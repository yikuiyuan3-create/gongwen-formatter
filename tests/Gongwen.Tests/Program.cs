using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Gongwen.Core;

var root = Path.GetFullPath(args.Length > 0 ? args[0] : ".");
var output = Path.Combine(root,"work","test-run-"+Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(output);
var sample = Path.Combine(root,"samples","01_通知_虚构示例.docx");
var formatter = new DocxFormatter();
var template = BuiltInTemplates.All[0];
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
var results = new List<object>();
int passed=0, failed=0;

void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
void Test(string name, Action test) {
 if(args.Contains("--v02-only")&&!name.StartsWith("v02_",StringComparison.Ordinal))return;
 if(args.Contains("--v03-only")&&!name.StartsWith("v03_",StringComparison.Ordinal))return;
 try { test(); passed++; results.Add(new {name,status="passed"}); Console.WriteLine("PASS "+name); }
 catch(Exception e) { failed++; results.Add(new{name,status="failed",errorType=e.GetType().Name,message=e.Message}); Console.WriteLine("FAIL "+name+": "+e.Message); }
}
void Reject(Action action) {
 try { action(); } catch(DocumentProcessingException) { return; }
 throw new Exception("Expected a safe document-processing rejection.");
}
string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
Dictionary<string,byte[]> Parts(string path) {
 using var z=ZipFile.OpenRead(path);
 return z.Entries.ToDictionary(e=>e.FullName,e=>{using var stream=e.Open();using var b=new MemoryStream();stream.CopyTo(b);return b.ToArray();});
}
XDocument Xml(string path) => XDocument.Parse(Encoding.UTF8.GetString(Parts(path)["word/document.xml"]));
string NewPath(string stem) => Path.Combine(output,stem+".docx");
IReadOnlyDictionary<int,ParagraphRole> Roles(DocumentAnalysis a) => a.Paragraphs.ToDictionary(p=>p.Index,p=>p.SuggestedRole);
FormatResult Run(string input,string target,bool acknowledge=true,FormatTemplate? t=null,IReadOnlyDictionary<int,ParagraphRole>? overrideRoles=null) {
 var a=formatter.Analyze(input);
 return formatter.Format(new(a,target,t??template,overrideRoles??Roles(a),acknowledge));
}
string Mutation(string name, Action<Dictionary<string,byte[]>> mutate) {
 var parts=Parts(sample);mutate(parts);var path=NewPath(name);
 using var z=ZipFile.Open(path,ZipArchiveMode.Create);
 foreach(var kv in parts) {var entry=z.CreateEntry(kv.Key);using var s=entry.Open();s.Write(kv.Value);}
 return path;
}

Test("read_fictional_document",()=>{
 var a=formatter.Analyze(sample);
 Check(a.CanFormat,"Basic DOCX must be eligible.");
 Check(a.Paragraphs.Count==11,"Expected 11 top-level paragraphs.");
 Check(a.Paragraphs[0].SuggestedRole==ParagraphRole.Title,"Expected a title suggestion.");
 Check(a.Paragraphs.Any(p=>p.SuggestedRole==ParagraphRole.Heading1),"Missing heading suggestion.");
});
var formatted=NewPath("formatted");
Test("format_and_preserve_original",()=>{
 var before=Hash(sample);var r=Run(sample,formatted);
 Check(Hash(sample)==before,"Source changed.");
 Check(File.Exists(formatted)&&r.FormattedParagraphCount>0,"No formatted output.");
 Check(r.OutputSha256.Equals(Hash(formatted),StringComparison.OrdinalIgnoreCase),"Output hash mismatch.");
});
Test("all_other_parts_byte_preserved",()=>{
 var a=Parts(sample);var b=Parts(formatted);
 Check(a.Keys.Order().SequenceEqual(b.Keys.Order()),"Package part names changed.");
 foreach(var key in a.Keys.Where(k=>k!="word/document.xml"))Check(a[key].SequenceEqual(b[key]),"Unrelated part changed: "+key);
});
Test("all_visible_text_preserved",()=>{
 var a=Xml(sample).Descendants(w+"t").Select(e=>e.Value);
 var b=Xml(formatted).Descendants(w+"t").Select(e=>e.Value);
 Check(a.SequenceEqual(b),"Visible text changed.");
});
Test("a4_and_margins_from_requirements",()=>{
 var d=Xml(formatted);var section=d.Descendants(w+"sectPr").Single();
 var size=section.Element(w+"pgSz")!;var margins=section.Element(w+"pgMar")!;
 Check((int?)size.Attribute(w+"w")==11906&&(int?)size.Attribute(w+"h")==16838,"Wrong A4 size.");
 foreach(var pair in new[]{("top",2098),("bottom",1984),("left",1587),("right",1474)})
  Check(Math.Abs(((int?)margins.Attribute(w+pair.Item1)??0)-pair.Item2)<=1,"Wrong margin "+pair.Item1);
 Check(section.Element(w+"docGrid")!=null,"Missing document grid.");
});
Test("body_16_point_and_two_character_indent",()=>{
 var p=Xml(formatted).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 Check((string?)p.Element(w+"pPr")?.Element(w+"ind")?.Attribute(w+"firstLineChars")=="200","Missing two-character indent.");
 Check(p.Descendants(w+"sz").Any(e=>(string?)e.Attribute(w+"val")=="32"),"Wrong body point size.");
});
Test("title_22_point_and_centered",()=>{
 var p=Xml(formatted).Root!.Element(w+"body")!.Elements(w+"p").First();
 Check((string?)p.Element(w+"pPr")?.Element(w+"jc")?.Attribute(w+"val")=="center","Title not centered.");
 Check(p.Descendants(w+"sz").Any(e=>(string?)e.Attribute(w+"val")=="44"),"Wrong title point size.");
});
Test("preserve_role_does_not_change_paragraph",()=>{
 var a=formatter.Analyze(sample);var roles=Roles(a).ToDictionary();roles[2]=ParagraphRole.Preserve;
 var target=NewPath("preserve");formatter.Format(new(a,target,template,roles,true));
 var p1=Xml(sample).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 var p2=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 Check(XNode.DeepEquals(p1,p2),"Preserved paragraph changed.");
});
Test("custom_font_names_are_explicit",()=>{
 var t=template with {BodyFont="测试仿宋字体",LatinFont="Test Latin"};
 var target=NewPath("custom-fonts");Run(sample,target,true,t);
 var p=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 Check(p.Descendants(w+"rFonts").Any(e=>(string?)e.Attribute(w+"eastAsia")=="测试仿宋字体"&&(string?)e.Attribute(w+"ascii")=="Test Latin"),"Custom fonts missing.");
});
Test("idempotent_formatting",()=>{
 var target=NewPath("second-pass");Run(formatted,target);
 Check(XNode.DeepEquals(Xml(formatted),Xml(target)),"Second pass changed the document XML.");
});
Test("same_source_destination_rejected",()=>Reject(()=>Run(sample,sample)));
Test("existing_output_rejected",()=>Reject(()=>Run(sample,formatted)));
Test("source_changed_after_analysis_rejected",()=>{
 var changed=NewPath("changed");File.Copy(sample,changed);var a=formatter.Analyze(changed);
 using(var s=new FileStream(changed,FileMode.Append))s.WriteByte(1);
 Reject(()=>formatter.Format(new(a,NewPath("stale-output"),template,Roles(a),true)));
});
Test("table_retained_and_explicit_acknowledgement",()=>{
 var input=Path.Combine(root,"samples","04_含表格_需要人工复核.docx");var a=formatter.Analyze(input);
 Check(a.TableCount==1&&a.Warnings.Count>0&&a.RequiresStructureReview,"Table must require structural review.");
 Reject(()=>formatter.Format(new(a,NewPath("table-not-acknowledged"),template,Roles(a),false)));
 var target=NewPath("table-acknowledged");Run(input,target);
 Check(XNode.DeepEquals(Xml(input).Descendants(w+"tbl").Single(),Xml(target).Descendants(w+"tbl").Single()),"Table changed.");
});
foreach(var fixture in new[]{"multisection"})Test(fixture+"_safe_block",()=>{
 var input=Path.Combine(root,"work","fixtures",fixture+".docx");var a=formatter.Analyze(input);
 Check(!a.CanFormat&&!string.IsNullOrWhiteSpace(a.BlockingReason),"Unsupported input not blocked.");
 Reject(()=>formatter.Format(new(a,NewPath(fixture+"-output"),template,Roles(a),true)));
});
Test("revisions_require_exact_paragraph_preservation",()=>{
 var input=Path.Combine(root,"work","fixtures","revisions.docx");var a=formatter.Analyze(input);
 Check(a.CanFormat&&a.Warnings.Count>0,"Revisions must be flagged for review.");
 Check(a.Paragraphs[2].SuggestedRole==ParagraphRole.Preserve,"Revised paragraph must require preservation.");
 var unsafeRoles=Roles(a).ToDictionary();unsafeRoles[2]=ParagraphRole.Body;
 Reject(()=>formatter.Format(new(a,NewPath("revision-unsafe"),template,unsafeRoles,true)));
 var target=NewPath("revision-preserved");Run(input,target);
 var before=Xml(input).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 var after=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 Check(XNode.DeepEquals(before,after),"Revision structure changed.");
});
Test("formatted_text_black_preserves_emphasis",()=>{
 var input=Mutation("colored",parts=>{
  var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
  var rp=d.Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2).Element(w+"r")!.Element(w+"rPr")!;
  rp.Add(new XElement(w+"b"),new XElement(w+"color",new XAttribute(w+"val","FF0000"),new XAttribute(w+"themeColor","accent1")));
  parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
 });
 var target=NewPath("black-output");Run(input,target);
 var p=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 Check(p.Descendants(w+"b").Any(),"Bold emphasis lost.");
 Check(p.Descendants(w+"color").All(c=>(string?)c.Attribute(w+"val")=="000000"&&c.Attribute(w+"themeColor")==null),"Color normalization failed.");
});
Test("grid_uses_default_font_basis",()=>{
 var input=Mutation("normal-11pt",parts=>{
  var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/styles.xml"]));
  d.Root!.Element(w+"style")!.Add(new XElement(w+"rPr",new XElement(w+"sz",new XAttribute(w+"val","22"))));
  parts["word/styles.xml"]=Encoding.UTF8.GetBytes(d.ToString());
 });
 var target=NewPath("grid-11pt");Run(input,target);
 var section=Xml(target).Descendants(w+"sectPr").Single();
 var margin=section.Element(w+"pgMar")!;
 double writableWidth=11906-(int)margin.Attribute(w+"left")!-(int)margin.Attribute(w+"right")!;
 int expected=(int)Math.Round((writableWidth/28-220)/20*4096);
 int actual=(int)section.Element(w+"docGrid")!.Attribute(w+"charSpace")!;
 Check(Math.Abs(actual-expected)<=1,"Grid not calculated against 11-point default font.");
});
string FieldFixture(string name,bool header,string[] fragments,bool simple=false){
 return Mutation(name,parts=>{
  XElement p=new(w+"p");
  if(simple)p.Add(new XElement(w+"fldSimple",new XAttribute(w+"instr",string.Concat(fragments)),new XElement(w+"r",new XElement(w+"t","虚构字段"))));
  else{
   p.Add(new XElement(w+"r",new XElement(w+"fldChar",new XAttribute(w+"fldCharType","begin"))));
   foreach(var fragment in fragments)p.Add(new XElement(w+"r",new XElement(w+"instrText",fragment)));
   p.Add(new XElement(w+"r",new XElement(w+"fldChar",new XAttribute(w+"fldCharType","end"))));
  }
  if(header){
   parts["word/header1.xml"]=Encoding.UTF8.GetBytes(new XDocument(new XElement(w+"hdr",p)).ToString());
   XNamespace ct="http://schemas.openxmlformats.org/package/2006/content-types";
   var types=XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));
   types.Root!.Add(new XElement(ct+"Override",new XAttribute("PartName","/word/header1.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
   parts["[Content_Types].xml"]=Encoding.UTF8.GetBytes(types.ToString());
  }else{
   var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
   d.Root!.Element(w+"body")!.AddFirst(p);
   parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
  }
 });
}
void BlockDangerousField(string input){
 var a=formatter.Analyze(input);
 Check(!a.CanFormat&&a.BlockingReason?.Contains("DDE")==true,"DDE field must block formatting.");
 Reject(()=>formatter.Format(new(a,NewPath("blocked-field-"+Guid.NewGuid().ToString("N")),template,Roles(a),true)));
}
Test("split_dde_instruction_rejected",()=>BlockDangerousField(FieldFixture("dde-split",false,new[]{" D","DE","AUTO fake fake fake "})));
Test("header_dde_instruction_rejected",()=>BlockDangerousField(FieldFixture("dde-header",true,new[]{" DDE fake fake fake "})));
Test("header_simple_dde_instruction_rejected",()=>BlockDangerousField(FieldFixture("dde-header-simple",true,new[]{" DDEAUTO fake fake fake "},true)));
Test("nested_field_dde_command_rejected",()=>{
 var input=Mutation("dde-nested",parts=>{
  var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
  XElement Marker(string type)=>new(w+"r",new XElement(w+"fldChar",new XAttribute(w+"fldCharType",type)));
  XElement Instruction(string text)=>new(w+"r",new XElement(w+"instrText",text));
  var p=new XElement(w+"p",Marker("begin"),Instruction(" D"),Marker("begin"),Instruction(" QUOTE fake "),Marker("separate"),new XElement(w+"r",new XElement(w+"t","DEAUTO")),Marker("end"),Instruction(" fake fake fake "),Marker("end"));
  d.Root!.Element(w+"body")!.AddFirst(p);
  parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
 });BlockDangerousField(input);
});
Test("ordinary_fields_preserved_without_execution",()=>{
 var input=FieldFixture("ordinary-field",false,new[]{" PAGE "});var a=formatter.Analyze(input);
 Check(a.Paragraphs[0].SuggestedRole==ParagraphRole.Preserve,"Field paragraph must be preserved.");
 var target=NewPath("ordinary-field-output");Run(input,target);
 Check(XNode.DeepEquals(Xml(input).Descendants(w+"p").First(),Xml(target).Descendants(w+"p").First()),"Field structure changed.");
});
Test("wrong_extension_rejected",()=>{
 var wrong=Path.Combine(output,"sample.doc");File.Copy(sample,wrong);Reject(()=>formatter.Analyze(wrong));
});
Test("invalid_zip_rejected",()=>{
 var bad=NewPath("not-zip");File.WriteAllText(bad,"not a document");Reject(()=>formatter.Analyze(bad));
});
Test("xml_dtd_rejected",()=>{
 var input=Mutation("dtd",parts=>parts["word/document.xml"]=Encoding.UTF8.GetBytes("<!DOCTYPE foo [<!ENTITY x 'unsafe'>]><foo>&x;</foo>"));
 Reject(()=>formatter.Analyze(input));
});
Test("zip_traversal_rejected",()=>{
 var input=Mutation("traversal",parts=>parts["../evil.txt"]=Encoding.UTF8.GetBytes("unsafe"));Reject(()=>formatter.Analyze(input));
});
Test("duplicate_parts_rejected",()=>{
 var input=NewPath("duplicate");File.Copy(sample,input);
 using(var z=ZipFile.Open(input,ZipArchiveMode.Update)){using var s=z.CreateEntry("word/document.xml").Open();s.Write(Parts(sample)["word/document.xml"]);}
 Reject(()=>formatter.Analyze(input));
});
Test("macro_payload_rejected",()=>{
 var input=Mutation("macro",parts=>parts["word/vbaProject.bin"]=new byte[]{1,2,3});Reject(()=>formatter.Analyze(input));
});
Test("template_invalid_page_bounds_rejected",()=>Reject(()=>Run(sample,NewPath("invalid-page"),true,template with{LeftMarginCm=20,RightMarginCm=20})));
Test("invalid_role_index_rejected",()=>{
 var a=formatter.Analyze(sample);var roles=Roles(a).ToDictionary();roles[99999]=ParagraphRole.Body;
 Reject(()=>formatter.Format(new(a,NewPath("bad-index"),template,roles,true)));
});
Test("v02_confident_plain_document_auto_roles",()=>{
 var a=formatter.Analyze(sample);
 Check(a.SuggestedTemplateId=="notice-v1","Clear notice title should auto-select notice.");
 Check(!a.RequiresStructureReview,"Plain document should not require structural acknowledgement.");
 Check(a.Paragraphs.Where(p=>p.SuggestedRole!=ParagraphRole.Preserve).All(p=>!p.NeedsReview),"Plain fictional notice should not require per-paragraph checkboxes.");
 Check(a.Paragraphs[1].SuggestedRole==ParagraphRole.Recipient,"Recipient missed.");
 Check(a.Paragraphs[8].SuggestedRole==ParagraphRole.AttachmentNote,"Attachment note missed.");
 Check(a.Paragraphs[9].SuggestedRole==ParagraphRole.Signature&&a.Paragraphs[10].SuggestedRole==ParagraphRole.Date,"Signature or date missed.");
 Check(a.Paragraphs[2].CanEdit,"Simple body should be editable.");
});
Test("v02_title_selects_request_or_report",()=>{
 foreach(var pair in new[]{("02_请示_虚构示例.docx","request-v1"),("03_报告_虚构示例.docx","report-v1")})
  Check(formatter.Analyze(Path.Combine(root,"samples",pair.Item1)).SuggestedTemplateId==pair.Item2,"Document type not derived from title.");
});
Test("v02_heading_style_inheritance",()=>{
 var input=Mutation("styled-heading",parts=>{
  var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));var p=d.Root!.Element(w+"body")!.Elements(w+"p").ElementAt(4);
  p.Descendants(w+"t").First().Value="整理范围";p.Element(w+"pPr")!.AddFirst(new XElement(w+"pStyle",new XAttribute(w+"val","CustomHeading")));
  parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
  var styles=XDocument.Parse(Encoding.UTF8.GetString(parts["word/styles.xml"]));
  styles.Root!.Add(new XElement(w+"style",new XAttribute(w+"type","paragraph"),new XAttribute(w+"styleId","Heading2"),new XElement(w+"pPr",new XElement(w+"outlineLvl",new XAttribute(w+"val","1")))),new XElement(w+"style",new XAttribute(w+"type","paragraph"),new XAttribute(w+"styleId","CustomHeading"),new XElement(w+"basedOn",new XAttribute(w+"val","Heading2"))));
  parts["word/styles.xml"]=Encoding.UTF8.GetBytes(styles.ToString());
 });
 Check(formatter.Analyze(input).Paragraphs[4].SuggestedRole==ParagraphRole.Heading2,"Inherited heading level missed.");
});
Test("v02_paragraph_edits_are_explicit_and_other_content_retained",()=>{
 var a=formatter.Analyze(sample);var target=NewPath("edited-body");var before=Hash(sample);
 const string edited="这是用户明确修改后的虚构正文，预算为4321.00元，编号EDIT-TEST。";
 formatter.Format(new(a,target,template,Roles(a),true,new Dictionary<int,string>{{2,edited}}));
 Check(Hash(sample)==before,"Editing changed source.");
 var original=Xml(sample).Root!.Element(w+"body")!.Elements(w+"p").Select(p=>string.Concat(p.Descendants(w+"t").Select(t=>t.Value))).ToArray();
 var actual=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").Select(p=>string.Concat(p.Descendants(w+"t").Select(t=>t.Value))).ToArray();
 Check(actual.Length==original.Length&&actual[2]==edited,"Explicit edit not exported correctly.");
 Check(original.Where((_,i)=>i!=2).SequenceEqual(actual.Where((_,i)=>i!=2)),"Unedited text changed.");
 foreach(var pair in Parts(sample).Where(kv=>kv.Key!="word/document.xml"))Check(pair.Value.SequenceEqual(Parts(target)[pair.Key]),"Other part changed by edit.");
});
Test("v02_edit_preserves_spaces_and_xml_characters",()=>{
 var a=formatter.Analyze(sample);var target=NewPath("edited-spaces");const string text="  测试 <XML> & 数字 100.00  ";
 formatter.Format(new(a,target,template,Roles(a),true,new Dictionary<int,string>{{2,text}}));
 var p=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 Check(string.Concat(p.Descendants(w+"t").Select(t=>t.Value))==text,"Edited whitespace or literal XML chars lost.");
 Check(p.Descendants(w+"t").Any(t=>(string?)t.Attribute(XNamespace.Xml+"space")=="preserve"),"Significant spaces must be marked preserve.");
});
Test("v02_unsafe_or_multiline_edits_rejected",()=>{
 var a=formatter.Analyze(sample);
 foreach(var value in new[]{"第一段\n第二段","禁止\0控制字符",new string('a',100001)})
  Reject(()=>formatter.Format(new(a,NewPath("invalid-edit-"+Guid.NewGuid().ToString("N")),template,Roles(a),true,new Dictionary<int,string>{{2,value}})));
 Reject(()=>formatter.Format(new(a,NewPath("invalid-edit-index"),template,Roles(a),true,new Dictionary<int,string>{{9999,"test"}})));
});
Test("v02_revision_edit_rejected",()=>{
 var a=formatter.Analyze(Path.Combine(root,"work","fixtures","revisions.docx"));
 Check(!a.Paragraphs[2].CanEdit,"Revised paragraphs must not be editable.");
 Reject(()=>formatter.Format(new(a,NewPath("revision-edit"),template,Roles(a),true,new Dictionary<int,string>{{2,"不可替换修订"}})));
});
Test("v02_bookmark_and_hyperlink_edits_rejected",()=>{
 foreach(var kind in new[]{"bookmark","hyperlink"}){
  var input=Mutation("special-edit-"+kind,parts=>{
   var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));var p=d.Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
   if(kind=="bookmark"){p.AddFirst(new XElement(w+"bookmarkStart",new XAttribute(w+"id","1"),new XAttribute(w+"name","TestBookmark")));p.Add(new XElement(w+"bookmarkEnd",new XAttribute(w+"id","1")));}
   else{var r=p.Element(w+"r")!;r.Remove();p.Add(new XElement(w+"hyperlink",new XAttribute(w+"anchor","test"),r));}
   parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
  });var a=formatter.Analyze(input);
  Check(!a.Paragraphs[2].CanEdit,"Special structure must not be editable: "+kind);
  Reject(()=>formatter.Format(new(a,NewPath("blocked-special-"+kind),template,Roles(a),true,new Dictionary<int,string>{{2,"替换会损害引用"}})));
 }
});
Test("v02_saved_document_with_writer_handle_can_be_formatted",()=>{
 var input=NewPath("wps-like-writer");File.Copy(sample,input);
 using var officeHandle=new FileStream(input,FileMode.Open,FileAccess.ReadWrite,FileShare.ReadWrite|FileShare.Delete);
 var a=formatter.Analyze(input);var target=NewPath("wps-like-output");
 formatter.Format(new(a,target,template,Roles(a),true));
 Check(File.Exists(target),"Concurrent office-style writer handle must not prevent saved-document import.");
});
string NumberedFixture(string name,bool inherited=false,bool protectedCopy=false) => Mutation(name,parts=>{
 var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));
 var p=d.Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 var numPr=new XElement(w+"numPr",new XElement(w+"ilvl",new XAttribute(w+"val","0")),new XElement(w+"numId",new XAttribute(w+"val","1")));
 var indent=new XElement(w+"ind",new XAttribute(w+"left","720"),new XAttribute(w+"hanging","360"));
 if(inherited){
  p.Element(w+"pPr")!.AddFirst(new XElement(w+"pStyle",new XAttribute(w+"val","ListBody")));
  var styles=XDocument.Parse(Encoding.UTF8.GetString(parts["word/styles.xml"]));
  styles.Root!.Add(new XElement(w+"style",new XAttribute(w+"type","paragraph"),new XAttribute(w+"styleId","ListBody"),new XElement(w+"basedOn",new XAttribute(w+"val","Normal")),new XElement(w+"pPr",numPr,indent)));
  parts["word/styles.xml"]=Encoding.UTF8.GetBytes(styles.ToString());
 }else p.Element(w+"pPr")!.Add(numPr,indent);
 parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
 parts["word/numbering.xml"]=Encoding.UTF8.GetBytes(new XDocument(new XElement(w+"numbering",new XElement(w+"abstractNum",new XAttribute(w+"abstractNumId","0"),new XElement(w+"lvl",new XAttribute(w+"ilvl","0"),new XElement(w+"start",new XAttribute(w+"val","1")),new XElement(w+"numFmt",new XAttribute(w+"val","decimal")),new XElement(w+"lvlText",new XAttribute(w+"val","%1.")),new XElement(w+"pPr",new XElement(indent)))),new XElement(w+"num",new XAttribute(w+"numId","1"),new XElement(w+"abstractNumId",new XAttribute(w+"val","0"))))).ToString());
 XNamespace ct="http://schemas.openxmlformats.org/package/2006/content-types";
 var types=XDocument.Parse(Encoding.UTF8.GetString(parts["[Content_Types].xml"]));types.Root!.Add(new XElement(ct+"Override",new XAttribute("PartName","/word/numbering.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")));parts["[Content_Types].xml"]=Encoding.UTF8.GetBytes(types.ToString());
 XNamespace rel="http://schemas.openxmlformats.org/package/2006/relationships";
 var relations=XDocument.Parse(Encoding.UTF8.GetString(parts["word/_rels/document.xml.rels"]));relations.Root!.Add(new XElement(rel+"Relationship",new XAttribute("Id","rIdNumbering"),new XAttribute("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),new XAttribute("Target","numbering.xml")));parts["word/_rels/document.xml.rels"]=Encoding.UTF8.GetBytes(relations.ToString());
 if(protectedCopy)parts["word/settings.xml"]=Encoding.UTF8.GetBytes(new XDocument(new XElement(w+"settings",new XElement(w+"trackRevisions"),new XElement(w+"documentProtection",new XAttribute(w+"edit","readOnly"),new XAttribute(w+"enforcement","1")),new XElement(w+"writeProtection",new XAttribute(w+"recommended","1")),new XElement(w+"zoom",new XAttribute(w+"percent","100")))).ToString());
});
Test("v03_edit_protection_detected_without_global_export_block",()=>{
 var a=formatter.Analyze(Path.Combine(root,"work","fixtures","protected.docx"));
 Check(a.HasEditProtection&&a.CanFormat,"Readable editing protection should allow an explicitly editable output copy.");
 Check(a.Paragraphs[2].CanEdit,"Simple protected document text should be editable in a new copy.");
});
Test("v03_protected_copy_requires_explicit_policy",()=>{
 var a=formatter.Analyze(Path.Combine(root,"work","fixtures","protected.docx"));
 Reject(()=>formatter.Format(new(a,NewPath("protected-without-policy"),template,Roles(a),true)));
 Reject(()=>formatter.Format(new(a with{HasEditProtection=false},NewPath("protected-forged-policy"),template,Roles(a),true)));
 Check(!File.Exists(NewPath("protected-without-policy")),"Rejected request left an output.");
});
Test("v03_protected_copy_edits_text_and_preserves_original",()=>{
 var input=Path.Combine(root,"work","fixtures","protected.docx");var a=formatter.Analyze(input);var before=Hash(input);var target=NewPath("protected-edited");
 var r=formatter.Format(new(a,target,template,Roles(a),true,new Dictionary<int,string>{{2,"这是在新副本中修改的虚构正文。"}},true));
 Check(r.EditProtectionRemoved&&Hash(input)==before,"Protection copy did not preserve source or record policy.");
 Check(!formatter.Analyze(target).HasEditProtection,"Output remains edit-protected.");
 Check(string.Concat(Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2).Descendants(w+"t").Select(t=>t.Value))=="这是在新副本中修改的虚构正文。","Protected copy edit lost.");
 foreach(var pair in Parts(input).Where(kv=>kv.Key!="word/document.xml"&&kv.Key!="word/settings.xml"))Check(pair.Value.SequenceEqual(Parts(target)[pair.Key]),"Unrelated protected package part changed.");
});
Test("v03_remove_only_edit_restrictions_keep_other_settings",()=>{
 var input=NumberedFixture("numbered-protected-input",protectedCopy:true);var a=formatter.Analyze(input);var target=NewPath("numbered-protected-output");
 formatter.Format(new(a,target,template,Roles(a),true,new Dictionary<int,string>{{2,"这是修改后的虚构编号正文，输出副本可编辑。"}},true));
 var original=XDocument.Parse(Encoding.UTF8.GetString(Parts(input)["word/settings.xml"]));original.Root!.Elements().Where(e=>e.Name==w+"documentProtection"||e.Name==w+"writeProtection").Remove();
 Check(XNode.DeepEquals(original,XDocument.Parse(Encoding.UTF8.GetString(Parts(target)["word/settings.xml"]))),"Settings beyond edit restrictions changed.");
 Check(Parts(input)["word/numbering.xml"].SequenceEqual(Parts(target)["word/numbering.xml"]),"Numbering definitions changed in editable copy.");
 Check(string.Concat(Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2).Descendants(w+"t").Select(t=>t.Value))=="这是修改后的虚构编号正文，输出副本可编辑。","Combined protected/numbered edit lost.");
});
Test("v03_section_form_protection_removed_only_in_copy",()=>{
 var input=Mutation("section-protected",parts=>{var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));d.Descendants(w+"sectPr").Single().Add(new XElement(w+"formProt",new XAttribute(w+"val","1")));parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());});
 var a=formatter.Analyze(input);var before=Hash(input);Check(a.HasEditProtection&&a.CanFormat,"Section edit restriction globally blocked.");
 var target=NewPath("section-unprotected");var r=formatter.Format(new(a,target,template,Roles(a),true,null,true));
 Check(r.EditProtectionRemoved&&Hash(input)==before&&!Xml(target).Descendants(w+"formProt").Any(),"Section restriction not removed safely.");
 Check(Parts(input)["word/settings.xml"].SequenceEqual(Parts(target)["word/settings.xml"]),"Unprotected settings should remain byte-identical.");
});
foreach(var inherited in new[]{false,true})Test("v03_"+(inherited?"inherited":"direct")+"_numbering_edit_preserves_sequence_and_indent",()=>{
 var input=NumberedFixture(inherited?"inherited-numbered":"direct-numbered",inherited);var a=formatter.Analyze(input);var p=a.Paragraphs[2];
 Check(p.HasAutomaticNumbering&&p.CanEdit&&p.SuggestedRole!=ParagraphRole.Preserve,"Simple automatic list falsely marked protected.");
 var target=NewPath(inherited?"inherited-numbered-edited":"direct-numbered-edited");
 var r=formatter.Format(new(a,target,template,Roles(a),true,new Dictionary<int,string>{{2,"修改后的虚构编号正文，保留原编号。"}}));
 Check(!r.EditProtectionRemoved,"Normal list falsely claims protection removal.");
 var original=Xml(input).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);var actual=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 Check(XNode.DeepEquals(original.Element(w+"pPr")?.Element(w+"numPr"),actual.Element(w+"pPr")?.Element(w+"numPr")),"Numbering reference changed.");
 Check(XNode.DeepEquals(original.Element(w+"pPr")?.Element(w+"ind"),actual.Element(w+"pPr")?.Element(w+"ind")),"Direct or inherited list indentation changed.");
 Check(string.Concat(actual.Descendants(w+"t").Select(t=>t.Value))=="修改后的虚构编号正文，保留原编号。","Numbered paragraph text edit lost.");
 foreach(var pair in Parts(input).Where(kv=>kv.Key!="word/document.xml"))Check(pair.Value.SequenceEqual(Parts(target)[pair.Key]),"List-related or unrelated package part changed.");
});
Test("v03_protection_copy_does_not_unlock_revision_structure",()=>{
 var input=Mutation("protected-revision",parts=>{parts["word/document.xml"]=Parts(Path.Combine(root,"work","fixtures","revisions.docx"))["word/document.xml"];parts["word/settings.xml"]=Parts(Path.Combine(root,"work","fixtures","protected.docx"))["word/settings.xml"];});
 var a=formatter.Analyze(input);Check(!a.Paragraphs[2].CanEdit,"Revision paragraph unlocked by document copy policy.");
 Reject(()=>formatter.Format(new(a,NewPath("unsafe-protected-revision"),template,Roles(a),true,new Dictionary<int,string>{{2,"禁止覆盖修订结构"}},true)));
 var target=NewPath("protected-revision-preserved");formatter.Format(new(a,target,template,Roles(a),true,null,true));
 Check(XNode.DeepEquals(Xml(input).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2),Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2)),"Revision lost while removing edit restrictions.");
});
Test("v03_protection_copy_keeps_unsafe_field_block",()=>{
 var input=FieldFixture("protected-dde",false,new[]{" DDEAUTO fake fake fake "});
 var path=NewPath("protected-dde-settings");var parts=Parts(input);parts["word/settings.xml"]=Parts(Path.Combine(root,"work","fixtures","protected.docx"))["word/settings.xml"];
 using(var z=ZipFile.Open(path,ZipArchiveMode.Create)){foreach(var pair in parts){using var s=z.CreateEntry(pair.Key).Open();s.Write(pair.Value);}}
 var a=formatter.Analyze(path);Check(!a.CanFormat,"Protection copy must not bypass unsafe field block.");
 Reject(()=>formatter.Format(new(a,NewPath("unsafe-field-copy"),template,Roles(a),true,null,true)));
});
Test("v03_unprotected_copy_policy_does_not_change_settings",()=>{
 var a=formatter.Analyze(sample);var target=NewPath("normal-copy-policy");var r=formatter.Format(new(a,target,template,Roles(a),true,null,true));
 Check(!r.EditProtectionRemoved&&Parts(sample)["word/settings.xml"].SequenceEqual(Parts(target)["word/settings.xml"]),"Copy policy changed ordinary settings.");
});
Test("v03_nondefault_settings_part_name_edit_copy",()=>{
 var input=Mutation("alternate-settings",parts=>{
  parts["word/customSettings.xml"]=Parts(Path.Combine(root,"work","fixtures","protected.docx"))["word/settings.xml"];parts.Remove("word/settings.xml");
  foreach(var key in new[]{"[Content_Types].xml","word/_rels/document.xml.rels"})parts[key]=Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(parts[key]).Replace("settings.xml","customSettings.xml",StringComparison.Ordinal));
 });var a=formatter.Analyze(input);var before=Hash(input);Check(a.HasEditProtection&&a.CanFormat,"Alternative valid settings part not detected.");
 var target=NewPath("alternate-settings-output");var r=formatter.Format(new(a,target,template,Roles(a),true,null,true));
 Check(r.EditProtectionRemoved&&Hash(input)==before&&!formatter.Analyze(target).HasEditProtection,"Alternative settings restriction not removed from new copy.");
 Check(Parts(input).Keys.Order().SequenceEqual(Parts(target).Keys.Order()),"Alternative settings part name changed.");
 foreach(var pair in Parts(input).Where(kv=>kv.Key!="word/document.xml"&&kv.Key!="word/customSettings.xml"))Check(pair.Value.SequenceEqual(Parts(target)[pair.Key]),"Unrelated part changed in alternative settings copy.");
});
Test("v03_numbered_heading_derived_from_inherited_outline",()=>{
 var path=NumberedFixture("numbered-heading-base",inherited:true);var parts=Parts(path);var styles=XDocument.Parse(Encoding.UTF8.GetString(parts["word/styles.xml"]));
 styles.Root!.Elements(w+"style").Single(s=>(string?)s.Attribute(w+"styleId")=="ListBody").Element(w+"pPr")!.Add(new XElement(w+"outlineLvl",new XAttribute(w+"val","0")));parts["word/styles.xml"]=Encoding.UTF8.GetBytes(styles.ToString());
 var input=NewPath("numbered-heading");using(var z=ZipFile.Open(input,ZipArchiveMode.Create)){foreach(var pair in parts){using var s=z.CreateEntry(pair.Key).Open();s.Write(pair.Value);}}
 var a=formatter.Analyze(input);Check(a.Paragraphs[2].SuggestedRole==ParagraphRole.Heading1&&a.Paragraphs[2].CanEdit,"Inherited numbered heading not recognized.");
 var target=NewPath("numbered-heading-output");Run(input,target);
 Check(Parts(input)["word/styles.xml"].SequenceEqual(Parts(target)["word/styles.xml"])&&Parts(input)["word/numbering.xml"].SequenceEqual(Parts(target)["word/numbering.xml"]),"Inherited numbered heading definitions changed.");
});
Test("v03_complex_numbering_remains_preserved",()=>{
 var path=NumberedFixture("complex-numbering-base");var parts=Parts(path);var d=Xml(path);var p=d.Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 p.Element(w+"pPr")!.Element(w+"numPr")!.Add(new XElement(w+"numberingChange",new XAttribute(w+"id","1"),new XAttribute(w+"author","虚构测试")));parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
 var input=NewPath("complex-numbering");using(var z=ZipFile.Open(input,ZipArchiveMode.Create)){foreach(var pair in parts){using var s=z.CreateEntry(pair.Key).Open();s.Write(pair.Value);}}
 var a=formatter.Analyze(input);Check(!a.Paragraphs[2].CanEdit&&a.Paragraphs[2].SuggestedRole==ParagraphRole.Preserve,"Complex numbering falsely editable.");
 var roles=Roles(a).ToDictionary();roles[2]=ParagraphRole.Body;Reject(()=>formatter.Format(new(a,NewPath("complex-numbering-unsafe"),template,roles,true)));
 var target=NewPath("complex-numbering-output");Run(input,target);
 Check(XNode.DeepEquals(Xml(input).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2),Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2)),"Complex numbering not preserved.");
});
Test("v03_numbered_paragraph_with_direct_tab_stops_is_editable",()=>{
 var path=NumberedFixture("numbered-tabs-base");var parts=Parts(path);var d=Xml(path);var p=d.Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 p.Element(w+"pPr")!.Add(new XElement(w+"tabs",new XElement(w+"tab",new XAttribute(w+"val","num"),new XAttribute(w+"pos","720"),new XAttribute(w+"leader","none"))));parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());
 var input=NewPath("numbered-tabs");using(var z=ZipFile.Open(input,ZipArchiveMode.Create)){foreach(var pair in parts){using var s=z.CreateEntry(pair.Key).Open();s.Write(pair.Value);}}
 var a=formatter.Analyze(input);Check(a.Paragraphs[2].CanEdit,"Ordinary numbered paragraph with direct tab stops falsely protected.");
 var target=NewPath("numbered-tabs-output");formatter.Format(new(a,target,template,Roles(a),true,new Dictionary<int,string>{{2,"编辑了保留编号制表位的虚构正文。"}}));
 var before=Xml(input).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);var after=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
 foreach(var name in new[]{"tabs","ind","numPr"})Check(XNode.DeepEquals(before.Element(w+"pPr")!.Element(w+name),after.Element(w+"pPr")!.Element(w+name)),"Numbered paragraph layout changed: "+name);
});
Test("v03_tab_content_and_invalid_tab_structure_edits_rejected",()=>{
 foreach(var kind in new[]{"run-tab","nested-tabs"}){
  var input=Mutation(kind,parts=>{var d=XDocument.Parse(Encoding.UTF8.GetString(parts["word/document.xml"]));var p=d.Root!.Element(w+"body")!.Elements(w+"p").ElementAt(2);
   if(kind=="run-tab")p.Element(w+"r")!.Add(new XElement(w+"tab"));
   else p.Element(w+"pPr")!.Add(new XElement(w+"tabs",new XElement(w+"tab",new XAttribute(w+"val","left"),new XAttribute(w+"pos","720"),new XElement(w+"unknown"))));
   parts["word/document.xml"]=Encoding.UTF8.GetBytes(d.ToString());});
  var a=formatter.Analyze(input);Check(!a.Paragraphs[2].CanEdit,"Unsafe tab content falsely editable: "+kind);
  Reject(()=>formatter.Format(new(a,NewPath("unsafe-tab-edit-"+kind),template,Roles(a),true,new Dictionary<int,string>{{2,"不能覆盖复杂制表内容"}})));
 }
});
Test("v03_shipping_protected_numbered_sample_copy",()=>{
 var input=Path.Combine(root,"samples","07_编辑保护与自动编号_虚构示例.docx");var a=formatter.Analyze(input);var before=Hash(input);
 Check(a.HasEditProtection&&a.Paragraphs.Count(p=>p.HasAutomaticNumbering)==2,"Shipping protected/numbered fixture missing its structures.");
 var target=NewPath("protected-numbered-shipping-output");var r=formatter.Format(new(a,target,template,Roles(a),true,new Dictionary<int,string>{{2,"这是修改后的虚构编号正文，输出副本可编辑。"}},true));
 Check(r.EditProtectionRemoved&&Hash(input)==before&&!formatter.Analyze(target).HasEditProtection,"Shipping edit-protected sample was not copied safely.");
 foreach(var index in new[]{2,7}){var original=Xml(input).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(index);var actual=Xml(target).Root!.Element(w+"body")!.Elements(w+"p").ElementAt(index);
  foreach(var name in new[]{"ind","numPr"})Check(XNode.DeepEquals(original.Element(w+"pPr")!.Element(w+name),actual.Element(w+"pPr")!.Element(w+name)),"Shipping sample numbering layout changed.");}
 Check(Parts(input)["word/numbering.xml"].SequenceEqual(Parts(target)["word/numbering.xml"]),"Shipping sample numbering definition changed.");
});
Test("no_uncommitted_temporary_outputs",()=>{
 Check(!Directory.EnumerateFiles(output).Any(p=>p.EndsWith(".tmp",StringComparison.OrdinalIgnoreCase)),"Leftover temporary output.");
});
var summary=new{passed,failed,total=passed+failed,environment=System.Runtime.InteropServices.RuntimeInformation.OSDescription+" core integration; WPS/Word host not exercised",outputDirectory=output,results};
Directory.CreateDirectory(Path.Combine(root,"evidence"));
File.WriteAllText(Path.Combine(root,"evidence","core-tests.json"),JsonSerializer.Serialize(summary,new JsonSerializerOptions{WriteIndented=true}));
Console.WriteLine($"RESULT {passed} passed, {failed} failed");
Environment.ExitCode=failed==0?0:1;
