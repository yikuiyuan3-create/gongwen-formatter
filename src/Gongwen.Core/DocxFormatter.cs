using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace Gongwen.Core;

/// <summary>本地 DOCX 文字稿排版。只修改主文档内已选择段落的直接格式和单节页面设置。</summary>
public sealed partial class DocxFormatter : IDocumentFormatter
{
    private const long MaximumInputBytes = 50L * 1024 * 1024;
    private const long MaximumExpandedBytes = 200L * 1024 * 1024;
    private const int MaximumEntries = 10_000;
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string MainContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string MainPart = "word/document.xml";

    public DocumentAnalysis Analyze(string sourcePath)
    {
        var package = ReadPackage(sourcePath);
        var analysis = AnalyzePackage(package);
        using var verifiedSource = OpenVerifiedSource(package);
        return analysis;
    }

    public FormatResult Format(FormatRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Analysis);
        ArgumentNullException.ThrowIfNull(request.Template);
        ArgumentNullException.ThrowIfNull(request.Roles);
        ValidateTemplate(request.Template);

        // Load one immutable snapshot and use it for analysis, mutation and verification.
        var package = ReadPackage(request.Analysis.SourcePath);
        if (!string.Equals(package.Sha256, request.Analysis.SourceSha256, StringComparison.OrdinalIgnoreCase))
            throw new DocumentProcessingException("原稿在分析后发生变化，请重新导入和分析。");
        var currentAnalysis = AnalyzePackage(package);
        if (!currentAnalysis.CanFormat)
            throw new DocumentProcessingException(currentAnalysis.BlockingReason ?? "当前文档不适合自动排版。");
        if (currentAnalysis.HasEditProtection && !request.AllowEditingProtectedCopy)
            throw new DocumentProcessingException("原稿有编辑限制，请启用“仅在新副本中移除编辑限制”；原稿不会被修改。加密文件仍需先在办公软件中合法解密。");
        if (currentAnalysis.Warnings.Count > 0 && !request.AcknowledgedWarnings)
            throw new DocumentProcessingException("请先确认文档检查提示，再执行排版。");

        var outputPath = ValidateOutputPath(package.SourcePath, request.OutputPath);
        var document = package.Document;
        var originalDocument = new XDocument(document);
        var body = document.Root!.Element(W + "body")!;
        var paragraphs = body.Elements(W + "p").ToList();
        foreach (var index in request.Roles.Keys)
            if (index < 0 || index >= paragraphs.Count)
                throw new DocumentProcessingException("段落选择已失效，请重新分析文档。");

        // Build an independently edited baseline. Only explicit plain-text edits
        // are allowed to differ from the source during subsequent verification.
        var edits = request.ParagraphEdits?.ToDictionary(pair => pair.Key, pair => pair.Value);
        ApplyParagraphEdits(document, edits, package);
        ApplyParagraphEdits(originalDocument, edits, package);
        var changedParts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var editProtectionRemoved = currentAnalysis.HasEditProtection && RemoveEditProtectionFromCopy(package, document, changedParts);

        var effectiveRoles = new ParagraphRole[paragraphs.Count];
        var formattedCount = 0;
        for (var index = 0; index < paragraphs.Count; index++)
        {
            var info = currentAnalysis.Paragraphs[index];
            var role = request.Roles.TryGetValue(index, out var selected)
                ? selected : info.NeedsReview ? ParagraphRole.Preserve : info.SuggestedRole;
            if (!Enum.IsDefined(role))
                throw new DocumentProcessingException("段落类型无效，请重新选择。");
            effectiveRoles[index] = role;
            if (role == ParagraphRole.Preserve)
                continue;
            if (HasRevision(paragraphs[index]))
                throw new DocumentProcessingException($"第 {index + 1} 段包含修订，必须选择“保留原格式”；请在办公软件中人工处理修订后重新导入。");
            if (!HasSafeNumberingProperties(paragraphs[index], package))
                throw new DocumentProcessingException($"第 {index + 1} 段的编号结构复杂或无效，必须选择“保留原格式”，避免改变编号布局。");
            if (HasProtectedStructure(paragraphs[index]))
                throw new DocumentProcessingException($"第 {index + 1} 段含字段、内容控件或特殊布局，必须选择“保留原格式”。");
            FormatParagraph(paragraphs[index], role, request.Template, HasAutomaticNumbering(paragraphs[index], package));
            formattedCount++;
        }
        FormatSection(body, request.Template, GetNormalFontSize(package) ?? request.Template.BodyFontPt);
        VerifyDocumentMutation(originalDocument, document, effectiveRoles, package);
        var changedMainPart = Serialize(document);
        changedParts[MainPart] = changedMainPart;

        var temporaryPath = Path.Combine(Path.GetDirectoryName(outputPath)!, $".gongwen-{Guid.NewGuid():N}.tmp");
        var moved = false;
        try
        {
            using (var output = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                using (var archive = new ZipArchive(output, ZipArchiveMode.Create, true, Encoding.UTF8))
                {
                    foreach (var part in package.Parts)
                    {
                        var entry = archive.CreateEntry(part.Name, CompressionLevel.Optimal);
                        entry.LastWriteTime = part.LastWriteTime;
                        entry.ExternalAttributes = part.ExternalAttributes;
                        using var destination = entry.Open();
                        destination.Write(changedParts.TryGetValue(part.Name, out var changedBytes) ? changedBytes : part.Bytes);
                    }
                }
                output.Flush(flushToDisk: true);
            }
            VerifyOutputPackage(temporaryPath, package, document, changedParts);
            var outputSha256 = HashFile(temporaryPath);
            // Recheck the saved source immediately before committing. Shared
            // access permits WPS / Word to remain open; source changes are rejected.
            using var finalSource = OpenVerifiedSource(package);
            // Atomic rename in the destination directory; no existing file is overwritten.
            File.Move(temporaryPath, outputPath, overwrite: false);
            moved = true;
            return new FormatResult(outputPath, package.Sha256, outputSha256, formattedCount, currentAnalysis.Warnings, editProtectionRemoved);
        }
        catch (DocumentProcessingException) { throw; }
        catch (IOException) { throw new DocumentProcessingException("无法保存新文档：请检查目标文件是否已存在、磁盘空间和目录写入权限。"); }
        catch (UnauthorizedAccessException) { throw new DocumentProcessingException("没有目标目录的写入权限，请选择可写目录。"); }
        finally
        {
            if (!moved)
            {
                try { File.Delete(temporaryPath); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }

    private static Package ReadPackage(string sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
            throw new DocumentProcessingException("请选择需要排版的 DOCX 文件。");
        string fullPath;
        try { fullPath = Path.GetFullPath(sourcePath); }
        catch (Exception error) when (error is ArgumentException or NotSupportedException or PathTooLongException)
        { throw new DocumentProcessingException("文档路径无效，请重新选择文件。"); }
        if (!string.Equals(Path.GetExtension(fullPath), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new DocumentProcessingException("首版仅支持 DOCX，请在 WPS 或 Word 中将原稿另存为 DOCX。");

        try
        {
            var bytes = ReadStableSource(fullPath);
            var sha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
            using var memory = new MemoryStream(bytes, writable: false);
            using var archive = new ZipArchive(memory, ZipArchiveMode.Read, false, Encoding.UTF8);
            if (archive.Entries.Count > MaximumEntries)
                throw new DocumentProcessingException("文档包含过多文件部件，不适合自动处理。");
            var parts = ReadParts(archive);
            var byName = parts.ToDictionary(part => part.Name, StringComparer.Ordinal);
            if (!byName.ContainsKey("[Content_Types].xml") || !byName.ContainsKey("_rels/.rels") || !byName.ContainsKey(MainPart))
                throw new DocumentProcessingException("文件缺少标准 DOCX 结构，请在 WPS 或 Word 中重新另存。");
            var xmlParts = new Dictionary<string, XDocument>(StringComparer.Ordinal);
            var xmlBytes = parts.Where(IsXmlPart).Sum(part => (long)part.Bytes.Length);
            if (xmlBytes > 50L * 1024 * 1024)
                throw new DocumentProcessingException("文档 XML 结构过大，不适合自动处理；请拆分文字稿后重试。");
            foreach (var part in parts.Where(IsXmlPart))
                xmlParts.Add(part.Name, ParseXml(part.Bytes));
            ValidateContentTypes(xmlParts["[Content_Types].xml"], parts);
            ValidateRelationships(xmlParts, byName);
            var document = xmlParts[MainPart];
            if (document.Root?.Name != W + "document" || document.Root.Elements(W + "body").Count() != 1)
                throw new DocumentProcessingException("首版仅支持标准 DOCX 文字稿；请在办公软件中另存为兼容格式后重试。");
            return new Package(fullPath, sha256, parts, byName, xmlParts, document);
        }
        catch (DocumentProcessingException) { throw; }
        catch (UnauthorizedAccessException) { throw new DocumentProcessingException("无法读取原稿，请检查文件读取权限。"); }
        catch (FileNotFoundException) { throw new DocumentProcessingException("未找到原稿，请重新选择文件。"); }
        catch (DirectoryNotFoundException) { throw new DocumentProcessingException("未找到原稿所在目录，请重新选择文件。"); }
        catch (Exception error) when (error is InvalidDataException or XmlException or IOException or ArgumentException or OverflowException)
        { throw new DocumentProcessingException("无法读取标准 DOCX：文件可能已加密、损坏、被占用或包含不安全的 XML。请在办公软件中检查后另存。"); }
    }

    private static List<Part> ReadParts(ZipArchive archive)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<Part>(archive.Entries.Count);
        long total = 0;
        foreach (var entry in archive.Entries)
        {
            ValidateEntryName(entry.FullName);
            if (!names.Add(entry.FullName))
                throw new DocumentProcessingException("文档包含重复的文件部件，已停止处理。");
            if (entry.Length < 0 || entry.Length > MaximumExpandedBytes || total > MaximumExpandedBytes - entry.Length)
                throw new DocumentProcessingException("文档解压后的部件总大小超过 200 MB，已停止处理。");
            using var source = entry.Open();
            using var buffer = new MemoryStream();
            var chunk = new byte[81920];
            long partSize = 0;
            int read;
            while ((read = source.Read(chunk, 0, chunk.Length)) != 0)
            {
                partSize += read;
                total += read;
                if (total > MaximumExpandedBytes || partSize > MaximumExpandedBytes)
                    throw new DocumentProcessingException("文档解压后的部件总大小超过 200 MB，已停止处理。");
                buffer.Write(chunk, 0, read);
            }
            if (partSize != entry.Length)
                throw new DocumentProcessingException("文档部件大小不一致，可能已经损坏。");
            result.Add(new Part(entry.FullName, buffer.ToArray(), entry.LastWriteTime, entry.ExternalAttributes));
        }
        return result;
    }

    private static void ValidateEntryName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.StartsWith('/') || name.Contains('\\') || name.Contains(':') || name.Any(char.IsControl))
            throw new DocumentProcessingException("文档包含不安全的文件部件路径，已停止处理。");
        var trimmed = name.EndsWith('/') ? name[..^1] : name;
        var segments = trimmed.Split('/');
        if (segments.Any(segment => segment is "" or "." or ".."))
            throw new DocumentProcessingException("文档包含不安全的文件部件路径，已停止处理。");
        var decoded = Uri.UnescapeDataString(name);
        if (decoded != name && (decoded.Contains('\\') || decoded.Contains(':') || decoded.Split('/').Any(segment => segment is "." or "..") || decoded.StartsWith('/')))
            throw new DocumentProcessingException("文档包含不安全的编码路径，已停止处理。");
    }

    private static bool IsXmlPart(Part part) =>
        part.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
        part.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
        part.Name.EndsWith(".vml", StringComparison.OrdinalIgnoreCase) ||
        part.Name.EndsWith(".svg", StringComparison.OrdinalIgnoreCase);

    private static XDocument ParseXml(byte[] bytes)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaximumExpandedBytes,
            MaxCharactersFromEntities = 0,
            IgnoreWhitespace = false,
            IgnoreComments = false
        };
        // Bound nesting before building an object tree, including parts that remain untouched.
        var nodeCount = 0;
        using (var stream = new MemoryStream(bytes, writable: false))
        using (var reader = XmlReader.Create(stream, settings))
            while (reader.Read())
            {
                if (reader.Depth > 128)
                    throw new DocumentProcessingException("文档 XML 嵌套过深，已停止处理。");
                if (++nodeCount > 500_000)
                    throw new DocumentProcessingException("文档 XML 节点过多，不适合自动处理。");
            }
        using var content = new MemoryStream(bytes, writable: false);
        using var xmlReader = XmlReader.Create(content, settings);
        return XDocument.Load(xmlReader, LoadOptions.PreserveWhitespace);
    }

    private static void ValidateContentTypes(XDocument contentTypeDocument, IReadOnlyList<Part> parts)
    {
        if (contentTypeDocument.Root?.Name != ContentTypes + "Types")
            throw new DocumentProcessingException("DOCX 内容类型结构无效。");
        var types = contentTypeDocument.Root.Elements().Select(element => (string?)element.Attribute("ContentType") ?? "").ToArray();
        if (types.Any(type => type.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase) ||
                type.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) ||
                type.Contains("vbaData", StringComparison.OrdinalIgnoreCase) ||
                type.Contains("activeX", StringComparison.OrdinalIgnoreCase)) ||
            parts.Any(part => part.Name.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) ||
                part.Name.Contains("vbaData", StringComparison.OrdinalIgnoreCase) ||
                part.Name.StartsWith("word/activeX/", StringComparison.OrdinalIgnoreCase)))
            throw new DocumentProcessingException("不支持含宏或 ActiveX 控件的文档，请另存为无宏 DOCX 后重试。");
        if (types.Any(type => type.Contains("digital-signature", StringComparison.OrdinalIgnoreCase)) ||
            parts.Any(part => part.Name.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)))
            throw new DocumentProcessingException("文档含数字签名，排版会使签名失效；请先由文件负责人处理签名。");
        var mainTypes = contentTypeDocument.Root.Elements(ContentTypes + "Override")
            .Where(element => (string?)element.Attribute("PartName") == "/" + MainPart).ToList();
        if (mainTypes.Count != 1 || (string?)mainTypes[0].Attribute("ContentType") != MainContentType)
            throw new DocumentProcessingException("文件不是受支持的无宏 DOCX 文字稿。");
    }

    private static void ValidateRelationships(IReadOnlyDictionary<string, XDocument> xmlParts, IReadOnlyDictionary<string, Part> parts)
    {
        foreach (var pair in xmlParts.Where(pair => pair.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            if (pair.Value.Root?.Name != PackageRelationships + "Relationships")
                throw new DocumentProcessingException("DOCX 关系结构无效。");
            var identifiers = new HashSet<string>(StringComparer.Ordinal);
            foreach (var relationship in pair.Value.Root.Elements(PackageRelationships + "Relationship"))
            {
                var id = (string?)relationship.Attribute("Id");
                var type = (string?)relationship.Attribute("Type") ?? "";
                var target = (string?)relationship.Attribute("Target");
                if (string.IsNullOrEmpty(id) || !identifiers.Add(id) || string.IsNullOrWhiteSpace(target) || string.IsNullOrWhiteSpace(type))
                    throw new DocumentProcessingException("文档关系记录不完整或重复。");
                if (type.Contains("digital-signature", StringComparison.OrdinalIgnoreCase))
                    throw new DocumentProcessingException("文档含数字签名，排版会使签名失效；请先由文件负责人处理签名。");
                if (type.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) || type.Contains("activeX", StringComparison.OrdinalIgnoreCase))
                    throw new DocumentProcessingException("不支持含宏或 ActiveX 控件的文档。");
                var external = (string?)relationship.Attribute("TargetMode") == "External";
                if (external)
                {
                    if (type.EndsWith("/attachedTemplate", StringComparison.OrdinalIgnoreCase))
                        throw new DocumentProcessingException("文档引用外部模板，请在办公软件中取消外部模板引用后重试。");
                    continue; // Never dereference any external target.
                }
                var resolved = ResolveRelationshipTarget(pair.Key, target);
                if (!parts.ContainsKey(resolved))
                    throw new DocumentProcessingException("文档引用了缺失的内部部件，请在办公软件中修复并另存。");
            }
        }
        var rootRelationships = xmlParts["_rels/.rels"].Root!.Elements(PackageRelationships + "Relationship")
            .Where(element => ((string?)element.Attribute("Type"))?.EndsWith("/officeDocument", StringComparison.Ordinal) == true).ToList();
        if (rootRelationships.Count != 1 ||
            (string?)rootRelationships[0].Attribute("TargetMode") == "External" ||
            ResolveRelationshipTarget("_rels/.rels", (string)rootRelationships[0].Attribute("Target")!) != MainPart)
            throw new DocumentProcessingException("DOCX 主文档关系不符合受支持的格式。");
    }

    private static string ResolveRelationshipTarget(string relationshipsName, string target)
    {
        var decoded = Uri.UnescapeDataString(target);
        if (decoded.Contains('\\') || decoded.Contains(':') || decoded.Any(char.IsControl))
            throw new DocumentProcessingException("文档包含不安全的内部引用，已停止处理。");
        var fragmentIndex = decoded.IndexOf('#');
        if (fragmentIndex >= 0) decoded = decoded[..fragmentIndex];
        if (decoded.Contains('?') || string.IsNullOrEmpty(decoded))
            throw new DocumentProcessingException("文档内部引用无效。");
        var segments = new List<string>();
        if (!decoded.StartsWith('/') && relationshipsName != "_rels/.rels")
        {
            var marker = relationshipsName.LastIndexOf("/_rels/", StringComparison.Ordinal);
            if (marker < 0) throw new DocumentProcessingException("文档关系部件位置无效。");
            segments.AddRange(relationshipsName[..marker].Split('/'));
        }
        foreach (var segment in decoded.TrimStart('/').Split('/'))
        {
            if (segment is "" or ".") continue;
            if (segment == "..")
            {
                if (segments.Count == 0) throw new DocumentProcessingException("文档内部引用超出文件范围，已停止处理。");
                segments.RemoveAt(segments.Count - 1);
            }
            else segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    private static DocumentAnalysis AnalyzePackage(Package package)
    {
        var body = package.Document.Root!.Element(W + "body")!;
        var paragraphs = body.Elements(W + "p").ToList();
        if (paragraphs.Count > 100_000)
            throw new DocumentProcessingException("文档段落过多，不适合自动处理。");
        var warnings = new List<string>();
        var blockingReasons = new List<string>();
        var tableCount = body.Descendants(W + "tbl").Count();
        var imageCount = body.Descendants().Count(element => element.Name == W + "drawing" || element.Name == W + "pict");
        if (tableCount > 0) warnings.Add($"检测到 {tableCount} 个表格：保留表格内容和格式；页面变化可能影响宽度与跨页，请在 WPS / Word 中复核。");
        if (imageCount > 0) warnings.Add($"检测到 {imageCount} 处图片或绘图：保留内容和格式；页面变化可能影响锚点与位置，请人工复核。");
        if (package.Parts.Any(part => part.Name.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase)) || body.Descendants(W + "object").Any())
            warnings.Add("检测到嵌入对象：程序保留对象且不打开它们，请在办公软件中人工检查。");
        if (body.Elements().Any(element => element.Name != W + "p" && element.Name != W + "tbl" && element.Name != W + "sectPr"))
            warnings.Add("检测到内容控件或特殊正文结构：只处理直接正文段落，其他结构保持原样并需人工复核。");
        if (package.Parts.Any(part => Regex.IsMatch(part.Name, @"^word/(header|footer)\d*\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) ||
            body.Descendants(W + "headerReference").Any() || body.Descendants(W + "footerReference").Any())
            warnings.Add("页眉、页脚及其中页码保持原样；首版不自动生成或校验页码，请检查打印结果。");
        if (package.Parts.Any(part => part.Name is "word/comments.xml" or "word/footnotes.xml" or "word/endnotes.xml"))
            warnings.Add("批注、脚注或尾注保留原样，请核对它们在新页面中的显示位置。");
        var relationships = package.XmlParts.Where(pair => pair.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            .SelectMany(pair => pair.Value.Root!.Elements(PackageRelationships + "Relationship"));
        if (relationships.Any(relationship => (string?)relationship.Attribute("TargetMode") == "External"))
            warnings.Add("文档包含外部链接或关联资源：程序不访问、不下载、不更新外部内容。");
        var sections = body.Descendants(W + "sectPr").ToList();
        if (sections.Count > 1 || sections.Any(section => section.Parent != body))
            blockingReasons.Add("文档有分节或多节布局，首版只支持单节文字稿；请在办公软件中人工排版。");
        if (sections.Any(section => section.Descendants(W + "sectPrChange").Any()))
            blockingReasons.Add("页面设置含修订，不能安全自动调整。");
        if (sections.Any(section => section.Elements(W + "cols").Any(columns =>
                ParseInt(columns.Attribute(W + "num"), 1) > 1 || columns.Elements(W + "col").Count() > 1)))
            blockingReasons.Add("文档采用多栏布局，首版暂不支持。");
        if (sections.Any(section => section.Element(W + "rtlGutter") is { } gutter && IsOn(gutter)))
            blockingReasons.Add("文档采用特殊装订布局，首版暂不支持。");
        if (sections.Any(section => section.Element(W + "pgBorders") is not null))
            warnings.Add("文档页面边框保持原样，请检查边框与新页边距的关系。");
        var hasEditProtection = HasEditProtection(package);
        if (hasEditProtection)
            warnings.Add("文档含编辑、写入或节保护：启用副本策略后，仅在新输出中移除这些限制；原稿、修订模式和其他设置保持原样。加密文件不在支持范围内。");
        foreach (var settingsPart in SettingsParts(package))
        {
            var settings = settingsPart.Value;
            if (settings.Descendants(W + "trackRevisions").Any(IsOn))
                warnings.Add("文档启用了修订模式：程序不接受或删除修订，输出后请核对修订设置。");
            if (settings.Descendants().Any(element => element.Name.Namespace == W &&
                    element.Name.LocalName is "mirrorMargins" or "bookFoldPrinting" or "bookFoldRevPrinting" or "printTwoOnOne" && IsOn(element)))
                blockingReasons.Add("文档设置了对称页边距或书册打印，首版暂不支持。");
        }
        // Fields can execute from headers, footers, notes and text boxes too.
        // Inspect every parsed part containing WordprocessingML, without evaluating fields.
        if (package.XmlParts.Values.Any(ContainsDdeInstruction))
            blockingReasons.Add("文档含 DDE 字段，不能安全自动处理。");
        if (body.Descendants().Any(HasRevisionElement))
            warnings.Add("文档包含修订记录：含修订的段落必须保留原格式，程序不会自动接受修订。");

        var infos = new List<ParagraphInfo>(paragraphs.Count);
        var texts = paragraphs.Select(ParagraphText).ToArray();
        var nonEmptyIndices = texts.Select((text, index) => (text, index))
            .Where(pair => !string.IsNullOrWhiteSpace(pair.text)).Select(pair => pair.index).ToList();
        var lastNonEmpty = nonEmptyIndices.Count == 0 ? -1 : nonEmptyIndices[^1];
        var dateIndex = nonEmptyIndices.LastOrDefault(index => index >= lastNonEmpty - 2 && IsPlausibleDate(texts[index].Trim()), -1);
        var firstNonEmpty = nonEmptyIndices.Count == 0 ? -1 : nonEmptyIndices[0];
        var signatureIndex = dateIndex >= 0 && nonEmptyIndices.IndexOf(dateIndex) > 0 ? nonEmptyIndices[nonEmptyIndices.IndexOf(dateIndex) - 1] : -1;
        var numberedCount = 0;
        var complexNumberedCount = 0;
        var protectedCount = 0;
        for (var index = 0; index < paragraphs.Count; index++)
        {
            var paragraph = paragraphs[index];
            var text = texts[index];
            var trimmed = text.Trim();
            var automaticNumbering = HasAutomaticNumbering(paragraph, package);
            if (automaticNumbering) numberedCount++;
            ParagraphRole role;
            bool review;
            string reason;
            if (HasRevision(paragraph))
                (role, review, reason) = (ParagraphRole.Preserve, true, "包含修订，必须保留原格式。");
            else if (!HasSafeNumberingProperties(paragraph, package))
            {
                (role, review, reason) = (ParagraphRole.Preserve, true, "编号结构复杂或无效，必须保留原格式。");
                complexNumberedCount++;
            }
            else if (HasProtectedStructure(paragraph))
            {
                (role, review, reason) = (ParagraphRole.Preserve, true, "包含字段、内容控件或特殊布局，必须保留原格式。");
                protectedCount++;
            }
            else if (trimmed.Length == 0)
                (role, review, reason) = (ParagraphRole.Preserve, false, "空白段落保留，不自动插入或删除空行。");
            else
                (role, review, reason) = SuggestParagraphRole(paragraph, package, trimmed, index, firstNonEmpty,
                    nonEmptyIndices.IndexOf(index), dateIndex, signatureIndex, texts, infos, automaticNumbering);
            infos.Add(new ParagraphInfo(index, text, role, review, reason, CanEditParagraph(paragraph, package), automaticNumbering));
        }
        if (numberedCount > 0) warnings.Add($"检测到 {numberedCount} 段自动编号：保留编号关系、序列、缩进与对齐；允许调整简单文字的字体、字号和内容，不将编号转成文本。");
        if (complexNumberedCount > 0) warnings.Add($"检测到 {complexNumberedCount} 段复杂编号结构：必须保留原格式，请在办公软件中人工处理。");
        if (protectedCount > 0) warnings.Add($"检测到 {protectedCount} 段字段、内容控件或特殊布局：保留这些段落格式，不更新字段。");
        if (infos.Any(info => info.SuggestedRole is ParagraphRole.Signature or ParagraphRole.Date))
            warnings.Add("署名和日期仅按右对齐初步排版；与正文、印章、附件及跨页关系必须人工复核。");
        if (GetNormalFontSize(package) is null)
            warnings.Add("原稿未明确基础字号，字符网格按模板正文估算；请在 WPS / Word 中检查每页行数与每行字数。");
        if (paragraphs.Count == 0)
            blockingReasons.Add("未找到可处理的直接正文段落，可能是扫描件或全部正文位于表格、文本框中。");
        var requiresStructureReview = tableCount > 0 || imageCount > 0 || complexNumberedCount > 0 || protectedCount > 0 ||
            body.Descendants().Any(HasRevisionElement) ||
            body.Elements().Any(element => element.Name != W + "p" && element.Name != W + "sectPr") ||
            package.Parts.Any(part => part.Name.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase) ||
                part.Name is "word/comments.xml" or "word/footnotes.xml" or "word/endnotes.xml" ||
                Regex.IsMatch(part.Name, @"^word/(header|footer)\d*\.xml$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)) ||
            body.Descendants().Any(element => element.Name.Namespace == W && element.Name.LocalName is
                "object" or "headerReference" or "footerReference" or "pgBorders" or "commentRangeStart" or "commentRangeEnd" or "commentReference" or "footnoteReference" or "endnoteReference") ||
            relationships.Any(relationship => (string?)relationship.Attribute("TargetMode") == "External") ||
            SettingsParts(package).Any(part => part.Value.Descendants(W + "trackRevisions").Any(IsOn));
        return new DocumentAnalysis(package.SourcePath, package.Sha256, infos, warnings.Distinct(StringComparer.Ordinal).ToArray(),
            tableCount, imageCount, blockingReasons.Count == 0, blockingReasons.Count == 0 ? null : string.Join("\n", blockingReasons.Distinct(StringComparer.Ordinal)), SuggestTemplate(infos), requiresStructureReview, hasEditProtection);
    }

    private static (ParagraphRole Role, bool Review, string Reason) SuggestParagraphRole(XElement paragraph, Package package,
        string text, int index, int firstNonEmpty, int nonEmptyOrdinal, int dateIndex, int signatureIndex,
        IReadOnlyList<string> allTexts, IReadOnlyList<ParagraphInfo> earlierInfos, bool automaticNumbering)
    {
        // Literal numbering is evidence for typography only. Never insert,
        // remove, normalize or renumber its text characters.
        if (Heading1Regex().IsMatch(text)) return (ParagraphRole.Heading1, false, "明文一级序号，序号文本原样保留。");
        if (Heading2Regex().IsMatch(text)) return (ParagraphRole.Heading2, false, "明文二级序号，序号文本原样保留。");
        if (Heading4Regex().IsMatch(text)) return (ParagraphRole.Heading4, text.Length > 160, "明文四级序号，序号文本原样保留；长条目请核对层次。");
        if (Heading3Regex().IsMatch(text)) return (ParagraphRole.Heading3, text.Length > 160, "明文三级序号，序号文本原样保留；长条目请核对层次。");
        var evidence = GetParagraphEvidence(paragraph, package);
        if (automaticNumbering)
        {
            if (evidence.HeadingLevel is { } numberedHeading)
                return (HeadingRole(numberedHeading), numberedHeading > 4, "自动编号标题按样式或大纲层次识别；编号和缩进保持原样。");
            if (evidence.IsTitleStyle)
                return (ParagraphRole.Title, text.Length > 100, "标题样式提供依据；自动编号及其布局保持原样。");
            return (ParagraphRole.Body, false, "自动编号正文：只处理文字，不改编号、序列或原编号布局。");
        }
        if (text.StartsWith("附件：", StringComparison.Ordinal) || text.StartsWith("附件:", StringComparison.Ordinal))
            return (ParagraphRole.AttachmentNote, false, "明确附件说明标记，附件名称与序号原样保留。");
        if (index == dateIndex) return (ParagraphRole.Date, false, "末尾独立且有效的年月日日期，按成文日期排版。");

        var shortLine = text.Length <= 100 && !SentencePunctuationRegex().IsMatch(text);
        if (index == signatureIndex && text.Length <= 80 && !SentencePunctuationRegex().IsMatch(text))
        {
            if (LooksLikeOrganizationName(text) || evidence.Alignment == "right")
                return (ParagraphRole.Signature, false, "有效成文日期之前的单位名称或右对齐落款。");
            if (!text.StartsWith("特此", StringComparison.Ordinal) && text.Length <= 40)
                return (ParagraphRole.Signature, true, "日期前短段落可能是署名，请核对机构或签发人。");
        }
        if (nonEmptyOrdinal is >= 0 and < 6 && LooksLikeRecipient(text))
            return (ParagraphRole.Recipient, false, "正文前明确的主送单位及冒号标记。");
        if (nonEmptyOrdinal is >= 0 and < 4 && shortLine &&
            (evidence.IsTitleStyle || StructuredTitleRegex().IsMatch(text) || evidence.Alignment == "center" && evidence.FontSize >= 18))
            return (ParagraphRole.Title, false, "标题样式、明确公文标题或居中较大字号提供可靠依据。");
        if (evidence.HeadingLevel is { } heading)
            return (HeadingRole(heading), heading > 4 || text.Length > 180, "段落样式或继承的大纲级别表明标题层次。");
        if (nonEmptyOrdinal > 0 && shortLine && (KnownHeadingRegex().IsMatch(text) || evidence.Bold && evidence.FontSize >= 18))
            return (ParagraphRole.Heading1, false, "短段落的标题用语或加粗较大字号表明无序号标题。");
        if (index == firstNonEmpty && shortLine && text.Length <= 80 && !BodyOpeningRegex().IsMatch(text))
        {
            var nextText = allTexts.Skip(index + 1).FirstOrDefault(next => !string.IsNullOrWhiteSpace(next))?.Trim();
            if (nextText is not null && LooksLikeRecipient(nextText))
                return (ParagraphRole.Title, false, "首段短标题之后是明确主送单位。");
            return (ParagraphRole.Title, true, "首个短段落可能是标题；缺少样式或结构证据，请核对。");
        }
        if (nonEmptyOrdinal is >= 0 and < 6 && shortLine && (text.EndsWith('：') || text.EndsWith(':')) &&
            earlierInfos.Any(info => info.SuggestedRole == ParagraphRole.Title))
            return (ParagraphRole.Recipient, true, "标题后的冒号短段落可能是主送机关，请确认。");
        return (ParagraphRole.Body, false, "按正文规则处理；未明确编辑的原文字内容保持不变。");
    }

    private static ParagraphRole HeadingRole(int level) => level switch
    {
        1 => ParagraphRole.Heading1, 2 => ParagraphRole.Heading2, 3 => ParagraphRole.Heading3, _ => ParagraphRole.Heading4
    };

    private static string? SuggestTemplate(IReadOnlyList<ParagraphInfo> paragraphs)
    {
        foreach (var paragraph in paragraphs.Where(paragraph => paragraph.SuggestedRole == ParagraphRole.Title && !paragraph.NeedsReview))
        {
            var title = paragraph.Text.Trim();
            if (title.EndsWith("通知", StringComparison.Ordinal)) return "notice-v1";
            if (title.EndsWith("请示", StringComparison.Ordinal)) return "request-v1";
            if (title.EndsWith("报告", StringComparison.Ordinal)) return "report-v1";
        }
        return null;
    }

    private static bool LooksLikeRecipient(string text)
    {
        if (text.Length is < 3 or > 100 || !(text.EndsWith('：') || text.EndsWith(':'))) return false;
        var recipient = text[..^1].Trim();
        return !SentencePunctuationRegex().IsMatch(recipient) &&
            (recipient.StartsWith("各", StringComparison.Ordinal) || LooksLikeOrganizationName(recipient) || recipient.EndsWith("同志", StringComparison.Ordinal));
    }

    private static bool LooksLikeOrganizationName(string text) =>
        OrganizationTailRegex().IsMatch(OrganizationNoteRegex().Replace(text, "").TrimEnd());

    private static ParagraphEvidence GetParagraphEvidence(XElement paragraph, Package package)
    {
        var styleChain = GetParagraphStyleChain(paragraph, package);
        XElement? ParagraphProperty(string name) => paragraph.Element(W + "pPr")?.Element(W + name) ??
            styleChain.Select(style => style.Element(W + "pPr")?.Element(W + name)).FirstOrDefault(property => property is not null);
        var dominantRun = paragraph.Elements(W + "r").Where(run => run.Elements(W + "t").Any()).OrderByDescending(run => run.Elements(W + "t").Sum(text => text.Value.Length)).FirstOrDefault();
        XElement? RunProperty(string name) => dominantRun?.Element(W + "rPr")?.Element(W + name) ??
            paragraph.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + name) ??
            styleChain.Select(style => style.Element(W + "rPr")?.Element(W + name)).FirstOrDefault(property => property is not null);
        var labels = styleChain.SelectMany(style => new[] { (string?)style.Attribute(W + "styleId"), (string?)style.Element(W + "name")?.Attribute(W + "val") })
            .Where(label => label is not null).Select(label => label!.Trim()).ToArray();
        var titleStyle = labels.Any(label => label.Equals("Title", StringComparison.OrdinalIgnoreCase) || label is "标题" or "文档标题");
        int? headingLevel = null;
        foreach (var label in labels)
        {
            var match = HeadingStyleRegex().Match(label);
            if (match.Success) { headingLevel = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture); break; }
        }
        if (headingLevel is not null) titleStyle = false;
        var outlineLevel = ParseInt(ParagraphProperty("outlineLvl")?.Attribute(W + "val"), -1);
        if (headingLevel is null && outlineLevel is >= 0 and <= 8) headingLevel = outlineLevel + 1;
        var fontHalfPoints = ParseInt(RunProperty("sz")?.Attribute(W + "val"), -1);
        var fontSize = fontHalfPoints > 0 ? fontHalfPoints / 2.0 : GetNormalFontSize(package);
        var bold = RunProperty("b") is { } boldProperty && IsOn(boldProperty);
        return new ParagraphEvidence(titleStyle, headingLevel, (string?)ParagraphProperty("jc")?.Attribute(W + "val"), fontSize, bold);
    }

    private static List<XElement> GetParagraphStyleChain(XElement paragraph, Package package)
    {
        var styleChain = new List<XElement>();
        if (package.XmlParts.TryGetValue("word/styles.xml", out var styles) && styles.Root is not null)
        {
            var styleId = (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            if (string.IsNullOrEmpty(styleId)) styleId = (string?)styles.Root.Elements(W + "style").FirstOrDefault(style =>
                (string?)style.Attribute(W + "type") == "paragraph" && IsTrueValue((string?)style.Attribute(W + "default")))?.Attribute(W + "styleId");
            var seen = new HashSet<string>(StringComparer.Ordinal);
            while (styleId is not null && seen.Add(styleId))
            {
                var style = styles.Root.Elements(W + "style").FirstOrDefault(candidate => (string?)candidate.Attribute(W + "styleId") == styleId);
                if (style is null) break;
                styleChain.Add(style);
                styleId = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
            }
        }
        return styleChain;
    }

    private static bool IsPlausibleDate(string text)
    {
        var match = DateTextRegex().Match(text);
        if (!match.Success) return false;
        static int Digit(char character) => "〇零一二三四五六七八九".IndexOf(character) is var index && index >= 0 ? Math.Max(0, index - 1) : -1;
        static int Number(string value)
        {
            if (int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var number)) return number;
            if (value.Length == 1) return value == "十" ? 10 : Digit(value[0]);
            var ten = value.IndexOf('十');
            if (ten is < 0 or > 1 || value.Length - ten - 1 > 1) return -1;
            var tens = ten == 0 ? 1 : Digit(value[0]);
            var units = ten == value.Length - 1 ? 0 : Digit(value[^1]);
            return tens >= 0 && units >= 0 ? tens * 10 + units : -1;
        }
        var yearText = match.Groups[1].Value;
        var year = int.TryParse(yearText, NumberStyles.None, CultureInfo.InvariantCulture, out var numericYear) ? numericYear :
            yearText.Aggregate(0, (value, character) => value * 10 + Digit(character));
        var month = Number(match.Groups[2].Value);
        var day = Number(match.Groups[3].Value);
        return year is >= 1 and <= 9999 && month is >= 1 and <= 12 && day >= 1 && day <= DateTime.DaysInMonth(year, month);
    }

    private static bool HasRevision(XElement paragraph) => paragraph.DescendantsAndSelf().Any(HasRevisionElement);

    private static bool ContainsDdeInstruction(XDocument document)
    {
        var fields = new Stack<FieldInstruction>();
        var orphanInstruction = new StringBuilder();
        var dangerous = false;

        static void AppendBounded(StringBuilder buffer, string value)
        {
            if (buffer.Length > 1024 * 1024 - value.Length)
                throw new DocumentProcessingException("字段指令或嵌套结果过大，不能安全自动检查。");
            buffer.Append(value);
        }

        void PushField(FieldInstruction field)
        {
            if (fields.Count >= 128)
                throw new DocumentProcessingException("字段嵌套过深，不能安全自动检查。");
            fields.Push(field);
        }

        void FinishField(FieldInstruction field)
        {
            if (DdeInstructionRegex().IsMatch(field.Instruction.ToString())) dangerous = true;
            // A nested field's cached result belongs to its parent's instruction or
            // result. Inspect the reconstructed outer command as well as each child.
            if (fields.TryPeek(out var parent))
                AppendBounded(parent.InInstruction ? parent.Instruction : parent.Result, field.Result.ToString());
        }

        void Visit(XElement element)
        {
            if (element.Name == W + "fldSimple")
            {
                var simple = new FieldInstruction(element, inInstruction: false);
                AppendBounded(simple.Instruction, (string?)element.Attribute(W + "instr") ?? "");
                PushField(simple);
                foreach (var child in element.Elements()) Visit(child);
                // Malformed/unclosed nested fields are still scanned, never skipped.
                while (fields.Count > 0)
                {
                    var finished = fields.Pop();
                    FinishField(finished);
                    if (ReferenceEquals(finished.SimpleElement, element)) break;
                }
                return;
            }
            if (element.Name == W + "fldChar")
            {
                var type = (string?)element.Attribute(W + "fldCharType");
                if (type == "begin") PushField(new FieldInstruction(null, inInstruction: true));
                else if (type == "separate" && fields.TryPeek(out var active) && active.SimpleElement is null)
                    active.InInstruction = false;
                else if (type == "end" && fields.TryPeek(out var closing) && closing.SimpleElement is null)
                    FinishField(fields.Pop());
            }
            else if (element.Name == W + "instrText" || element.Name == W + "delInstrText")
            {
                // Do not insert spaces between runs: DD + EAUTO is one command.
                if (fields.TryPeek(out var active) && active.InInstruction) AppendBounded(active.Instruction, element.Value);
                else AppendBounded(orphanInstruction, element.Value);
            }
            else if (element.Name == W + "t" || element.Name == W + "delText")
            {
                if (fields.TryPeek(out var active))
                {
                    if (active.InInstruction) AppendBounded(active.Instruction, element.Value);
                    else if (fields.Count > 1) AppendBounded(active.Result, element.Value);
                }
            }
            foreach (var child in element.Elements()) Visit(child);
        }

        if (document.Root is not null) Visit(document.Root);
        // Retain detection even for a corrupt field with no closing marker.
        while (fields.Count > 0) FinishField(fields.Pop());
        return dangerous || DdeInstructionRegex().IsMatch(orphanInstruction.ToString());
    }

    private static bool HasRevisionElement(XElement element) => element.Name.Namespace == W &&
        (element.Name.LocalName.EndsWith("Change", StringComparison.Ordinal) || element.Name.LocalName is
            "ins" or "del" or "delText" or "delInstrText" or "moveFrom" or "moveTo" or
            "moveFromRangeStart" or "moveFromRangeEnd" or "moveToRangeStart" or "moveToRangeEnd" or
            "customXmlInsRangeStart" or "customXmlInsRangeEnd" or "customXmlDelRangeStart" or "customXmlDelRangeEnd" or
            "customXmlMoveFromRangeStart" or "customXmlMoveFromRangeEnd" or "customXmlMoveToRangeStart" or "customXmlMoveToRangeEnd" or "cellIns" or "cellDel" or "cellMerge");

    private static bool HasProtectedStructure(XElement paragraph) => paragraph.Descendants().Any(element => element.Name.Namespace == W &&
        element.Name.LocalName is "fldChar" or "instrText" or "fldSimple" or "sdt" or "framePr" or "altChunk" or "txbxContent");

    private static readonly HashSet<string> EditableParagraphProperties = new(StringComparer.Ordinal)
    {
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "suppressLineNumbers", "snapToGrid",
        "spacing", "ind", "contextualSpacing", "jc", "textAlignment", "outlineLvl", "rPr", "numPr", "tabs"
    };
    private static readonly HashSet<string> EditableRunProperties = new(StringComparer.Ordinal)
    {
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "noProof",
        "snapToGrid", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "vertAlign", "rtl", "cs", "em", "lang"
    };

    private static bool CanEditParagraph(XElement paragraph, Package package)
    {
        if (HasRevision(paragraph) || !HasSafeNumberingProperties(paragraph, package) || HasProtectedStructure(paragraph)) return false;
        if (paragraph.Elements(W + "pPr").Count() > 1 || paragraph.Elements().Any(element => element.Name != W + "pPr" && element.Name != W + "r")) return false;
        static bool SimpleRunProperties(XElement properties) => properties.Elements().All(property =>
            property.Name.Namespace == W && EditableRunProperties.Contains(property.Name.LocalName) && !property.HasElements);
        if (paragraph.Element(W + "pPr") is { } paragraphProperties)
            foreach (var property in paragraphProperties.Elements())
            {
                if (property.Name.Namespace != W || !EditableParagraphProperties.Contains(property.Name.LocalName)) return false;
                if (property.Name == W + "rPr") { if (!SimpleRunProperties(property)) return false; }
                else if (property.Name == W + "numPr") { if (!IsSimpleNumberingProperties(property)) return false; }
                else if (property.Name == W + "tabs") { if (!IsSimpleTabStops(property)) return false; }
                else if (property.HasElements) return false;
            }
        foreach (var run in paragraph.Elements(W + "r"))
        {
            if (run.Elements(W + "rPr").Count() > 1 || run.Elements().Any(element => element.Name != W + "rPr" && element.Name != W + "t")) return false;
            if (run.Element(W + "rPr") is { } properties && !SimpleRunProperties(properties)) return false;
            if (run.Elements(W + "t").Any(text => text.Nodes().Any(node => node is not XText) || text.Attributes().Any(attribute => attribute.Name != XNamespace.Xml + "space"))) return false;
        }
        return true;
    }

    private static bool IsSimpleTabStops(XElement tabs)
    {
        if (tabs.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || tabs.Elements().Count() > 256) return false;
        foreach (var tab in tabs.Elements())
        {
            if (tab.Name != W + "tab" || tab.HasElements || tab.Attributes().Any(attribute =>
                !attribute.IsNamespaceDeclaration && attribute.Name != W + "val" && attribute.Name != W + "pos" && attribute.Name != W + "leader")) return false;
            var alignment = (string?)tab.Attribute(W + "val");
            if (alignment is not ("clear" or "left" or "center" or "right" or "decimal" or "bar" or "num" or "start" or "end")) return false;
            var position = ParseInt(tab.Attribute(W + "pos"), int.MinValue);
            if (position is < -31680 or > 31680) return false;
            var leader = (string?)tab.Attribute(W + "leader");
            if (leader is not (null or "none" or "dot" or "hyphen" or "underscore" or "heavy" or "middleDot")) return false;
        }
        return true;
    }

    private static void ApplyParagraphEdits(XDocument document, IReadOnlyDictionary<int, string>? edits, Package package)
    {
        if (edits is null || edits.Count == 0) return;
        var paragraphs = document.Root!.Element(W + "body")!.Elements(W + "p").ToList();
        foreach (var (index, newText) in edits)
        {
            if (index < 0 || index >= paragraphs.Count)
                throw new DocumentProcessingException("编辑段落选择已失效，请重新导入文档。");
            if (!CanEditParagraph(paragraphs[index], package))
                throw new DocumentProcessingException($"第 {index + 1} 段包含受保护或复杂结构，请在 WPS / Word 中修改后保存，再重新导入。");
            if (newText is null || newText.Length > 100_000)
                throw new DocumentProcessingException("每段编辑文本不得超过 100,000 个字符。");
            if (newText.IndexOfAny(['\r', '\n', '\t', '\u2028', '\u2029']) >= 0)
                throw new DocumentProcessingException("段落编辑不支持换行或制表符，请在办公软件中修改段落结构。");
            try { XmlConvert.VerifyXmlChars(newText); }
            catch (XmlException) { throw new DocumentProcessingException("编辑文本包含非法 XML 字符，请移除控制字符后重试。"); }
            if (ParagraphText(paragraphs[index]) == newText) continue;
            var texts = paragraphs[index].Elements(W + "r").SelectMany(run => run.Elements(W + "t")).ToList();
            if (texts.Count == 0)
            {
                var run = paragraphs[index].Elements(W + "r").FirstOrDefault();
                if (run is null) { run = new XElement(W + "r"); paragraphs[index].Add(run); }
                var text = new XElement(W + "t");
                run.Add(text);
                texts.Add(text);
            }
            var offset = 0;
            for (var textIndex = 0; textIndex < texts.Count; textIndex++)
            {
                var text = texts[textIndex];
                var length = textIndex == texts.Count - 1 ? newText.Length - offset : Math.Min(text.Value.Length, newText.Length - offset);
                // Avoid splitting a new surrogate pair across formatting runs.
                if (length > 0 && offset + length < newText.Length && char.IsHighSurrogate(newText[offset + length - 1]) && char.IsLowSurrogate(newText[offset + length])) length--;
                text.Value = newText.Substring(offset, length);
                text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                offset += length;
            }
        }
    }

    private static bool HasAutomaticNumbering(XElement paragraph, Package package)
    {
        var sawNumberingProperties = false;
        foreach (var properties in ParagraphNumberingProperties(paragraph, package))
        {
            sawNumberingProperties = true;
            if (properties.Element(W + "numId") is { } numberId)
                return ParseInt(numberId.Attribute(W + "val"), -1) != 0;
        }
        // An incomplete numPr is conservatively treated as numbering so its
        // positioning is never replaced with body indentation.
        return sawNumberingProperties;
    }

    private static IEnumerable<XElement> ParagraphNumberingProperties(XElement paragraph, Package package)
    {
        foreach (var properties in paragraph.Element(W + "pPr")?.Elements(W + "numPr") ?? []) yield return properties;
        foreach (var style in GetParagraphStyleChain(paragraph, package))
            foreach (var properties in style.Element(W + "pPr")?.Elements(W + "numPr") ?? []) yield return properties;
    }

    private static bool HasSafeNumberingProperties(XElement paragraph, Package package) =>
        (paragraph.Element(W + "pPr")?.Elements(W + "numPr").Count() ?? 0) <= 1 &&
        ParagraphNumberingProperties(paragraph, package).All(IsSimpleNumberingProperties);

    private static bool IsSimpleNumberingProperties(XElement properties)
    {
        if (properties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)) return false;
        if (properties.Elements(W + "numId").Count() > 1 || properties.Elements(W + "ilvl").Count() > 1) return false;
        foreach (var child in properties.Elements())
        {
            if (child.Name != W + "numId" && child.Name != W + "ilvl" || child.HasElements ||
                child.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != W + "val")) return false;
            var value = ParseInt(child.Attribute(W + "val"), -1);
            if (value < 0 || child.Name == W + "ilvl" && value > 8) return false;
        }
        return true;
    }

    private static IEnumerable<KeyValuePair<string, XDocument>> SettingsParts(Package package) =>
        package.XmlParts.Where(part => part.Value.Root?.Name == W + "settings");

    private static bool HasEditProtection(Package package) =>
        SettingsParts(package).Any(part => part.Value.Descendants().Any(element => element.Name == W + "documentProtection" || element.Name == W + "writeProtection")) ||
        package.Document.Descendants(W + "sectPr").Elements(W + "formProt").Any();

    private static bool RemoveEditProtectionFromCopy(Package package, XDocument document, IDictionary<string, byte[]> changedParts)
    {
        var removed = false;
        foreach (var part in SettingsParts(package))
        {
            var copy = new XDocument(part.Value);
            var restrictions = copy.Descendants().Where(element => element.Name == W + "documentProtection" || element.Name == W + "writeProtection").ToList();
            if (restrictions.Count == 0) continue;
            foreach (var restriction in restrictions) restriction.Remove();
            var expected = new XDocument(part.Value);
            foreach (var restriction in expected.Descendants().Where(element => element.Name == W + "documentProtection" || element.Name == W + "writeProtection").ToList()) restriction.Remove();
            if (!XNode.DeepEquals(expected.Root, copy.Root))
                throw new DocumentProcessingException("输出副本校验失败：编辑限制之外的设置发生变化，未保存输出。");
            changedParts.Add(part.Key, Serialize(copy));
            removed = true;
        }
        foreach (var restriction in document.Descendants(W + "sectPr").Elements(W + "formProt").ToList())
        {
            restriction.Remove();
            removed = true;
        }
        return removed;
    }

    private static double? GetNormalFontSize(Package package)
    {
        if (!package.XmlParts.TryGetValue("word/styles.xml", out var styles) || styles.Root is null) return null;
        var paragraphStyles = styles.Root.Elements(W + "style").Where(style => (string?)style.Attribute(W + "type") == "paragraph").ToList();
        var normal = paragraphStyles.FirstOrDefault(style => (string?)style.Attribute(W + "styleId") == "Normal")
            ?? paragraphStyles.FirstOrDefault(style => string.Equals((string?)style.Element(W + "name")?.Attribute(W + "val"), "Normal", StringComparison.OrdinalIgnoreCase))
            ?? paragraphStyles.FirstOrDefault(style => IsTrueValue((string?)style.Attribute(W + "default")));
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (normal is not null)
        {
            var styleId = (string?)normal.Attribute(W + "styleId");
            if (styleId is null || !visited.Add(styleId)) break;
            var halfPoints = ParseInt(normal.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"), -1);
            if (halfPoints is >= 2 and <= 400) return halfPoints / 2.0;
            var parentId = (string?)normal.Element(W + "basedOn")?.Attribute(W + "val");
            normal = paragraphStyles.FirstOrDefault(style => (string?)style.Attribute(W + "styleId") == parentId);
        }
        var defaultHalfPoints = ParseInt(styles.Root.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"), -1);
        return defaultHalfPoints is >= 2 and <= 400 ? defaultHalfPoints / 2.0 : null;
    }

    private static string ParagraphText(XElement paragraph)
    {
        var text = new StringBuilder();
        foreach (var element in paragraph.Descendants())
        {
            if (element.Ancestors().Any(ancestor => ancestor.Name == W + "del" || ancestor.Name == W + "moveFrom")) continue;
            if (element.Name == W + "t") text.Append(element.Value);
            else if (element.Name == W + "tab") text.Append('\t');
            else if (element.Name == W + "br" || element.Name == W + "cr") text.Append('\n');
        }
        return text.ToString();
    }

    private static void FormatParagraph(XElement paragraph, ParagraphRole role, FormatTemplate template, bool automaticNumbering)
    {
        var paragraphProperties = paragraph.Element(W + "pPr");
        if (paragraphProperties is null)
        {
            paragraphProperties = new XElement(W + "pPr");
            paragraph.AddFirst(paragraphProperties);
        }
        var title = role == ParagraphRole.Title;
        var heading = role is ParagraphRole.Heading1 or ParagraphRole.Heading2 or ParagraphRole.Heading3 or ParagraphRole.Heading4;
        var rightAligned = role is ParagraphRole.Signature or ParagraphRole.Date;
        if (!automaticNumbering)
            SetProperty(paragraphProperties, "jc", ("val", title ? "center" : rightAligned ? "right" : "both"));
        SetProperty(paragraphProperties, "keepNext", ("val", title || heading ? "1" : "0"));
        SetProperty(paragraphProperties, "keepLines", ("val", title || heading ? "1" : "0"));
        SetProperty(paragraphProperties, "snapToGrid", ("val", title ? "0" : "1"));
        var spacing = EnsureProperty(paragraphProperties, "spacing");
        foreach (var name in new[] { "beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing" }) spacing.Attribute(W + name)?.Remove();
        SetProperty(paragraphProperties, "spacing", ("before", "0"), ("after", "0"), ("line", "240"), ("lineRule", "auto"));
        if (!automaticNumbering)
        {
            var indent = EnsureProperty(paragraphProperties, "ind");
            foreach (var name in new[] { "left", "right", "start", "end", "leftChars", "rightChars", "startChars", "endChars", "firstLine", "firstLineChars", "hanging", "hangingChars" })
                indent.Attribute(W + name)?.Remove();
            indent.SetAttributeValue(W + "left", "0");
            indent.SetAttributeValue(W + "right", "0");
            if (title || rightAligned || role == ParagraphRole.Recipient)
                indent.SetAttributeValue(W + "firstLine", "0");
            else
                indent.SetAttributeValue(W + "firstLineChars", "200");
        }
        SetProperty(paragraphProperties, "outlineLvl", ("val", role switch
        {
            ParagraphRole.Heading1 => "0", ParagraphRole.Heading2 => "1", ParagraphRole.Heading3 => "2", ParagraphRole.Heading4 => "3", _ => "9"
        }));

        var font = role switch
        {
            ParagraphRole.Title => template.TitleFont,
            ParagraphRole.Heading1 => template.Heading1Font,
            ParagraphRole.Heading2 => template.Heading2Font,
            _ => template.BodyFont
        };
        var pointSize = title ? template.TitleFontPt : template.BodyFontPt;
        // Paragraph-mark formatting avoids inherited mark fonts changing line metrics.
        var paragraphRunProperties = EnsureProperty(paragraphProperties, "rPr");
        SetRunFont(paragraphRunProperties, font, template.LatinFont, pointSize, !title);
        foreach (var run in paragraph.Descendants(W + "r"))
        {
            if (!run.Elements(W + "t").Any() || run.Ancestors(W + "drawing").Any() || run.Ancestors(W + "pict").Any()) continue;
            var runProperties = run.Element(W + "rPr");
            if (runProperties is null)
            {
                runProperties = new XElement(W + "rPr");
                run.AddFirst(runProperties);
            }
            SetRunFont(runProperties, font, template.LatinFont, pointSize, !title);
        }
    }

    private static void SetRunFont(XElement runProperties, string eastAsiaFont, string latinFont, double pointSize, bool snapToGrid)
    {
        var fonts = EnsureProperty(runProperties, "rFonts");
        foreach (var name in new[] { "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme" }) fonts.Attribute(W + name)?.Remove();
        fonts.SetAttributeValue(W + "ascii", latinFont);
        fonts.SetAttributeValue(W + "hAnsi", latinFont);
        fonts.SetAttributeValue(W + "eastAsia", eastAsiaFont);
        fonts.SetAttributeValue(W + "cs", latinFont);
        fonts.SetAttributeValue(W + "hint", "eastAsia");
        SetProperty(runProperties, "sz", ("val", Math.Round(pointSize * 2, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture)));
        SetProperty(runProperties, "szCs", ("val", Math.Round(pointSize * 2, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture)));
        SetProperty(runProperties, "snapToGrid", ("val", snapToGrid ? "1" : "0"));
        var color = EnsureProperty(runProperties, "color");
        foreach (var name in new[] { "themeColor", "themeTint", "themeShade" }) color.Attribute(W + name)?.Remove();
        color.SetAttributeValue(W + "val", "000000");
    }

    private static void FormatSection(XElement body, FormatTemplate template, double normalFontSize)
    {
        var section = body.Element(W + "sectPr");
        if (section is null)
        {
            section = new XElement(W + "sectPr");
            body.Add(section);
        }
        const int pageWidth = 11906;
        const int pageHeight = 16838;
        SetProperty(section, "pgSz", ("w", pageWidth.ToString(CultureInfo.InvariantCulture)), ("h", pageHeight.ToString(CultureInfo.InvariantCulture)), ("orient", "portrait"));
        var margins = EnsureProperty(section, "pgMar");
        margins.SetAttributeValue(W + "top", CmToTwips(template.TopMarginCm));
        margins.SetAttributeValue(W + "bottom", CmToTwips(template.BottomMarginCm));
        margins.SetAttributeValue(W + "left", CmToTwips(template.LeftMarginCm));
        margins.SetAttributeValue(W + "right", CmToTwips(template.RightMarginCm));
        margins.SetAttributeValue(W + "gutter", "0");
        if (margins.Attribute(W + "header") is null) margins.SetAttributeValue(W + "header", "850");
        if (margins.Attribute(W + "footer") is null) margins.SetAttributeValue(W + "footer", "992");
        var textHeight = pageHeight - CmToTwips(template.TopMarginCm) - CmToTwips(template.BottomMarginCm);
        var textWidthPoints = (pageWidth - CmToTwips(template.LeftMarginCm) - CmToTwips(template.RightMarginCm)) / 20.0;
        var linePitch = (int)Math.Round(textHeight / (double)template.LinesPerPage, MidpointRounding.AwayFromZero);
        // OOXML defines charSpace as a 1/4096-point delta from the Normal font.
        // Keep the styles part untouched and derive its existing default size.
        var characterSpacing = (int)Math.Round((textWidthPoints / template.CharactersPerLine - normalFontSize) * 4096, MidpointRounding.AwayFromZero);
        SetProperty(section, "docGrid", ("type", "linesAndChars"), ("linePitch", linePitch.ToString(CultureInfo.InvariantCulture)), ("charSpace", characterSpacing.ToString(CultureInfo.InvariantCulture)));
    }

    private static readonly string[] ParagraphPropertyOrder = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
    private static readonly string[] RunPropertyOrder = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"];
    private static readonly string[] SectionPropertyOrder = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange"];

    private static XElement EnsureProperty(XElement parent, string name)
    {
        var existing = parent.Elements(W + name).ToList();
        if (existing.Count > 0)
        {
            foreach (var duplicate in existing.Skip(1)) duplicate.Remove();
            return existing[0];
        }
        var element = new XElement(W + name);
        var order = parent.Name.LocalName switch { "pPr" => ParagraphPropertyOrder, "rPr" => RunPropertyOrder, "sectPr" => SectionPropertyOrder, _ => [] };
        var index = Array.IndexOf(order, name);
        var next = parent.Elements().FirstOrDefault(child => child.Name.Namespace == W && Array.IndexOf(order, child.Name.LocalName) > index);
        if (next is null) parent.Add(element); else next.AddBeforeSelf(element);
        return element;
    }

    private static void SetProperty(XElement parent, string name, params (string Name, string Value)[] attributes)
    {
        var element = EnsureProperty(parent, name);
        foreach (var (attributeName, value) in attributes) element.SetAttributeValue(W + attributeName, value);
    }

    private static void VerifyDocumentMutation(XDocument original, XDocument changed, IReadOnlyList<ParagraphRole> roles, Package package)
    {
        var originalBody = original.Root!.Element(W + "body")!;
        var changedBody = changed.Root!.Element(W + "body")!;
        var originalParagraphs = originalBody.Elements(W + "p").ToList();
        var changedParagraphs = changedBody.Elements(W + "p").ToList();
        if (originalParagraphs.Count != changedParagraphs.Count || !TextFingerprint(original).SequenceEqual(TextFingerprint(changed)))
            throw new DocumentProcessingException("输出校验失败：正文或字段内容发生变化，未保存输出。");
        for (var index = 0; index < originalParagraphs.Count; index++)
        {
            if (roles[index] == ParagraphRole.Preserve && !XNode.DeepEquals(originalParagraphs[index], changedParagraphs[index]))
                throw new DocumentProcessingException("输出校验失败：保留段落发生变化，未保存输出。");
            if (HasAutomaticNumbering(originalParagraphs[index], package))
                foreach (var name in new[] { "numPr", "ind", "jc", "tabs", "pStyle" })
                    if (!XNode.DeepEquals(originalParagraphs[index].Element(W + "pPr")?.Element(W + name), changedParagraphs[index].Element(W + "pPr")?.Element(W + name)))
                        throw new DocumentProcessingException("输出校验失败：自动编号或其布局发生变化，未保存输出。");
        }
        var originalProtected = originalBody.Elements().Where(element => element.Name != W + "p" && element.Name != W + "sectPr").ToList();
        var changedProtected = changedBody.Elements().Where(element => element.Name != W + "p" && element.Name != W + "sectPr").ToList();
        if (originalProtected.Count != changedProtected.Count || originalProtected.Where((element, index) => !XNode.DeepEquals(element, changedProtected[index])).Any())
            throw new DocumentProcessingException("输出校验失败：表格或特殊结构发生变化，未保存输出。");
        var protectedNames = new HashSet<XName> { W + "drawing", W + "pict", W + "object", W + "bookmarkStart", W + "bookmarkEnd", W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference" };
        var originalObjects = original.Descendants().Where(element => protectedNames.Contains(element.Name)).ToList();
        var changedObjects = changed.Descendants().Where(element => protectedNames.Contains(element.Name)).ToList();
        if (originalObjects.Count != changedObjects.Count || originalObjects.Where((element, index) => !XNode.DeepEquals(element, changedObjects[index])).Any())
            throw new DocumentProcessingException("输出校验失败：图片或引用标记发生变化，未保存输出。");
    }

    private static byte[] TextFingerprint(XDocument document)
    {
        using var memory = new MemoryStream();
        using (var writer = new BinaryWriter(memory, Encoding.UTF8, leaveOpen: true))
            foreach (var element in document.Descendants().Where(element => element.Name.Namespace == W && element.Name.LocalName is "t" or "delText" or "instrText" or "delInstrText" or "tab" or "br" or "cr" or "fldChar" or "fldSimple"))
            {
                writer.Write(element.Name.LocalName);
                writer.Write(element.Value);
                foreach (var attribute in element.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal))
                {
                    writer.Write(attribute.Name.ToString());
                    writer.Write(attribute.Value);
                }
                writer.Write("\0");
            }
        return SHA256.HashData(memory.ToArray());
    }

    private static void VerifyOutputPackage(string temporaryPath, Package source, XDocument expectedDocument, IReadOnlyDictionary<string, byte[]> changedParts)
    {
        using var output = new FileStream(temporaryPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var zip = new ZipArchive(output, ZipArchiveMode.Read, false, Encoding.UTF8);
        var parts = ReadParts(zip);
        if (parts.Count != source.Parts.Count)
            throw new DocumentProcessingException("输出校验失败：文件部件数量变化，未保存输出。");
        var byName = parts.ToDictionary(part => part.Name, StringComparer.Ordinal);
        foreach (var originalPart in source.Parts)
        {
            if (!byName.TryGetValue(originalPart.Name, out var part))
                throw new DocumentProcessingException("输出校验失败：缺少文件部件，未保存输出。");
            if (originalPart.Name == MainPart)
            {
                if (!XNode.DeepEquals(expectedDocument.Root, ParseXml(part.Bytes).Root))
                    throw new DocumentProcessingException("输出校验失败：主文档格式不一致，未保存输出。");
            }
            else if (changedParts.TryGetValue(originalPart.Name, out var expectedPart))
            {
                if (!source.XmlParts.TryGetValue(originalPart.Name, out var sourceXml) || sourceXml.Root?.Name != W + "settings")
                    throw new DocumentProcessingException("输出校验失败：存在未授权修改的文件部件，未保存输出。");
                var permitted = new XDocument(sourceXml);
                foreach (var restriction in permitted.Descendants().Where(element => element.Name == W + "documentProtection" || element.Name == W + "writeProtection").ToList()) restriction.Remove();
                if (!XNode.DeepEquals(permitted.Root, ParseXml(part.Bytes).Root) || !expectedPart.AsSpan().SequenceEqual(part.Bytes))
                    throw new DocumentProcessingException("输出校验失败：编辑限制之外的设置发生变化，未保存输出。");
            }
            else if (!originalPart.Bytes.AsSpan().SequenceEqual(part.Bytes))
                throw new DocumentProcessingException("输出校验失败：未修改的文件部件内容发生变化，未保存输出。");
        }
    }

    private static byte[] Serialize(XDocument document)
    {
        using var memory = new MemoryStream();
        var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, OmitXmlDeclaration = false, NewLineHandling = NewLineHandling.None };
        using (var writer = XmlWriter.Create(memory, settings)) document.Save(writer);
        return memory.ToArray();
    }

    private static string ValidateOutputPath(string sourcePath, string requestedOutput)
    {
        if (string.IsNullOrWhiteSpace(requestedOutput)) throw new DocumentProcessingException("请选择新文档的保存位置。");
        string outputPath;
        try { outputPath = Path.GetFullPath(requestedOutput); }
        catch (Exception error) when (error is ArgumentException or NotSupportedException or PathTooLongException)
        { throw new DocumentProcessingException("输出路径无效。"); }
        if (!string.Equals(Path.GetExtension(outputPath), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new DocumentProcessingException("输出文件必须使用 .docx 扩展名。");
        if (string.Equals(sourcePath, outputPath, StringComparison.OrdinalIgnoreCase))
            throw new DocumentProcessingException("必须另存新文件，不能覆盖原稿。");
        if (File.Exists(outputPath) || Directory.Exists(outputPath))
            throw new DocumentProcessingException("目标文件已存在，请使用新文件名；程序不会覆盖已有文档。");
        if (!Directory.Exists(Path.GetDirectoryName(outputPath)))
            throw new DocumentProcessingException("目标目录不存在，请选择已有且可写的目录。");
        return outputPath;
    }

    private static void ValidateTemplate(FormatTemplate template)
    {
        var fonts = new[] { template.TitleFont, template.BodyFont, template.Heading1Font, template.Heading2Font, template.LatinFont };
        if (fonts.Any(font => string.IsNullOrWhiteSpace(font) || font.Length > 100 || font.Any(char.IsControl)))
            throw new DocumentProcessingException("模板字体名称无效。");
        var margins = new[] { template.TopMarginCm, template.BottomMarginCm, template.LeftMarginCm, template.RightMarginCm };
        if (margins.Any(margin => !double.IsFinite(margin) || margin < 0.5 || margin > 8) ||
            template.TopMarginCm + template.BottomMarginCm >= 24.7 || template.LeftMarginCm + template.RightMarginCm >= 16 ||
            !double.IsFinite(template.BodyFontPt) || template.BodyFontPt is < 5 or > 48 ||
            !double.IsFinite(template.TitleFontPt) || template.TitleFontPt is < 5 or > 72 ||
            template.LinesPerPage is < 8 or > 60 || template.CharactersPerLine is < 10 or > 60)
            throw new DocumentProcessingException("模板页面、字号或网格参数超出安全范围。");
    }

    private static int CmToTwips(double centimeters) => (int)Math.Round(centimeters * 1440 / 2.54, MidpointRounding.AwayFromZero);
    private static int ParseInt(XAttribute? attribute, int fallback) => int.TryParse(attribute?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : fallback;
    private static bool IsOn(XElement element) => element.Attribute(W + "val") is not { } attribute || IsTrueValue(attribute.Value);
    private static bool IsTrueValue(string? value) => value is "1" or "true" or "on";
    private static string HashFile(string path)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(file)).ToLowerInvariant();
    }

    private static byte[] ReadStableSource(string path)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var length = input.Length;
        if (length is <= 0 or > MaximumInputBytes)
            throw new DocumentProcessingException("输入文档必须在 50 MB 以内，且不能为空。");
        var bytes = new byte[checked((int)length)];
        input.ReadExactly(bytes);
        if (input.Length != length)
            throw new DocumentProcessingException("原稿正在保存或同步，读取期间长度发生变化；请完成保存后重新导入。");
        var firstHash = Convert.ToHexString(SHA256.HashData(bytes));
        // Reopen by path so an editor's replace-and-rename save is also detected.
        using var check = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var secondHash = HashSourceHandle(check, out var secondLength);
        if (secondLength != length || !string.Equals(firstHash, secondHash, StringComparison.Ordinal))
            throw new DocumentProcessingException("原稿正在保存或同步，读取期间内容发生变化；请完成保存后重新导入。");
        return bytes;
    }

    private static string HashSourceHandle(FileStream source, out long length)
    {
        length = source.Length;
        if (length is <= 0 or > MaximumInputBytes)
            throw new DocumentProcessingException("原稿为空或超过 50 MB，无法安全核验。");
        source.Position = 0;
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[81920];
        var remaining = length;
        while (remaining > 0)
        {
            var read = source.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
            if (read == 0)
                throw new DocumentProcessingException("原稿在读取期间发生变化，请完成保存后重新导入。");
            hash.AppendData(buffer, 0, read);
            remaining -= read;
        }
        if (source.Length != length)
            throw new DocumentProcessingException("原稿在读取期间发生变化，请完成保存后重新导入。");
        return Convert.ToHexString(hash.GetHashAndReset());
    }

    private static FileStream OpenVerifiedSource(Package package)
    {
        FileStream? source = null;
        try
        {
            source = new FileStream(package.SourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            var firstHash = HashSourceHandle(source, out var firstLength);
            using var second = new FileStream(package.SourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            var secondHash = HashSourceHandle(second, out var secondLength);
            if (firstLength != secondLength || !string.Equals(firstHash, secondHash, StringComparison.Ordinal) ||
                !string.Equals(secondHash, package.Sha256, StringComparison.OrdinalIgnoreCase))
                throw new DocumentProcessingException("原稿在排版期间发生变化，已停止保存；请重新导入和分析。");
            return source;
        }
        catch (DocumentProcessingException)
        {
            source?.Dispose();
            throw;
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            source?.Dispose();
            throw new DocumentProcessingException("保存前无法重新核验原稿，请检查原稿是否被移动、删除或占用后重新导入。");
        }
    }

    [GeneratedRegex(@"\bDDE(?:AUTO)?\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex DdeInstructionRegex();

    [GeneratedRegex(@"[，,。！？!?；;：:]", RegexOptions.CultureInvariant)]
    private static partial Regex SentencePunctuationRegex();
    [GeneratedRegex(@"(?:公司|集团|委员会|人民政府|政府|办公厅|办公室|部门|单位|局|厅|院|处|部|中心|协会|学校|大学|工会|分公司)$", RegexOptions.CultureInvariant)]
    private static partial Regex OrganizationTailRegex();
    [GeneratedRegex(@"(?:（[^（）]{1,16}）|\([^()]{1,16}\))$", RegexOptions.CultureInvariant)]
    private static partial Regex OrganizationNoteRegex();
    [GeneratedRegex(@"^(?:关于.+(?:通知|请示|报告|决定|公告|通告|通报|函|意见|办法|规定|方案|纪要)|.+(?:的通知|的请示|的报告|工作报告|情况报告|调研报告|年度报告|总结报告|实施方案|工作方案|会议纪要|管理办法|管理规定)|通知|请示|报告)$", RegexOptions.CultureInvariant)]
    private static partial Regex StructuredTitleRegex();
    [GeneratedRegex(@"^(?:总体要求|工作要求|主要任务|保障措施|实施步骤|工作目标|工作安排|有关事项|组织领导|注意事项|具体要求)$", RegexOptions.CultureInvariant)]
    private static partial Regex KnownHeadingRegex();
    [GeneratedRegex(@"^(?:根据|按照|为进一步|为落实|现将|经研究|我们|本次|请各|本单位|本公司|依据)", RegexOptions.CultureInvariant)]
    private static partial Regex BodyOpeningRegex();
    [GeneratedRegex(@"^(?:heading|标题)[\s_-]*([1-9])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex HeadingStyleRegex();

    [GeneratedRegex(@"^[一二三四五六七八九十百零〇]+、", RegexOptions.CultureInvariant)]
    private static partial Regex Heading1Regex();
    [GeneratedRegex(@"^[（(][一二三四五六七八九十百零〇]+[）)]", RegexOptions.CultureInvariant)]
    private static partial Regex Heading2Regex();
    [GeneratedRegex(@"^\d{1,3}[.．、](?!\d)", RegexOptions.CultureInvariant)]
    private static partial Regex Heading3Regex();
    [GeneratedRegex(@"^[（(]\d{1,3}[）)]", RegexOptions.CultureInvariant)]
    private static partial Regex Heading4Regex();
    [GeneratedRegex(@"^(\d{4}|[〇零一二三四五六七八九]{4})年(\d{1,2}|[一二三四五六七八九十]{1,3})月(\d{1,2}|[一二三四五六七八九十]{1,3})日[。\s]*$", RegexOptions.CultureInvariant)]
    private static partial Regex DateTextRegex();

    private sealed record Part(string Name, byte[] Bytes, DateTimeOffset LastWriteTime, int ExternalAttributes);
    private sealed record Package(string SourcePath, string Sha256, List<Part> Parts, Dictionary<string, Part> ByName, Dictionary<string, XDocument> XmlParts, XDocument Document);
    private sealed record ParagraphEvidence(bool IsTitleStyle, int? HeadingLevel, string? Alignment, double? FontSize, bool Bold);
    private sealed class FieldInstruction(XElement? simpleElement, bool inInstruction)
    {
        public XElement? SimpleElement { get; } = simpleElement;
        public bool InInstruction { get; set; } = inInstruction;
        public StringBuilder Instruction { get; } = new();
        public StringBuilder Result { get; } = new();
    }
}
