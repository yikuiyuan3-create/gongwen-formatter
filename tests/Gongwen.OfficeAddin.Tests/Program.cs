using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Gongwen.OfficeAddin;

// Behavioral tests run without Office, COM activation or starting a process.
// Temporary source documents contain only explicitly fictional text.
var root = Path.Combine(Path.GetTempPath(), "gongwen-office-addin-test-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
var appDirectory = Path.Combine(root, "app-0.4.0");
var addinDirectory = Path.Combine(appDirectory, "office-addin");
var sourceDirectory = Path.Combine(root, "中文 办公文档");
Directory.CreateDirectory(addinDirectory);
Directory.CreateDirectory(sourceDirectory);
var assemblyLocation = Path.Combine(addinDirectory, "Gongwen.OfficeAddin.dll");
var fixedExecutable = Path.Combine(appDirectory, LaunchPolicy.ExecutableName);
File.WriteAllText(assemblyLocation, "Fictional metadata fixture; never loaded.");
File.WriteAllText(fixedExecutable, "Fictional executable fixture; never executed.");
var source = Path.Combine(sourceDirectory, "关于虚构测试工作的通知.docx");
CreateDocument(source);
var failed = 0;
var passed = 0;

try
{
    Test("saved_docx_plans_fixed_parent_executable", () =>
    {
        var plan = LaunchPolicy.ForCurrentDocument(Host(source), assemblyLocation);
        Equal(fixedExecutable, plan.ExecutablePath);
        Sequence(["--open", source], ParseWindowsCommandLine(plan.Arguments));
        var start = plan.ToProcessStartInfo();
        Check(!start.UseShellExecute && start.FileName == fixedExecutable, "Must use fixed executable without shell.");
        Equal(appDirectory, start.WorkingDirectory);
    });

    Test("unsaved_changes_rejected_without_saving_host", () =>
    {
        var document = new FakeDocument(source, sourceDirectory, saved: false);
        var application = new FakeApplication(document);
        Reject(() => LaunchPolicy.ForCurrentDocument(new ReflectionDocumentHost(application), assemblyLocation), "Ctrl+S");
        Check(document.SaveCount == 0 && document.SaveAsCount == 0, "Must not save the host.");
    });

    Test("missing_saved_path_and_new_document_rejected", () =>
    {
        Reject(() => LaunchPolicy.ForCurrentDocument(new FakeHost(new("", "", true)), assemblyLocation), "Ctrl+S");
        Reject(() => LaunchPolicy.ForCurrentDocument(new FakeHost(new(source, "", true)), assemblyLocation), "Ctrl+S");
        Reject(() => LaunchPolicy.ForCurrentDocument(new FakeHost(new("Document1", "", false)), assemblyLocation), "DOCX");
    });

    Test("old_word_extension_requires_docx_save_as", () =>
    {
        var legacy = Path.Combine(sourceDirectory, "虚构旧文档.doc");
        CreateDocument(legacy);
        Reject(() => LaunchPolicy.ForCurrentDocument(Host(legacy), assemblyLocation), "另存为 DOCX");
    });

    Test("missing_directory_empty_and_inaccessible_sources_rejected", () =>
    {
        Reject(() => LaunchPolicy.ForCurrentDocument(Host(Path.Combine(sourceDirectory, "已移走.docx")), assemblyLocation));
        var folder = Path.Combine(sourceDirectory, "不是文件.docx");
        Directory.CreateDirectory(folder);
        Reject(() => LaunchPolicy.ForCurrentDocument(Host(folder), assemblyLocation));
        var empty = Path.Combine(sourceDirectory, "尚未保存完整.docx");
        File.WriteAllBytes(empty, []);
        Reject(() => LaunchPolicy.ForCurrentDocument(Host(empty), assemblyLocation));
    });

    Test("mismatched_host_path_rejected", () =>
        Reject(() => LaunchPolicy.ForCurrentDocument(new FakeHost(new(source, addinDirectory, true)), assemblyLocation)));

    Test("source_read_only_and_no_host_save_or_save_as", () =>
    {
        var hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source)));
        var document = new FakeDocument(source, sourceDirectory, true);
        var application = new FakeApplication(document);
        var plan = LaunchPolicy.ForCurrentDocument(new ReflectionDocumentHost(application), assemblyLocation);
        Sequence(["--open", source], ParseWindowsCommandLine(plan.Arguments));
        Check(document.SaveCount == 0 && document.SaveAsCount == 0 && application.ActiveReads == 1, "Host read-only boundary failed.");
        Equal(hash, Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(source))));
    });

    Test("existing_office_writer_handle_supported", () =>
    {
        using var writer = new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete);
        var plan = LaunchPolicy.ForCurrentDocument(Host(source), assemblyLocation);
        Check(plan.Arguments.StartsWith("--open ", StringComparison.Ordinal), "Saved source must remain readable with a writer handle.");
    });

    Test("exclusive_unreadable_docx_is_a_controlled_rejection", () =>
    {
        using var exclusive = new FileStream(source, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        Reject(() => LaunchPolicy.ForCurrentDocument(Host(source), assemblyLocation));
    });

    Test("host_com_failure_has_no_raw_path_or_body_in_message", () =>
    {
        var error = Reject(() => LaunchPolicy.ForCurrentDocument(new ReflectionDocumentHost(new ThrowingApplication()), assemblyLocation));
        Check(!error.Message.Contains("FICTIONAL-PRIVATE") && error.InnerException is null, "Host exception leaked raw details.");
    });

    Test("no_active_document_is_a_controlled_error", () =>
        Reject(() => LaunchPolicy.ForCurrentDocument(new ReflectionDocumentHost(new FakeApplication(null)), assemblyLocation)));

    Test("missing_deployment_does_not_search_document_folder_or_path", () =>
    {
        var missingAddinDirectory = Path.Combine(root, "missing-app", "office-addin");
        Directory.CreateDirectory(missingAddinDirectory);
        File.WriteAllText(Path.Combine(sourceDirectory, LaunchPolicy.ExecutableName), "Never executed.");
        Reject(() => LaunchPolicy.ForCurrentDocument(Host(source), Path.Combine(missingAddinDirectory, "Gongwen.OfficeAddin.dll")));
    });

    Test("open_tool_requires_no_document_and_no_arguments", () =>
    {
        var plan = LaunchPolicy.ForTool(assemblyLocation);
        Equal(fixedExecutable, plan.ExecutablePath);
        Equal("", plan.Arguments);
        Check(!plan.ToProcessStartInfo().UseShellExecute, "Open tool must not use a shell.");
    });

    Test("document_shell_metacharacters_cannot_change_executable_or_switches", () =>
    {
        var injected = Path.Combine(sourceDirectory, "虚构 & echo UNEXECUTED; $HOME --other.docx");
        CreateDocument(injected);
        var plan = LaunchPolicy.ForCurrentDocument(Host(injected), assemblyLocation);
        Equal(fixedExecutable, plan.ExecutablePath);
        Sequence(["--open", injected], ParseWindowsCommandLine(plan.Arguments));
        Check(!plan.ToProcessStartInfo().UseShellExecute, "Payload must remain data.");
    });

    Test("http_and_relative_document_locations_rejected_without_fetch", () =>
    {
        Reject(() => LaunchPolicy.ForCurrentDocument(new FakeHost(new("https://example.invalid/private.docx", "https://example.invalid", true)), assemblyLocation));
        Reject(() => LaunchPolicy.ForCurrentDocument(new FakeHost(new("relative.docx", ".", true)), assemblyLocation));
    });

    Test("windows_quoting_known_empty_space_quote_and_backslash_vectors", () =>
    {
        Equal("\"\"", WindowsCommandLine.QuoteArgument(""));
        Equal("\"plain\"", WindowsCommandLine.QuoteArgument("plain"));
        Equal("\"two words\"", WindowsCommandLine.QuoteArgument("two words"));
        Equal("\"a\\\"b\"", WindowsCommandLine.QuoteArgument("a\"b"));
        Equal("\"C:\\中文\\目录\\\\\"", WindowsCommandLine.QuoteArgument("C:\\中文\\目录\\"));
        foreach (var value in new[]
        {
            @"C:\办公 文件\关于测试工作的通知.docx", @"\\server\共享\文件.docx",
            "C:\\a\\\\\"b.docx", "trailing\\\\", "\" --other \"", "中文; & | %NAME% $() `data`"
        }) Sequence(["--open", value], ParseWindowsCommandLine("--open " + WindowsCommandLine.QuoteArgument(value)));
    });

    Test("windows_quoting_deterministic_adversarial_round_trips", () =>
    {
        var random = new Random(1024);
        const string alphabet = "abc中文 空格\\\"&;|()%$-";
        for (var i = 0; i < 1000; i++)
        {
            var value = new string(Enumerable.Range(0, random.Next(0, 100)).Select(_ => alphabet[random.Next(alphabet.Length)]).ToArray());
            Sequence(["--open", value], ParseWindowsCommandLine("--open " + WindowsCommandLine.QuoteArgument(value)));
        }
        try { WindowsCommandLine.QuoteArgument("invalid\0argument"); }
        catch (ArgumentException) { return; }
        throw new Exception("NUL must be rejected.");
    });

    Test("ribbon_xml_contains_exact_local_callbacks_and_no_remote_assets", () =>
    {
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        using var reader = XmlReader.Create(new StringReader(RibbonDefinition.Xml), settings);
        var xml = XDocument.Load(reader);
        XNamespace ui = "http://schemas.microsoft.com/office/2006/01/customui";
        Equal(ui + "customUI", xml.Root!.Name);
        Equal("公文排版", (string?)xml.Descendants(ui + "tab").Single().Attribute("label"));
        var buttons = xml.Descendants(ui + "button").ToArray();
        Sequence(["排版当前文档", "打开排版工具", "使用说明"], buttons.Select(button => (string)button.Attribute("label")!).ToArray());
        var callbacks = typeof(IGongwenRibbonCallbacks).GetMethods().Select(method => method.Name).Order().ToArray();
        Sequence(callbacks, buttons.Select(button => (string)button.Attribute("onAction")!).Order().ToArray());
        Check(!xml.Descendants().Attributes().Any(attribute => attribute.Name.LocalName is "src" or "href" or "image" or "getImage"), "Ribbon must not fetch remote content.");
        Check(!xml.Descendants().Any(element => element.Name.LocalName is "script" or "webView"), "Ribbon must remain declarative.");
    });

    Test("com_interface_guids_dispatch_ids_and_interface_kinds_are_stable", () =>
    {
        Equal(new Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), typeof(IDTExtensibility2).GUID);
        Equal(ComInterfaceType.InterfaceIsDual, typeof(IDTExtensibility2).GetCustomAttribute<InterfaceTypeAttribute>()!.Value);
        Sequence([1, 2, 3, 4, 5], typeof(IDTExtensibility2).GetMethods().Select(method => method.GetCustomAttribute<DispIdAttribute>()!.Value).Order().ToArray());
        Equal(new Guid("000C0396-0000-0000-C000-000000000046"), typeof(IRibbonExtensibility).GUID);
        // Office's original typelib marks this interface FDual (TypeLibType 4160).
        Equal(ComInterfaceType.InterfaceIsDual, typeof(IRibbonExtensibility).GetCustomAttribute<InterfaceTypeAttribute>()!.Value);
        Equal(new Guid("C8B34047-5D6E-4A52-9D46-85B523D6F2D1"), typeof(IGongwenRibbonCallbacks).GUID);
        Equal(ComInterfaceType.InterfaceIsIDispatch, typeof(IGongwenRibbonCallbacks).GetCustomAttribute<InterfaceTypeAttribute>()!.Value);
        Sequence([101, 102, 103], typeof(IGongwenRibbonCallbacks).GetMethods().Select(method => method.GetCustomAttribute<DispIdAttribute>()!.Value).Order().ToArray());
    });
}
finally
{
    try { Directory.Delete(root, recursive: true); }
    catch (IOException) { }
    catch (UnauthorizedAccessException) { }
}

Console.WriteLine($"RESULT {passed} passed, {failed} failed");
return failed == 0 ? 0 : 1;

void Test(string name, Action action)
{
    try { action(); passed++; Console.WriteLine("PASS " + name); }
    catch (Exception error) { failed++; Console.WriteLine("FAIL " + name + " (" + error.GetType().Name + ")"); }
}

static void Check(bool value, string message) { if (!value) throw new Exception(message); }
static void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception("Unexpected value."); }
static void Sequence<T>(T[] expected, T[] actual) { if (!expected.SequenceEqual(actual)) throw new Exception("Unexpected argument or contract sequence."); }
static FakeHost Host(string source) => new(new ActiveDocumentSnapshot(source, Path.GetDirectoryName(source)!, true));
static AddinOperationException Reject(Action action, string? message = null)
{
    try { action(); }
    catch (AddinOperationException error)
    {
        if (message is not null) Check(error.Message.Contains(message, StringComparison.Ordinal), "Controlled error should explain recovery.");
        return error;
    }
    throw new Exception("Expected controlled rejection.");
}

static void CreateDocument(string path)
{
    using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
    var content = zip.CreateEntry("[Content_Types].xml");
    using (var writer = new StreamWriter(content.Open(), new UTF8Encoding(false))) writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
    var relationships = zip.CreateEntry("_rels/.rels");
    using (var writer = new StreamWriter(relationships.Open(), new UTF8Encoding(false))) writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
    var main = zip.CreateEntry("word/document.xml");
    using (var writer = new StreamWriter(main.Open(), new UTF8Encoding(false))) writer.Write("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>虚构测试内容，不含真实公文或个人数据。</w:t></w:r></w:p></w:body></w:document>");
}

// Independent command-line consumer, used only for asserting the actual argv
// boundary. It does not call QuoteArgument or launch a child process.
static string[] ParseWindowsCommandLine(string commandLine)
{
    var arguments = new List<string>();
    var index = 0;
    while (index < commandLine.Length)
    {
        while (index < commandLine.Length && commandLine[index] is ' ' or '\t') index++;
        if (index == commandLine.Length) break;
        var argument = new StringBuilder();
        var quoted = false;
        while (index < commandLine.Length && (quoted || commandLine[index] is not (' ' or '\t')))
        {
            var count = 0;
            while (index < commandLine.Length && commandLine[index] == '\\') { count++; index++; }
            if (index < commandLine.Length && commandLine[index] == '"')
            {
                argument.Append('\\', count / 2);
                if (count % 2 == 1) argument.Append('"'); else quoted = !quoted;
                index++;
            }
            else
            {
                argument.Append('\\', count);
                if (index < commandLine.Length && (quoted || commandLine[index] is not (' ' or '\t'))) argument.Append(commandLine[index++]);
            }
        }
        arguments.Add(argument.ToString());
    }
    return arguments.ToArray();
}

internal sealed class FakeHost(ActiveDocumentSnapshot snapshot) : IActiveDocumentHost
{
    public ActiveDocumentSnapshot ReadActiveDocument() => snapshot;
}

public sealed class FakeApplication(FakeDocument? document)
{
    public int ActiveReads { get; private set; }
    public FakeDocument? ActiveDocument { get { ActiveReads++; return document; } }
}

public sealed class FakeDocument(string fullName, string path, bool saved)
{
    public string FullName => fullName;
    public string Path => path;
    public bool Saved => saved;
    public int SaveCount { get; private set; }
    public int SaveAsCount { get; private set; }
    public void Save() { SaveCount++; throw new Exception("Host must remain read-only."); }
    public void SaveAs(string target) { SaveAsCount++; throw new Exception("Host must remain read-only."); }
}

public sealed class ThrowingApplication
{
    public object ActiveDocument => throw new COMException("FICTIONAL-PRIVATE-PATH-AND-BODY must not escape.");
}
