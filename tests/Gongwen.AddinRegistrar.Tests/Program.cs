using Gongwen.AddinRegistrar;

var firstEnumGuid = new Guid("A7F05841-8528-41B4-800A-44DFB8C12301");
var secondEnumGuid = new Guid("A7F05841-8528-41B4-800A-44DFB8C12302");
var metadata = Metadata("file:///C:/CI 空格/中文#字符/office-addin/Gongwen.OfficeAddin.dll");
var plan = RegistrationPlanBuilder.Build(metadata);
const string progId = @"Software\Classes\Gongwen.OfficeAddin.Connect";
const string clsid = @"Software\Classes\CLSID\{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}";
var inproc = clsid + @"\InprocServer32";
var passed = 0;
var failed = 0;

Test("cli_accepts_only_exact_register_or_unregister", () =>
{
    Check(RegistrarCommandLine.TryParse(["--register"], out var register) && register == RegistrationOperation.Register);
    Check(RegistrarCommandLine.TryParse(["--unregister"], out var unregister) && unregister == RegistrationOperation.Unregister);
    foreach (var arguments in new[] { Array.Empty<string>(), new[] { "--register", "other.dll" }, new[] { "--register=other.dll" }, new[] { "--REGISTER" }, new[] { "--help" }, new[] { "--unregister", "{foreign-clsid}" } })
        Check(!RegistrarCommandLine.TryParse(arguments, out _));
});

Test("progid_and_clsid_defaults_match_regasm_class_name", () =>
{
    Equal(metadata.ClassName, Value(progId, "").Value);
    Equal("{3B873652-4D36-4CE7-B43C-B93B9F1B7A01}", Value(progId + @"\CLSID", "").Value);
    Equal(metadata.ClassName, Value(clsid, "").Value);
    Equal("Gongwen.OfficeAddin.Connect", Value(clsid + @"\ProgId", "").Value);
    Check(!plan.Nodes.Any(node => node.Values.Any(value => value.Name == "RootClass")));
    Equal(1, Node(clsid).Values.Count);
});

Test("inproc_contains_framework_server_and_all_real_metadata", () =>
{
    Equal("mscoree.dll", Value(inproc, "").Value);
    Equal("Both", Value(inproc, "ThreadingModel").Value);
    Equal(6, Node(inproc).Values.Count);
    AssertMetadata(inproc, metadata.ClassName);
});

Test("version_subkey_mirrors_exact_four_values", () =>
{
    var version = inproc + @"\0.4.0.0";
    Equal(4, Node(version).Values.Count);
    AssertMetadata(version, metadata.ClassName);
    Check(Node(version).Values.All(value => value.Kind == PlanValueKind.String));
});

Test("managed_category_is_an_empty_own_clsid_key", () =>
{
    var category = Node(clsid + @"\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");
    Equal(0, category.Values.Count);
    Check(!plan.Nodes.Any(node => node.Path.Contains(@"\Component Categories\", StringComparison.Ordinal)));
});

Test("both_enum_records_use_runtime_guids_version_and_enum_class", () =>
{
    var records = plan.Nodes.Where(node => node.Path.StartsWith(@"Software\Classes\Record\", StringComparison.Ordinal)).ToArray();
    Equal(2, records.Length);
    foreach (var enumeration in metadata.Enums)
    {
        var record = @"Software\Classes\Record\" + enumeration.Guid.ToString("B").ToUpperInvariant() + @"\0.4.0.0";
        Equal(4, Node(record).Values.Count);
        AssertMetadata(record, enumeration.FullName);
        Check(plan.UnregisterPaths.Contains(record));
        Check(!plan.UnregisterPaths.Contains(record[..record.LastIndexOf('\\')]));
    }
});

Test("word_and_wps_only_use_own_addin_and_dword_load_behavior", () =>
{
    foreach (var host in new[] { RegistrationPlanBuilder.WordAddins, RegistrationPlanBuilder.WpsAddins })
    {
        Equal("公文一键排版", Value(host, "FriendlyName").Value);
        Equal(3, Value(host, "LoadBehavior").Value);
        Equal(PlanValueKind.DWord, Value(host, "LoadBehavior").Kind);
        Equal(3, Node(host).Values.Count);
    }
});

Test("codebase_is_preserved_without_uri_normalization_or_preescaping", () =>
{
    foreach (var codeBase in new[]
    {
        "file:///C:/CI 空格/中文#字符/office-addin/Gongwen.OfficeAddin.dll",
        "file:///C:/CI%20space/%E4%B8%AD%E6%96%87%23/office-addin/Gongwen.OfficeAddin.dll"
    })
    {
        var next = RegistrationPlanBuilder.Build(Metadata(codeBase));
        Equal(4, next.Nodes.SelectMany(node => node.Values).Count(value => value.Name == "CodeBase"));
        Check(next.Nodes.SelectMany(node => node.Values).Where(value => value.Name == "CodeBase").All(value => (string)value.Value == codeBase));
    }
});

Test("register_plan_has_only_fixed_product_scope", () =>
{
    Equal(11, plan.Nodes.Count);
    Equal(6, plan.UnregisterPaths.Count);
    foreach (var node in plan.Nodes)
        Check(node.Path == RegistrationPlanBuilder.WordAddins || node.Path == RegistrationPlanBuilder.WpsAddins ||
            node.Path == progId || node.Path.StartsWith(progId + "\\", StringComparison.Ordinal) ||
            node.Path == clsid || node.Path.StartsWith(clsid + "\\", StringComparison.Ordinal) ||
            metadata.Enums.Any(enumeration => node.Path == @"Software\Classes\Record\" + enumeration.Guid.ToString("B").ToUpperInvariant() + @"\0.4.0.0"));
    Check(!plan.UnregisterPaths.Any(path => path is @"Software\Classes" or @"Software\Classes\CLSID" or @"Software\Classes\Record" || path.Contains("Security", StringComparison.Ordinal)));
});

Test("registration_writes_both_views", () =>
{
    var store = new FakeUserRegistry();
    Equal(0, RegistrationRunner.Execute(RegistrationOperation.Register, metadata, store));
    Equal(22, store.Nodes.Count);
    Equal(11, store.Writes.Count(item => item.View == RegistrationView.Registry32));
    Equal(11, store.Writes.Count(item => item.View == RegistrationView.Registry64));
    Equal(0, store.Deletes.Count);
});

Test("second_view_failure_rolls_back_only_product_keys_in_both_views", () =>
{
    var store = new FakeUserRegistry { FailWriteView = RegistrationView.Registry64 };
    var foreign = @"Software\Classes\CLSID\{C03C1B07-FD87-4215-9A6C-F08B89599100}";
    var otherVersion = @"Software\Classes\Record\" + firstEnumGuid.ToString("B").ToUpperInvariant() + @"\0.3.0.0";
    foreach (var view in Views()) { store.Seed(view, foreign); store.Seed(view, otherVersion); }
    Equal(2, RegistrationRunner.Execute(RegistrationOperation.Register, metadata, store));
    Equal(12, store.Deletes.Count);
    foreach (var view in Views()) { Check(store.Exists(view, foreign)); Check(store.Exists(view, otherVersion)); }
    Check(store.Nodes.Keys.All(key => key.EndsWith(foreign, StringComparison.Ordinal) || key.EndsWith(otherVersion, StringComparison.Ordinal)));
});

Test("unregister_keeps_foreign_software_security_and_other_record_versions", () =>
{
    var store = new FakeUserRegistry();
    Equal(0, RegistrationRunner.Execute(RegistrationOperation.Register, metadata, store));
    var neighbors = new[]
    {
        @"Software\Classes\Another.Product", @"Software\Microsoft\Office\Word\Security",
        @"Software\Microsoft\Office\Word\Addins\Another.Product",
        @"Software\Classes\Record\" + firstEnumGuid.ToString("B").ToUpperInvariant() + @"\0.3.0.0",
        @"Software\Classes\Record\" + firstEnumGuid.ToString("B").ToUpperInvariant() + @"\0.4.0.01",
        @"Software\Classes\Component Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    };
    foreach (var view in Views()) foreach (var neighbor in neighbors) store.Seed(view, neighbor);
    Equal(0, RegistrationRunner.Execute(RegistrationOperation.Unregister, metadata, store));
    Equal(neighbors.Length * 2, store.Nodes.Count);
    foreach (var view in Views()) foreach (var neighbor in neighbors) Check(store.Exists(view, neighbor));
});

Test("unregister_missing_keys_is_idempotent", () =>
{
    var store = new FakeUserRegistry();
    Equal(0, RegistrationRunner.Execute(RegistrationOperation.Unregister, metadata, store));
    Equal(0, RegistrationRunner.Execute(RegistrationOperation.Unregister, metadata, store));
    Equal(0, store.Nodes.Count);
});

Test("unregister_failure_returns_two_and_attempts_all_owned_paths", () =>
{
    var store = new FakeUserRegistry { FailDeletePath = progId };
    Equal(2, RegistrationRunner.Execute(RegistrationOperation.Unregister, metadata, store));
    Equal(12, store.Deletes.Count);
    Check(store.Deletes.All(item => plan.UnregisterPaths.Contains(item.Path)));
});

Test("foreign_or_invalid_metadata_cannot_start_registry_mutations", () =>
{
    var bad = new[]
    {
        new AddinRegistrationMetadata("Foreign.Connect", metadata.ClassGuid, metadata.AssemblyFullName, metadata.Version, metadata.RuntimeVersion, metadata.CodeBase, metadata.Enums),
        new AddinRegistrationMetadata(metadata.ClassName, Guid.NewGuid(), metadata.AssemblyFullName, metadata.Version, metadata.RuntimeVersion, metadata.CodeBase, metadata.Enums),
        new AddinRegistrationMetadata(metadata.ClassName, metadata.ClassGuid, "Foreign, Version=0.4.0.0, Culture=neutral, PublicKeyToken=null", metadata.Version, metadata.RuntimeVersion, metadata.CodeBase, metadata.Enums),
        new AddinRegistrationMetadata(metadata.ClassName, metadata.ClassGuid, metadata.AssemblyFullName, metadata.Version, metadata.RuntimeVersion, "https://example.invalid/foreign.dll", metadata.Enums),
        new AddinRegistrationMetadata(metadata.ClassName, metadata.ClassGuid, metadata.AssemblyFullName, metadata.Version, metadata.RuntimeVersion, metadata.CodeBase,
            [new EnumRegistrationMetadata("Foreign.Enum", firstEnumGuid), new EnumRegistrationMetadata("Gongwen.OfficeAddin.ExtDisconnectMode", secondEnumGuid)])
    };
    foreach (var invalid in bad)
    {
        var store = new FakeUserRegistry();
        Equal(2, RegistrationRunner.Execute(RegistrationOperation.Register, invalid, store));
        Equal(0, store.Writes.Count);
        Equal(0, store.Deletes.Count);
    }
});

Test("invalid_operation_has_no_registry_side_effect", () =>
{
    var store = new FakeUserRegistry();
    Equal(2, RegistrationRunner.Execute(RegistrationOperation.Invalid, metadata, store));
    Equal(0, store.Writes.Count);
    Equal(0, store.Deletes.Count);
});

Console.WriteLine($"RESULT {passed} passed, {failed} failed");
return failed == 0 ? 0 : 1;

AddinRegistrationMetadata Metadata(string codeBase) => new("Gongwen.OfficeAddin.Connect", RegistrationPlanBuilder.ConnectGuid,
    "Gongwen.OfficeAddin, Version=0.4.0.0, Culture=neutral, PublicKeyToken=null", new Version(0, 4, 0, 0), "v4.0.30319", codeBase,
    [new EnumRegistrationMetadata("Gongwen.OfficeAddin.ExtConnectMode", firstEnumGuid), new EnumRegistrationMetadata("Gongwen.OfficeAddin.ExtDisconnectMode", secondEnumGuid)]);

RegistrationNode Node(string path) => plan.Nodes.Single(node => node.Path == path);
RegistrationValue Value(string path, string name) => Node(path).Values.Single(value => value.Name == name);
void AssertMetadata(string path, string className)
{
    Equal(className, Value(path, "Class").Value);
    Equal(metadata.AssemblyFullName, Value(path, "Assembly").Value);
    Equal(metadata.RuntimeVersion, Value(path, "RuntimeVersion").Value);
    Equal(metadata.CodeBase, Value(path, "CodeBase").Value);
}
static RegistrationView[] Views() => [RegistrationView.Registry32, RegistrationView.Registry64];
static void Check(bool condition) { if (!condition) throw new Exception("Contract failed."); }
static void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception("Unexpected value."); }
void Test(string name, Action action)
{
    try { action(); passed++; Console.WriteLine("PASS " + name); }
    catch (Exception error) { failed++; Console.WriteLine("FAIL " + name + " (" + error.GetType().Name + ")"); }
}

internal sealed class FakeUserRegistry : IUserRegistry
{
    internal Dictionary<string, RegistrationNode> Nodes { get; } = new(StringComparer.OrdinalIgnoreCase);
    internal List<(RegistrationView View, string Path)> Writes { get; } = [];
    internal List<(RegistrationView View, string Path)> Deletes { get; } = [];
    internal RegistrationView? FailWriteView { get; set; }
    internal string? FailDeletePath { get; set; }
    internal void Seed(RegistrationView view, string path) => Nodes[Key(view, path)] = new RegistrationNode(path);
    internal bool Exists(RegistrationView view, string path) => Nodes.ContainsKey(Key(view, path));
    public void CreateOrUpdate(RegistrationView view, RegistrationNode node)
    {
        Writes.Add((view, node.Path));
        if (view == FailWriteView) throw new UnauthorizedAccessException("Fictional access failure.");
        Nodes[Key(view, node.Path)] = node;
    }
    public void DeleteOwnedTree(RegistrationView view, string path)
    {
        Deletes.Add((view, path));
        if (path == FailDeletePath) throw new UnauthorizedAccessException("Fictional access failure.");
        var key = Key(view, path);
        foreach (var existing in Nodes.Keys.Where(existing => existing.Equals(key, StringComparison.OrdinalIgnoreCase) || existing.StartsWith(key + "\\", StringComparison.OrdinalIgnoreCase)).ToArray()) Nodes.Remove(existing);
    }
    private static string Key(RegistrationView view, string path) => (int)view + ":" + path;
}
