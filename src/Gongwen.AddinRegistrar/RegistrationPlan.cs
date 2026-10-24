using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Gongwen.AddinRegistrar;

internal enum RegistrationView { Registry32 = 32, Registry64 = 64 }
internal enum PlanValueKind { String, DWord }

internal sealed class RegistrationValue
{
    internal RegistrationValue(string name, object value, PlanValueKind kind = PlanValueKind.String)
    { Name = name; Value = value; Kind = kind; }
    internal string Name { get; }
    internal object Value { get; }
    internal PlanValueKind Kind { get; }
}

internal sealed class RegistrationNode
{
    internal RegistrationNode(string path, params RegistrationValue[] values) { Path = path; Values = values; }
    internal string Path { get; }
    internal IReadOnlyList<RegistrationValue> Values { get; }
}

internal sealed class EnumRegistrationMetadata
{
    internal EnumRegistrationMetadata(string fullName, Guid guid) { FullName = fullName; Guid = guid; }
    internal string FullName { get; }
    internal Guid Guid { get; }
}

internal sealed class AddinRegistrationMetadata
{
    internal AddinRegistrationMetadata(string className, Guid classGuid, string assemblyFullName, Version version,
        string runtimeVersion, string codeBase, IReadOnlyList<EnumRegistrationMetadata> enums)
    {
        ClassName = className; ClassGuid = classGuid; AssemblyFullName = assemblyFullName; Version = version;
        RuntimeVersion = runtimeVersion; CodeBase = codeBase; Enums = enums;
    }
    internal string ClassName { get; }
    internal Guid ClassGuid { get; }
    internal string AssemblyFullName { get; }
    internal Version Version { get; }
    internal string RuntimeVersion { get; }
    internal string CodeBase { get; }
    internal IReadOnlyList<EnumRegistrationMetadata> Enums { get; }
}

internal sealed class RegistrationPlan
{
    internal RegistrationPlan(IReadOnlyList<RegistrationNode> nodes, IReadOnlyList<string> unregisterPaths)
    { Nodes = nodes; UnregisterPaths = unregisterPaths; }
    internal IReadOnlyList<RegistrationNode> Nodes { get; }
    internal IReadOnlyList<string> UnregisterPaths { get; }
}

internal static class RegistrationPlanBuilder
{
    internal const string ProgId = "Gongwen.OfficeAddin.Connect";
    internal const string ClassName = "Gongwen.OfficeAddin.Connect";
    internal const string AssemblyName = "Gongwen.OfficeAddin";
    internal const string ClassesRoot = @"Software\Classes";
    internal const string WordAddins = @"Software\Microsoft\Office\Word\Addins\" + ProgId;
    internal const string WpsAddins = @"Software\Kingsoft\Office\WPS\Addins\" + ProgId;
    internal static readonly Guid ConnectGuid = new Guid("3B873652-4D36-4CE7-B43C-B93B9F1B7A01");
    internal const string ManagedCategoryGuid = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";

    internal static RegistrationPlan Build(AddinRegistrationMetadata metadata)
    {
        Validate(metadata);
        var clsid = ClassesRoot + @"\CLSID\" + GuidKey(metadata.ClassGuid);
        var progId = ClassesRoot + @"\" + ProgId;
        var inproc = clsid + @"\InprocServer32";
        var nodes = new List<RegistrationNode>
        {
            new RegistrationNode(progId, Text("", metadata.ClassName)),
            new RegistrationNode(progId + @"\CLSID", Text("", GuidKey(metadata.ClassGuid))),
            new RegistrationNode(clsid, Text("", metadata.ClassName)),
            new RegistrationNode(inproc, new[] { Text("", "mscoree.dll"), Text("ThreadingModel", "Both") }.Concat(TypeValues(metadata, metadata.ClassName)).ToArray()),
            new RegistrationNode(inproc + @"\" + metadata.Version, TypeValues(metadata, metadata.ClassName)),
            new RegistrationNode(clsid + @"\ProgId", Text("", ProgId)),
            new RegistrationNode(clsid + @"\Implemented Categories\" + ManagedCategoryGuid),
        };
        var unregister = new List<string> { WordAddins, WpsAddins, progId, clsid };
        foreach (var enumeration in metadata.Enums)
        {
            var record = ClassesRoot + @"\Record\" + GuidKey(enumeration.Guid) + @"\" + metadata.Version;
            nodes.Add(new RegistrationNode(record, TypeValues(metadata, enumeration.FullName)));
            // Only this version. Other versions and the shared Record root remain.
            unregister.Add(record);
        }
        foreach (var host in new[] { WordAddins, WpsAddins })
            nodes.Add(new RegistrationNode(host, Text("FriendlyName", "公文一键排版"),
                Text("Description", "本地公文排版，另存新DOCX，原稿保留"), new RegistrationValue("LoadBehavior", 3, PlanValueKind.DWord)));
        return new RegistrationPlan(nodes.AsReadOnly(), unregister.AsReadOnly());
    }

    private static RegistrationValue Text(string name, string value) => new RegistrationValue(name, value);
    private static string GuidKey(Guid guid) => guid.ToString("B").ToUpperInvariant();
    private static RegistrationValue[] TypeValues(AddinRegistrationMetadata metadata, string typeName) => new[]
    {
        Text("Class", typeName), Text("Assembly", metadata.AssemblyFullName),
        Text("RuntimeVersion", metadata.RuntimeVersion), Text("CodeBase", metadata.CodeBase)
    };

    private static void Validate(AddinRegistrationMetadata metadata)
    {
        if (metadata is null || metadata.ClassName != ClassName || metadata.ClassGuid != ConnectGuid ||
            metadata.Version is null || metadata.Version.Revision < 0 || metadata.RuntimeVersion != "v4.0.30319" ||
            string.IsNullOrWhiteSpace(metadata.CodeBase) || !metadata.CodeBase.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ||
            metadata.CodeBase.Any(char.IsControl) || metadata.Enums is null || metadata.Enums.Count != 2)
            throw new ArgumentException("加载项元数据无效。");
        var assembly = new AssemblyName(metadata.AssemblyFullName);
        if (assembly.Name != AssemblyName || assembly.Version != metadata.Version ||
            metadata.Enums.Any(enumeration => enumeration is null || enumeration.Guid == Guid.Empty || enumeration.Guid == ConnectGuid) ||
            metadata.Enums.Select(enumeration => enumeration.Guid).Distinct().Count() != 2 ||
            !metadata.Enums.Select(enumeration => enumeration.FullName).OrderBy(name => name, StringComparer.Ordinal)
                .SequenceEqual(new[] { "Gongwen.OfficeAddin.ExtConnectMode", "Gongwen.OfficeAddin.ExtDisconnectMode" }))
            throw new ArgumentException("加载项元数据不属于本产品。");
    }
}
