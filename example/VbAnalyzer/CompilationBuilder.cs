using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;

namespace VbAnalyzer;

public static class CompilationBuilder
{
    /// <summary>
    /// 把所有 .vb 檔 + reference assemblies 組成一個 Compilation。
    /// 用 AdhocWorkspace 概念（手動組裝），不走 MSBuild。
    /// Partial class 在這一步自動合併。
    /// </summary>
    public static (VisualBasicCompilation compilation, int errorCount, List<string> missingTypes) Build(
        List<string> vbFiles, string libsDir, string? vbprojPath = null)
    {
        // ── 1. 解析所有 .vb 為 SyntaxTree ──

        var syntaxTrees = new List<SyntaxTree>();
        foreach (var file in vbFiles)
        {
            try
            {
                var text = File.ReadAllText(file);
                var tree = VisualBasicSyntaxTree.ParseText(
                    text,
                    path: file,
                    options: new VisualBasicParseOptions(LanguageVersion.Latest)
                );
                syntaxTrees.Add(tree);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"       [WARN] Cannot read {file}: {ex.Message}");
            }
        }

        // ── 2. 收集 reference assemblies ──

        var references = new List<MetadataReference>();

        // 2a. .NET Framework 4.8 reference assemblies（NuGet 套件提供）
        //     安裝 Microsoft.NETFramework.ReferenceAssemblies.net48 後，
        //     DLL 位於 NuGet cache 的 build/.NETFramework/v4.8/ 下
        var refAsmPaths = FindNet48RefAssemblies();
        foreach (var dllPath in refAsmPaths)
        {
            references.Add(MetadataReference.CreateFromFile(dllPath));
        }
        Console.Error.WriteLine($"       .NET Framework 4.8 refs: {refAsmPaths.Count} assemblies");

        // 2b. 第三方 DLL（從 Windows 複製過來的）
        if (Directory.Exists(libsDir))
        {
            var libDlls = Directory.GetFiles(libsDir, "*.dll", SearchOption.AllDirectories);
            foreach (var dll in libDlls)
            {
                try
                {
                    references.Add(MetadataReference.CreateFromFile(dll));
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"       [WARN] Cannot load DLL {dll}: {ex.Message}");
                }
            }
            Console.Error.WriteLine($"       Third-party libs: {libDlls.Length} DLLs from {libsDir}");
        }
        else
        {
            Console.Error.WriteLine($"       [INFO] No libs dir at {libsDir} (third-party types will be unresolved)");
        }

        // ── 3. 從 .vbproj 讀取 global imports 和 rootNamespace ──

        var (rootNamespace, globalImports) = ParseVbproj(vbprojPath);
        Console.Error.WriteLine($"       Root namespace: '{rootNamespace}'");
        Console.Error.WriteLine($"       Global imports: {globalImports.Count} ({string.Join(", ", globalImports.Take(5))}{(globalImports.Count > 5 ? "..." : "")})");

        var compilation = VisualBasicCompilation.Create(
            assemblyName: "VbAnalysis",
            syntaxTrees: syntaxTrees,
            references: references,
            options: new VisualBasicCompilationOptions(
                OutputKind.DynamicallyLinkedLibrary,
                optionStrict: OptionStrict.Off,
                optionExplicit: true,
                optionInfer: true,
                rootNamespace: rootNamespace,
                globalImports: globalImports.Select(GlobalImport.Parse)
            )
        );

        // ── 4. 收集診斷資訊 ──

        var errors = compilation.GetDiagnostics()
            .Where(d => d.Severity == DiagnosticSeverity.Error)
            .ToList();

        // 找出缺失的型別（BC30002 = Type 'xxx' is not defined）
        var missingTypes = errors
            .Where(d => d.Id == "BC30002")
            .Select(d =>
            {
                // 從錯誤訊息提取型別名稱
                var msg = d.GetMessage();
                var start = msg.IndexOf('\'');
                var end = msg.LastIndexOf('\'');
                return start >= 0 && end > start ? msg[(start + 1)..end] : msg;
            })
            .Distinct()
            .OrderBy(t => t)
            .ToList();

        if (missingTypes.Count > 0)
        {
            Console.Error.WriteLine($"       Missing types (top 10):");
            foreach (var t in missingTypes.Take(10))
                Console.Error.WriteLine($"         - {t}");
            if (missingTypes.Count > 10)
                Console.Error.WriteLine($"         ... and {missingTypes.Count - 10} more (see stats.json)");
        }

        return (compilation, errors.Count, missingTypes);
    }

    /// <summary>
    /// 找到 NuGet cache 中 .NET Framework 4.8 reference assemblies 的位置。
    /// 套件 Microsoft.NETFramework.ReferenceAssemblies.net48 安裝後，
    /// DLL 在 ~/.nuget/packages/microsoft.netframework.referenceassemblies.net48/1.0.3/build/.NETFramework/v4.8/
    /// </summary>
    static List<string> FindNet48RefAssemblies()
    {
        var nugetCache = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".nuget", "packages"
        );

        var packageDir = Path.Combine(nugetCache, "microsoft.netframework.referenceassemblies.net48");

        if (!Directory.Exists(packageDir))
        {
            Console.Error.WriteLine("       [ERROR] .NET Framework 4.8 reference assemblies not found.");
            Console.Error.WriteLine("       Run: dotnet restore");
            return [];
        }

        // 找最新版本
        var versionDir = Directory.GetDirectories(packageDir)
            .OrderByDescending(d => d)
            .FirstOrDefault();

        if (versionDir == null) return [];

        var refDir = Path.Combine(versionDir, "build", ".NETFramework", "v4.8");
        if (!Directory.Exists(refDir))
        {
            // 嘗試不同路徑結構
            refDir = Path.Combine(versionDir, "build", ".NETFramework", "v4.8", "Facades");
            if (!Directory.Exists(refDir))
            {
                Console.Error.WriteLine($"       [ERROR] Cannot find ref assemblies in {versionDir}");
                return [];
            }
        }

        // 收集主目錄 + Facades 子目錄的所有 DLL
        var dlls = new List<string>();
        dlls.AddRange(Directory.GetFiles(refDir, "*.dll"));

        var facadesDir = Path.Combine(refDir, "Facades");
        if (Directory.Exists(facadesDir))
        {
            dlls.AddRange(Directory.GetFiles(facadesDir, "*.dll"));
        }

        return dlls;
    }

    /// <summary>
    /// 從 .vbproj 讀取 RootNamespace 和 Global Imports。
    /// VB.NET 專案預設有 global imports（System, System.Data 等），
    /// 不設定的話 Roslyn 無法解析 DataTable、DateTime 等短名稱。
    /// </summary>
    static (string rootNamespace, List<string> globalImports) ParseVbproj(string? vbprojPath)
    {
        // VB.NET 標準 global imports（即使讀不到 .vbproj 也用這些）
        var defaultImports = new List<string>
        {
            "Microsoft.VisualBasic",
            "System",
            "System.Collections",
            "System.Collections.Generic",
            "System.Data",
            "System.Drawing",
            "System.Diagnostics",
            "System.Windows.Forms",
            "System.Linq",
            "System.Xml.Linq",
            "System.Threading.Tasks"
        };

        if (vbprojPath == null || !File.Exists(vbprojPath))
            return ("", defaultImports);

        try
        {
            var doc = XDocument.Load(vbprojPath);
            var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

            // 讀 RootNamespace
            var rootNs = doc.Descendants(ns + "RootNamespace").FirstOrDefault()?.Value ?? "";

            // 讀 <Import Include="System.Data" /> 等
            var imports = doc.Descendants(ns + "Import")
                .Select(e => e.Attribute("Include")?.Value)
                .Where(v => v != null && !v.Contains("$(") && !v.EndsWith(".props") && !v.EndsWith(".targets"))
                .Cast<string>()
                .ToList();

            if (imports.Count == 0)
                imports = defaultImports;

            return (rootNs, imports);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"       [WARN] Cannot parse .vbproj: {ex.Message}, using defaults");
            return ("", defaultImports);
        }
    }
}
