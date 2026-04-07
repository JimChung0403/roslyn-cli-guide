using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;

namespace VbAnalyzer;

public static class CompilationBuilder
{
    /// <summary>
    /// 多 project 編譯策略：
    /// 1. 找出 Form 所在的「主 project」
    /// 2. 其他 project 先各自編譯（各自的 RootNamespace + GlobalImports）
    /// 3. 其他 project 的編譯產物當作 MetadataReference 餵給主 project
    /// 4. 主 project 用自己的 RootNamespace + GlobalImports 編譯
    ///
    /// 這樣每個 project 的 RootNamespace 都正確，跨 project 的型別引用也能解析。
    /// </summary>
    public static (VisualBasicCompilation compilation, int errorCount, List<string> missingTypes) Build(
        List<ProjectInfo> projects, string libsDir, string formName)
    {
        // ── 1. 收集共用 reference assemblies ──

        var sharedReferences = new List<MetadataReference>();

        var refAsmPaths = FindNet48RefAssemblies();
        foreach (var dllPath in refAsmPaths)
            sharedReferences.Add(MetadataReference.CreateFromFile(dllPath));
        Console.Error.WriteLine($"       .NET Framework 4.8 refs: {refAsmPaths.Count} assemblies");

        if (Directory.Exists(libsDir))
        {
            var libDlls = Directory.GetFiles(libsDir, "*.dll", SearchOption.AllDirectories);
            foreach (var dll in libDlls)
            {
                try { sharedReferences.Add(MetadataReference.CreateFromFile(dll)); }
                catch (Exception ex) { Console.Error.WriteLine($"       [WARN] Cannot load DLL {dll}: {ex.Message}"); }
            }
            Console.Error.WriteLine($"       Third-party libs: {libDlls.Length} DLLs from {libsDir}");
        }
        else
        {
            Console.Error.WriteLine($"       [INFO] No libs dir at {libsDir} (third-party types will be unresolved)");
        }

        // ── 2. 找出主 project（包含 Form 的那個） ──

        var mainProject = FindMainProject(projects, formName);
        if (mainProject == null)
        {
            Console.Error.WriteLine($"       [WARN] Cannot determine main project for '{formName}', using first project");
            mainProject = projects[0];
        }
        Console.Error.WriteLine($"       Main project: {mainProject.ProjectName}");

        // ── 3. 編譯非主 project，產出 in-memory DLL 作為 reference ──

        var dependencyReferences = new List<MetadataReference>();

        foreach (var proj in projects)
        {
            if (proj == mainProject) continue;

            var (rootNs, imports) = ParseVbproj(proj.VbprojPath);
            Console.Error.WriteLine($"       Compiling dependency: {proj.ProjectName} (namespace: '{rootNs}', {proj.VbFiles.Count} files)");

            var depTrees = ParseFiles(proj.VbFiles);
            var depCompilation = VisualBasicCompilation.Create(
                assemblyName: proj.ProjectName,
                syntaxTrees: depTrees,
                references: sharedReferences,
                options: new VisualBasicCompilationOptions(
                    OutputKind.DynamicallyLinkedLibrary,
                    optionStrict: OptionStrict.Off,
                    optionExplicit: true,
                    optionInfer: true,
                    rootNamespace: rootNs,
                    globalImports: imports.Select(GlobalImport.Parse)
                )
            );

            // 不需要實際 emit 到檔案，用 in-memory reference
            var depRef = depCompilation.ToMetadataReference();
            dependencyReferences.Add(depRef);

            var depErrors = depCompilation.GetDiagnostics().Count(d => d.Severity == DiagnosticSeverity.Error);
            Console.Error.WriteLine($"         → {depErrors} compilation errors");
        }

        // ── 4. 編譯主 project ──

        var (mainRootNs, mainImports) = ParseVbproj(mainProject.VbprojPath);

        // 合併所有 dependency project 的 RootNamespace 到 global imports
        // 這樣主 project 可以直接用 dependency 的型別短名稱
        foreach (var proj in projects)
        {
            if (proj == mainProject) continue;
            var (depRootNs, _) = ParseVbproj(proj.VbprojPath);
            if (!string.IsNullOrEmpty(depRootNs) && !mainImports.Contains(depRootNs))
            {
                mainImports.Add(depRootNs);
                Console.Error.WriteLine($"       [INFO] Added '{depRootNs}' to global imports (from {proj.ProjectName})");
            }
        }

        Console.Error.WriteLine($"       Main project namespace: '{mainRootNs}'");
        Console.Error.WriteLine($"       Global imports: {mainImports.Count} ({string.Join(", ", mainImports.Take(5))}{(mainImports.Count > 5 ? "..." : "")})");

        // 主 project 的 .vb 包含自己的 + 所有 dependency 的（for iterative expansion 追蹤）
        // 但 dependency 的 .vb 已經透過 in-memory reference 提供型別定義，
        // 這裡還是把 dependency 的 .vb 也放進來，讓 Roslyn 能解析到原始碼位置
        var allVbFiles = new List<string>();
        foreach (var proj in projects)
            allVbFiles.AddRange(proj.VbFiles);

        var mainTrees = ParseFiles(allVbFiles);

        var allReferences = new List<MetadataReference>();
        allReferences.AddRange(sharedReferences);
        allReferences.AddRange(dependencyReferences);

        var compilation = VisualBasicCompilation.Create(
            assemblyName: "VbAnalysis",
            syntaxTrees: mainTrees,
            references: allReferences,
            options: new VisualBasicCompilationOptions(
                OutputKind.DynamicallyLinkedLibrary,
                optionStrict: OptionStrict.Off,
                optionExplicit: true,
                optionInfer: true,
                rootNamespace: mainRootNs,
                globalImports: mainImports.Select(GlobalImport.Parse)
            )
        );

        // ── 5. 收集診斷 ──

        var errors = compilation.GetDiagnostics()
            .Where(d => d.Severity == DiagnosticSeverity.Error)
            .ToList();

        var missingTypes = errors
            .Where(d => d.Id == "BC30002")
            .Select(d =>
            {
                var msg = d.GetMessage();
                var start = msg.IndexOf('\'');
                var end = msg.LastIndexOf('\'');
                return start >= 0 && end > start ? msg[(start + 1)..end] : msg;
            })
            .Distinct().OrderBy(t => t).ToList();

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
    /// 找出包含目標 Form 的 project。
    /// 掃描每個 project 的 .vb 檔案，看哪個含有 Form class 定義。
    /// </summary>
    static ProjectInfo? FindMainProject(List<ProjectInfo> projects, string formName)
    {
        foreach (var proj in projects)
        {
            foreach (var file in proj.VbFiles)
            {
                var fileName = Path.GetFileNameWithoutExtension(file);
                if (string.Equals(fileName, formName, StringComparison.OrdinalIgnoreCase))
                    return proj;
            }
        }
        // fallback: 看哪個 project 有最多 .vb 檔
        return projects.OrderByDescending(p => p.VbFiles.Count).FirstOrDefault();
    }

    static List<SyntaxTree> ParseFiles(List<string> vbFiles)
    {
        var trees = new List<SyntaxTree>();
        foreach (var file in vbFiles)
        {
            try
            {
                var text = File.ReadAllText(file);
                trees.Add(VisualBasicSyntaxTree.ParseText(text, path: file,
                    options: new VisualBasicParseOptions(LanguageVersion.Latest)));
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"       [WARN] Cannot read {file}: {ex.Message}");
            }
        }
        return trees;
    }

    static List<string> FindNet48RefAssemblies()
    {
        var nugetCache = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages");
        var packageDir = Path.Combine(nugetCache, "microsoft.netframework.referenceassemblies.net48");

        if (!Directory.Exists(packageDir))
        {
            Console.Error.WriteLine("       [ERROR] .NET Framework 4.8 reference assemblies not found. Run: dotnet restore");
            return [];
        }

        var versionDir = Directory.GetDirectories(packageDir).OrderByDescending(d => d).FirstOrDefault();
        if (versionDir == null) return [];

        var refDir = Path.Combine(versionDir, "build", ".NETFramework", "v4.8");
        if (!Directory.Exists(refDir)) return [];

        var dlls = new List<string>();
        dlls.AddRange(Directory.GetFiles(refDir, "*.dll"));
        var facadesDir = Path.Combine(refDir, "Facades");
        if (Directory.Exists(facadesDir))
            dlls.AddRange(Directory.GetFiles(facadesDir, "*.dll"));

        return dlls;
    }

    public static (string rootNamespace, List<string> globalImports) ParseVbproj(string? vbprojPath)
    {
        var defaultImports = new List<string>
        {
            "Microsoft.VisualBasic", "System", "System.Collections", "System.Collections.Generic",
            "System.Data", "System.Drawing", "System.Diagnostics", "System.Windows.Forms",
            "System.Linq", "System.Xml.Linq", "System.Threading.Tasks"
        };

        if (vbprojPath == null || !File.Exists(vbprojPath))
            return ("", defaultImports);

        try
        {
            var doc = XDocument.Load(vbprojPath);
            var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
            var rootNs = doc.Descendants(ns + "RootNamespace").FirstOrDefault()?.Value ?? "";
            var imports = doc.Descendants(ns + "Import")
                .Select(e => e.Attribute("Include")?.Value)
                .Where(v => v != null && !v.Contains("$(") && !v.EndsWith(".props") && !v.EndsWith(".targets"))
                .Cast<string>().ToList();

            if (imports.Count == 0) imports = defaultImports;
            return (rootNs, imports);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"       [WARN] Cannot parse .vbproj: {ex.Message}, using defaults");
            return ("", defaultImports);
        }
    }
}
