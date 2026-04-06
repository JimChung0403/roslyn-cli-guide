using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;

namespace VbAnalyzer;

public static class CompilationBuilder
{
    public static (VisualBasicCompilation compilation, int errorCount, List<string> missingTypes) Build(
        List<string> vbFiles, string libsDir, string? vbprojPath = null)
    {
        var syntaxTrees = new List<SyntaxTree>();
        foreach (var file in vbFiles)
        {
            try
            {
                var text = File.ReadAllText(file);
                var tree = VisualBasicSyntaxTree.ParseText(text, path: file,
                    options: new VisualBasicParseOptions(LanguageVersion.Latest));
                syntaxTrees.Add(tree);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"       [WARN] Cannot read {file}: {ex.Message}");
            }
        }

        var references = new List<MetadataReference>();

        var refAsmPaths = FindNet48RefAssemblies();
        foreach (var dllPath in refAsmPaths)
            references.Add(MetadataReference.CreateFromFile(dllPath));
        Console.Error.WriteLine($"       .NET Framework 4.8 refs: {refAsmPaths.Count} assemblies");

        if (Directory.Exists(libsDir))
        {
            var libDlls = Directory.GetFiles(libsDir, "*.dll", SearchOption.AllDirectories);
            foreach (var dll in libDlls)
            {
                try { references.Add(MetadataReference.CreateFromFile(dll)); }
                catch (Exception ex) { Console.Error.WriteLine($"       [WARN] Cannot load DLL {dll}: {ex.Message}"); }
            }
            Console.Error.WriteLine($"       Third-party libs: {libDlls.Length} DLLs from {libsDir}");
        }
        else
        {
            Console.Error.WriteLine($"       [INFO] No libs dir at {libsDir} (third-party types will be unresolved)");
        }

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

    static (string rootNamespace, List<string> globalImports) ParseVbproj(string? vbprojPath)
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
