using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace VbAnalyzer;

public record ProjectInfo
{
    public string VbprojPath { get; init; } = "";
    public string ProjectDir { get; init; } = "";
    public string ProjectName { get; init; } = "";
    public List<string> VbFiles { get; init; } = [];
}

public static class SolutionLoader
{
    public static List<ProjectInfo> CollectProjects(string? slnPath, string? vbprojPath)
    {
        var projects = new List<ProjectInfo>();

        if (slnPath != null)
        {
            var slnFullPath = Path.GetFullPath(slnPath);
            if (!File.Exists(slnFullPath))
            {
                Console.Error.WriteLine($"[ERROR] .sln not found: {slnFullPath}");
                return [];
            }

            var slnDir = Path.GetDirectoryName(slnFullPath)!;
            var projPattern = new Regex(@"""([^""]+\.vbproj)""", RegexOptions.IgnoreCase);

            foreach (var line in File.ReadLines(slnFullPath))
            {
                var match = projPattern.Match(line);
                if (match.Success)
                {
                    var projRelPath = match.Groups[1].Value.Replace('\\', '/');
                    var projFullPath = Path.GetFullPath(Path.Combine(slnDir, projRelPath));

                    if (File.Exists(projFullPath))
                    {
                        var info = CollectProjectFiles(projFullPath);
                        projects.Add(info);
                    }
                    else
                    {
                        Console.Error.WriteLine($"       [WARN] Project not found: {projFullPath}");
                    }
                }
            }
        }
        else if (vbprojPath != null)
        {
            var fullPath = Path.GetFullPath(vbprojPath);
            if (!File.Exists(fullPath))
            {
                Console.Error.WriteLine($"[ERROR] .vbproj not found: {fullPath}");
                return [];
            }
            projects.Add(CollectProjectFiles(fullPath));
        }

        return projects;
    }

    /// <summary>
    /// 收集一個 project 的所有 .vb 檔案。兩種來源：
    /// 1. 目錄遞迴掃描（project 目錄下所有 .vb）
    /// 2. .vbproj 裡的 Compile Include 外部引用（如 ..\..\Common\Model\clsString.vb）
    ///
    /// Visual Studio 支援用 Compile Include + Link 引用 project 目錄外的 .vb 檔：
    ///   <Compile Include="..\..\Common\Model\clsString.vb">
    ///       <Link>Model\clsString.vb</Link>
    ///   </Compile>
    /// 這些檔案物理位置在外部目錄，但編譯時屬於這個 project。
    /// </summary>
    static ProjectInfo CollectProjectFiles(string vbprojFullPath)
    {
        var projDir = Path.GetDirectoryName(vbprojFullPath)!;
        var projName = Path.GetFileNameWithoutExtension(vbprojFullPath);
        var excludeDirs = new[] { "/obj/", "/bin/", "\\obj\\", "\\bin\\" };

        // 1. 目錄遞迴掃描
        var vbFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in Directory.GetFiles(projDir, "*.vb", SearchOption.AllDirectories))
        {
            if (!excludeDirs.Any(ex => f.Contains(ex)))
                vbFiles.Add(Path.GetFullPath(f));
        }

        var dirFileCount = vbFiles.Count;

        // 2. 解析 .vbproj 的 <Compile Include="..."> 找外部引用
        var linkedFiles = 0;
        try
        {
            var doc = XDocument.Load(vbprojFullPath);
            var ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;

            foreach (var compile in doc.Descendants(ns + "Compile"))
            {
                var include = compile.Attribute("Include")?.Value;
                if (include == null) continue;

                // 把 Windows 路徑轉成當前平台路徑
                var relativePath = include.Replace('\\', Path.DirectorySeparatorChar);
                var fullPath = Path.GetFullPath(Path.Combine(projDir, relativePath));

                // 只處理 .vb 檔，且在 project 目錄外的（目錄內的已經被遞迴掃描到了）
                if (fullPath.EndsWith(".vb", StringComparison.OrdinalIgnoreCase)
                    && !fullPath.StartsWith(projDir, StringComparison.OrdinalIgnoreCase)
                    && File.Exists(fullPath)
                    && !vbFiles.Contains(fullPath))
                {
                    vbFiles.Add(fullPath);
                    linkedFiles++;
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"       [WARN] Cannot parse {vbprojFullPath} for Compile Include: {ex.Message}");
        }

        var sortedFiles = vbFiles.OrderBy(f => f).ToList();

        Console.Error.WriteLine($"       Project: {projName} ({dirFileCount} dir files + {linkedFiles} linked files = {sortedFiles.Count} total)");

        return new ProjectInfo
        {
            VbprojPath = vbprojFullPath,
            ProjectDir = projDir,
            ProjectName = projName,
            VbFiles = sortedFiles
        };
    }
}
