using System.Text.RegularExpressions;

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
    /// <summary>
    /// 回傳每個 project 的資訊（vbproj 路徑、.vb 檔案清單），
    /// 支援多 project 各自有不同的 RootNamespace。
    /// </summary>
    public static List<ProjectInfo> CollectProjects(string? slnPath, string? vbprojPath)
    {
        var projects = new List<ProjectInfo>();
        var excludeDirs = new[] { "/obj/", "/bin/", "\\obj\\", "\\bin\\" };

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
                        var projDir = Path.GetDirectoryName(projFullPath)!;
                        var projName = Path.GetFileNameWithoutExtension(projFullPath);

                        var vbFiles = Directory.GetFiles(projDir, "*.vb", SearchOption.AllDirectories)
                            .Where(f => !excludeDirs.Any(ex => f.Contains(ex)))
                            .Select(Path.GetFullPath)
                            .Distinct().OrderBy(f => f).ToList();

                        Console.Error.WriteLine($"       Project: {projRelPath} ({vbFiles.Count} .vb files)");

                        projects.Add(new ProjectInfo
                        {
                            VbprojPath = projFullPath,
                            ProjectDir = projDir,
                            ProjectName = projName,
                            VbFiles = vbFiles
                        });
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

            var projDir = Path.GetDirectoryName(fullPath)!;
            var vbFiles = Directory.GetFiles(projDir, "*.vb", SearchOption.AllDirectories)
                .Where(f => !excludeDirs.Any(ex => f.Contains(ex)))
                .Select(Path.GetFullPath)
                .Distinct().OrderBy(f => f).ToList();

            projects.Add(new ProjectInfo
            {
                VbprojPath = fullPath,
                ProjectDir = projDir,
                ProjectName = Path.GetFileNameWithoutExtension(fullPath),
                VbFiles = vbFiles
            });
        }

        return projects;
    }
}
