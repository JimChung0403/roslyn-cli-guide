using System.Text.RegularExpressions;

namespace VbAnalyzer;

public static class SolutionLoader
{
    /// <summary>
    /// 從 .sln 或 .vbproj 收集所有 .vb 檔案路徑。
    /// 用純文字解析，不依賴 MSBuild（因為 .NET Framework 專案在 Linux 上無法 MSBuild）。
    /// </summary>
    /// <summary>
    /// Returns (vbFiles, firstVbprojPath) — the first .vbproj is used to read global imports.
    /// </summary>
    public static (List<string> vbFiles, string? vbprojPath) CollectVbFiles(string? slnPath, string? vbprojPath)
    {
        var projectDirs = new List<string>();
        string? firstVbproj = null;

        if (slnPath != null)
        {
            var slnFullPath = Path.GetFullPath(slnPath);
            if (!File.Exists(slnFullPath))
            {
                Console.Error.WriteLine($"[ERROR] .sln not found: {slnFullPath}");
                return ([], null);
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
                        firstVbproj ??= projFullPath;
                        var projDir = Path.GetDirectoryName(projFullPath)!;
                        projectDirs.Add(projDir);
                        Console.Error.WriteLine($"       Project: {projRelPath}");
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
                return ([], null);
            }
            firstVbproj = fullPath;
            projectDirs.Add(Path.GetDirectoryName(fullPath)!);
        }

        if (projectDirs.Count == 0)
        {
            Console.Error.WriteLine("[WARN] No project directories found");
            return ([], null);
        }

        // 遞迴收集所有 .vb 檔，排除 obj/ 和 bin/
        var excludeDirs = new[] { "/obj/", "/bin/", "\\obj\\", "\\bin\\" };

        var files = projectDirs
            .SelectMany(dir =>
            {
                if (!Directory.Exists(dir)) return [];
                return Directory.GetFiles(dir, "*.vb", SearchOption.AllDirectories);
            })
            .Where(f => !excludeDirs.Any(ex => f.Contains(ex)))
            .Select(Path.GetFullPath)
            .Distinct()
            .OrderBy(f => f)
            .ToList();

        return (files, firstVbproj);
    }
}
