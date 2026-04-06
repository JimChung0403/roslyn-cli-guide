using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class FileAnalyzer
{
    public static List<FileEntry> Analyze(VisualBasicCompilation compilation, string formName, string projectRoot)
    {
        var results = new List<FileEntry>();
        var formFiles = new HashSet<string>();

        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            var root = tree.GetRoot();

            foreach (var classBlock in root.DescendantNodes().OfType<ClassBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(classBlock);
                if (symbol == null) continue;
                if (string.Equals(symbol.Name, formName, StringComparison.OrdinalIgnoreCase))
                    formFiles.Add(tree.FilePath);
            }
        }

        foreach (var filePath in formFiles)
        {
            var relPath = RelPath(filePath, projectRoot);
            var role = ClassifyRole(filePath, formName);
            var reason = role switch
            {
                "main" => "Form main code file",
                "designer" => "Designer-generated partial class",
                _ => "Partial class file"
            };

            results.Add(new FileEntry
            {
                Path = relPath,
                Role = role,
                Reason = reason
            });
        }

        // 加入 Form 引用的其他檔案（helper、service 等）
        foreach (var tree in compilation.SyntaxTrees)
        {
            if (formFiles.Contains(tree.FilePath)) continue;

            var model = compilation.GetSemanticModel(tree);
            var root = tree.GetRoot();
            var hasRelevantClass = false;

            foreach (var classBlock in root.DescendantNodes().OfType<ClassBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(classBlock);
                if (symbol == null) continue;

                // 檢查 Form 是否有引用這個 class
                if (IsReferencedByForm(compilation, formName, symbol.Name))
                {
                    hasRelevantClass = true;
                    break;
                }
            }

            foreach (var moduleBlock in root.DescendantNodes().OfType<ModuleBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(moduleBlock);
                if (symbol == null) continue;
                if (IsReferencedByForm(compilation, formName, symbol.Name))
                {
                    hasRelevantClass = true;
                    break;
                }
            }

            if (hasRelevantClass)
            {
                results.Add(new FileEntry
                {
                    Path = RelPath(tree.FilePath, projectRoot),
                    Role = "resolved-dependency",
                    Reason = "referenced by Form methods"
                });
            }
        }

        return results.OrderBy(f => f.Role switch
        {
            "main" => 0, "designer" => 1, "partial" => 2, _ => 3
        }).ToList();
    }

    static bool IsReferencedByForm(VisualBasicCompilation compilation, string formName, string typeName)
    {
        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            var root = tree.GetRoot();

            foreach (var classBlock in root.DescendantNodes().OfType<ClassBlockSyntax>())
            {
                var classSymbol = model.GetDeclaredSymbol(classBlock);
                if (classSymbol == null) continue;
                if (!string.Equals(classSymbol.Name, formName, StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (var invocation in classBlock.DescendantNodes().OfType<InvocationExpressionSyntax>())
                {
                    var symbolInfo = model.GetSymbolInfo(invocation);
                    if (symbolInfo.Symbol is IMethodSymbol method
                        && string.Equals(method.ContainingType?.Name, typeName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
        }
        return false;
    }

    static string ClassifyRole(string filePath, string formName)
    {
        if (filePath.EndsWith(".Designer.vb", StringComparison.OrdinalIgnoreCase))
            return "designer";

        var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
        if (string.Equals(fileName, formName, StringComparison.OrdinalIgnoreCase))
            return "main";

        return "partial";
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
