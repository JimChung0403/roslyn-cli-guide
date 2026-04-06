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
        // 找 Form 自身的檔案（main + designer + partial）
        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            foreach (var classBlock in tree.GetRoot().DescendantNodes().OfType<ClassBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(classBlock);
                if (symbol != null && string.Equals(symbol.Name, formName, StringComparison.OrdinalIgnoreCase))
                    formFiles.Add(tree.FilePath);
            }
        }

        foreach (var filePath in formFiles)
        {
            var relPath = RelPath(filePath, projectRoot);
            var isDesigner = filePath.EndsWith(".Designer.vb", StringComparison.OrdinalIgnoreCase);
            var isMain = string.Equals(System.IO.Path.GetFileNameWithoutExtension(filePath), formName, StringComparison.OrdinalIgnoreCase);

            // 對齊 Python 的 role 名稱
            var role = isDesigner ? "form-designer" : isMain ? "form-main" : "partial";
            var reason = isDesigner ? "Form Designer 檔案" : isMain ? "Form 主檔案" : "Partial class 檔案";

            results.Add(new FileEntry { Path = relPath, Role = role, Reason = reason });
        }

        // 收集 Form 引用的所有外部型別（方法呼叫、欄位型別、member access、Imports）
        var referencedTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            foreach (var classBlock in tree.GetRoot().DescendantNodes().OfType<ClassBlockSyntax>())
            {
                var cs = model.GetDeclaredSymbol(classBlock);
                if (cs == null || !string.Equals(cs.Name, formName, StringComparison.OrdinalIgnoreCase)) continue;

                // 方法呼叫的 ContainingType
                foreach (var inv in classBlock.DescendantNodes().OfType<InvocationExpressionSyntax>())
                {
                    if (model.GetSymbolInfo(inv).Symbol is IMethodSymbol m && m.ContainingType?.Name != null)
                        referencedTypes.Add(m.ContainingType.Name);
                }

                // 欄位型別
                foreach (var field in classBlock.DescendantNodes().OfType<FieldDeclarationSyntax>())
                    foreach (var decl in field.Declarators)
                        if (decl.AsClause is SimpleAsClauseSyntax ac)
                        {
                            var tn = model.GetTypeInfo(ac.Type).Type?.Name;
                            if (tn != null) referencedTypes.Add(tn);
                        }

                // MemberAccess（如 CNPJEmpresa、modFuncoes.xxx）
                foreach (var ma in classBlock.DescendantNodes().OfType<MemberAccessExpressionSyntax>())
                {
                    var sym = model.GetSymbolInfo(ma).Symbol;
                    if (sym?.ContainingType?.Name != null && !IsExcludedType(sym.ContainingType))
                        referencedTypes.Add(sym.ContainingType.Name);
                }

                // IdentifierName（如直接引用 Module 成員 CNPJEmpresa）
                foreach (var id in classBlock.DescendantNodes().OfType<IdentifierNameSyntax>())
                {
                    var sym = model.GetSymbolInfo(id).Symbol;
                    if (sym?.ContainingType?.Name != null
                        && !string.Equals(sym.ContainingType.Name, formName, StringComparison.OrdinalIgnoreCase)
                        && !IsExcludedType(sym.ContainingType))
                        referencedTypes.Add(sym.ContainingType.Name);
                }
            }
        }

        // 找出包含被引用型別的檔案
        foreach (var tree in compilation.SyntaxTrees)
        {
            if (formFiles.Contains(tree.FilePath)) continue;
            var model = compilation.GetSemanticModel(tree);
            var root = tree.GetRoot();

            var hasRelevant = root.DescendantNodes().OfType<ClassBlockSyntax>()
                .Any(c => { var s = model.GetDeclaredSymbol(c); return s != null && referencedTypes.Contains(s.Name); })
                || root.DescendantNodes().OfType<ModuleBlockSyntax>()
                .Any(m => { var s = model.GetDeclaredSymbol(m); return s != null && referencedTypes.Contains(s.Name); });

            if (hasRelevant)
            {
                var relPath = RelPath(tree.FilePath, projectRoot);
                // 找出這個檔案裡被引用的 symbol 名稱
                var symbolNames = new List<string>();
                foreach (var c in root.DescendantNodes().OfType<ClassBlockSyntax>())
                {
                    var s = model.GetDeclaredSymbol(c);
                    if (s != null && referencedTypes.Contains(s.Name)) symbolNames.Add(s.Name);
                }
                foreach (var m in root.DescendantNodes().OfType<ModuleBlockSyntax>())
                {
                    var s = model.GetDeclaredSymbol(m);
                    if (s != null && referencedTypes.Contains(s.Name)) symbolNames.Add(s.Name);
                }

                results.Add(new FileEntry
                {
                    Path = relPath,
                    Role = "related-helper",
                    Reason = $"檔名含相關 symbol `{string.Join("`, `", symbolNames)}`"
                });
            }
        }

        // 排序對齊 Python：form-main → form-designer → partial → related-helper
        return results
            .OrderBy(f => f.Role switch { "form-main" => 0, "form-designer" => 1, "partial" => 2, "related-helper" => 3, _ => 4 })
            .ThenBy(f => f.Path)
            .ToList();
    }

    static bool IsExcludedType(INamedTypeSymbol type)
    {
        var name = type.Name;
        var ns = type.ContainingNamespace?.ToDisplayString() ?? "";
        if (name is "Resources" or "MySettings" or "MyProject") return true;
        if (ns.Contains("My.Resources") || ns.Contains("My.")) return true;
        if (ns.StartsWith("System.") || ns == "System" || ns.StartsWith("Microsoft.")) return true;
        return false;
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
