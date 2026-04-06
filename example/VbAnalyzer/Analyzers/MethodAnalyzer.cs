using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class MethodAnalyzer
{
    public static List<MethodEntry> Analyze(VisualBasicCompilation compilation, string formName, string projectRoot,
        HashSet<string>? additionalTypes = null)
    {
        var results = new List<MethodEntry>();
        var relevantTypes = FindRelevantTypes(compilation, formName);
        if (additionalTypes != null)
            foreach (var t in additionalTypes)
                relevantTypes.Add(t);

        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            var root = tree.GetRoot();
            var relFile = RelPath(tree.FilePath, projectRoot);

            foreach (var methodBlock in root.DescendantNodes().OfType<MethodBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(methodBlock);
                if (symbol == null) continue;
                var ownerName = symbol.ContainingType?.Name ?? "Unknown";
                if (!relevantTypes.Contains(ownerName)) continue;

                var lineSpan = methodBlock.GetLocation().GetLineSpan();
                var returnType = symbol.ReturnsVoid ? "Sub" : "Function";
                var paramStr = string.Join(", ", symbol.Parameters.Select(p => $"{p.Name} As {p.Type?.ToDisplayString() ?? "Object"}"));

                results.Add(new MethodEntry
                {
                    Name = symbol.Name, Owner = ownerName, File = relFile,
                    StartLine = lineSpan.StartLinePosition.Line + 1,
                    EndLine = lineSpan.EndLinePosition.Line + 1,
                    Signature = $"{returnType} {symbol.Name}({paramStr})",
                });
            }

            foreach (var propBlock in root.DescendantNodes().OfType<PropertyBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(propBlock);
                if (symbol == null) continue;
                var ownerName = symbol.ContainingType?.Name ?? "Unknown";
                if (!relevantTypes.Contains(ownerName)) continue;

                var lineSpan = propBlock.GetLocation().GetLineSpan();
                var paramStr = string.Join(", ", symbol.Parameters.Select(p => $"{p.Name} As {p.Type?.ToDisplayString() ?? "Object"}"));

                results.Add(new MethodEntry
                {
                    Name = symbol.Name, Owner = ownerName, File = relFile,
                    StartLine = lineSpan.StartLinePosition.Line + 1,
                    EndLine = lineSpan.EndLinePosition.Line + 1,
                    Signature = $"Property {symbol.Name}({paramStr}) As {symbol.Type?.ToDisplayString() ?? "Object"}",
                });
            }
        }

        return results.OrderBy(m => m.File).ThenBy(m => m.StartLine).ToList();
    }

    static HashSet<string> FindRelevantTypes(VisualBasicCompilation compilation, string formName)
    {
        var types = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { formName };

        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            foreach (var classBlock in tree.GetRoot().DescendantNodes().OfType<ClassBlockSyntax>())
            {
                var classSymbol = model.GetDeclaredSymbol(classBlock);
                if (classSymbol == null || !string.Equals(classSymbol.Name, formName, StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (var inv in classBlock.DescendantNodes().OfType<InvocationExpressionSyntax>())
                {
                    if (model.GetSymbolInfo(inv).Symbol is IMethodSymbol ms && ms.ContainingType?.Name != null)
                        types.Add(ms.ContainingType.Name);
                }

                foreach (var field in classBlock.DescendantNodes().OfType<FieldDeclarationSyntax>())
                    foreach (var decl in field.Declarators)
                        if (decl.AsClause is SimpleAsClauseSyntax ac)
                        {
                            var tn = model.GetTypeInfo(ac.Type).Type?.Name;
                            if (tn != null) types.Add(tn);
                        }

                // MemberAccess（如 modFuncoes.xxx）
                foreach (var ma in classBlock.DescendantNodes().OfType<MemberAccessExpressionSyntax>())
                {
                    var sym = model.GetSymbolInfo(ma).Symbol;
                    if (sym?.ContainingType?.Name != null && !IsExcludedType(sym.ContainingType))
                        types.Add(sym.ContainingType.Name);
                }

                // IdentifierName（如直接引用 Module 成員 CNPJEmpresa、removeAcentos）
                foreach (var id in classBlock.DescendantNodes().OfType<IdentifierNameSyntax>())
                {
                    var sym = model.GetSymbolInfo(id).Symbol;
                    if (sym?.ContainingType?.Name != null
                        && !string.Equals(sym.ContainingType.Name, formName, StringComparison.OrdinalIgnoreCase)
                        && !IsExcludedType(sym.ContainingType))
                        types.Add(sym.ContainingType.Name);
                }
            }
        }
        return types;
    }

    /// <summary>
    /// 排除非程式邏輯的型別：Resources、My.xxx、Framework 內建型別。
    /// 這些型別的方法不應出現在 methods index 中。
    /// </summary>
    static bool IsExcludedType(INamedTypeSymbol type)
    {
        var name = type.Name;
        var ns = type.ContainingNamespace?.ToDisplayString() ?? "";

        // My.Resources.Resources, My.MySettings 等
        if (name == "Resources" || name == "MySettings" || name == "MyProject")
            return true;
        if (ns.Contains("My.Resources") || ns.Contains("My."))
            return true;

        // .NET Framework / System 內建
        if (ns.StartsWith("System.") || ns == "System" || ns.StartsWith("Microsoft."))
            return true;

        return false;
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
