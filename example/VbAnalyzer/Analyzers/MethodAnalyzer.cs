using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class MethodAnalyzer
{
    public static List<MethodEntry> Analyze(VisualBasicCompilation compilation, string formName, string projectRoot)
    {
        var results = new List<MethodEntry>();
        var relevantTypes = FindRelevantTypes(compilation, formName);

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

                // 組裝 signature，與 Python 的 signature 欄位格式一致
                var returnType = symbol.ReturnsVoid ? "Sub" : "Function";
                var paramStr = string.Join(", ", symbol.Parameters.Select(p =>
                    $"{p.Name} As {p.Type?.ToDisplayString() ?? "Object"}"));
                var signature = $"{returnType} {symbol.Name}({paramStr})";

                results.Add(new MethodEntry
                {
                    Name = symbol.Name,
                    Owner = ownerName,
                    File = relFile,
                    StartLine = lineSpan.StartLinePosition.Line + 1,
                    EndLine = lineSpan.EndLinePosition.Line + 1,
                    Signature = signature,
                });
            }

            foreach (var propBlock in root.DescendantNodes().OfType<PropertyBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(propBlock);
                if (symbol == null) continue;

                var ownerName = symbol.ContainingType?.Name ?? "Unknown";
                if (!relevantTypes.Contains(ownerName)) continue;

                var lineSpan = propBlock.GetLocation().GetLineSpan();
                var paramStr = string.Join(", ", symbol.Parameters.Select(p =>
                    $"{p.Name} As {p.Type?.ToDisplayString() ?? "Object"}"));
                var signature = $"Property {symbol.Name}({paramStr}) As {symbol.Type?.ToDisplayString() ?? "Object"}";

                results.Add(new MethodEntry
                {
                    Name = symbol.Name,
                    Owner = ownerName,
                    File = relFile,
                    StartLine = lineSpan.StartLinePosition.Line + 1,
                    EndLine = lineSpan.EndLinePosition.Line + 1,
                    Signature = signature,
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
                    if (symbolInfo.Symbol is IMethodSymbol methodSymbol)
                    {
                        var targetType = methodSymbol.ContainingType?.Name;
                        if (targetType != null) types.Add(targetType);
                    }
                }

                foreach (var field in classBlock.DescendantNodes().OfType<FieldDeclarationSyntax>())
                {
                    foreach (var declarator in field.Declarators)
                    {
                        if (declarator.AsClause is SimpleAsClauseSyntax asClause)
                        {
                            var typeInfo = model.GetTypeInfo(asClause.Type);
                            var typeName = typeInfo.Type?.Name;
                            if (typeName != null) types.Add(typeName);
                        }
                    }
                }
            }
        }

        return types;
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
