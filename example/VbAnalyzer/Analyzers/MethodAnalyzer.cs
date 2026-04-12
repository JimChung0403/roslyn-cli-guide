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
                // 直接從 syntax tree 讀原始宣告行，保留原始碼的型別短名/修飾詞，跟 Python 對齊
                // 去掉 attribute（<...> _）前綴，只保留 Sub/Function 宣告本身
                var sigRaw = methodBlock.SubOrFunctionStatement.ToString();
                var sigLines = sigRaw.Split('\n').Select(l => l.Trim()).Where(l => !l.StartsWith("<") && l.Length > 0);
                var sig = string.Join(" ", sigLines).Trim();

                results.Add(new MethodEntry
                {
                    Name = symbol.Name, Owner = ownerName, File = relFile,
                    StartLine = lineSpan.StartLinePosition.Line + 1,
                    EndLine = lineSpan.EndLinePosition.Line + 1,
                    Signature = sig,
                });
            }

            foreach (var propBlock in root.DescendantNodes().OfType<PropertyBlockSyntax>())
            {
                var symbol = model.GetDeclaredSymbol(propBlock);
                if (symbol == null) continue;
                var ownerName = symbol.ContainingType?.Name ?? "Unknown";
                if (!relevantTypes.Contains(ownerName)) continue;

                var lineSpan = propBlock.GetLocation().GetLineSpan();
                var propSig = propBlock.PropertyStatement.ToString().Trim();

                results.Add(new MethodEntry
                {
                    Name = symbol.Name, Owner = ownerName, File = relFile,
                    StartLine = lineSpan.StartLinePosition.Line + 1,
                    EndLine = lineSpan.EndLinePosition.Line + 1,
                    Signature = propSig,
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

    /// <summary>
    /// 產生跟 Python 一致的 VB.NET method signature。
    /// 包含 Public/Private/Friend + Sub/Function + ByVal/ByRef + 回傳型別。
    /// </summary>
    static string FormatMethodSignature(IMethodSymbol symbol)
    {
        var access = FormatAccessibility(symbol.DeclaredAccessibility);
        var overrides = symbol.IsOverride ? "Overrides " : "";
        var kind = symbol.ReturnsVoid ? "Sub" : "Function";

        var paramParts = symbol.Parameters.Select(p =>
        {
            var refKind = p.RefKind == RefKind.Ref ? "ByRef " : "ByVal ";
            return $"{refKind}{p.Name} As {p.Type?.ToDisplayString() ?? "Object"}";
        });
        var paramStr = string.Join(", ", paramParts);

        var returnPart = symbol.ReturnsVoid ? "" : $" As {symbol.ReturnType?.ToDisplayString() ?? "Object"}";
        return $"{access}{overrides}{kind} {symbol.Name}({paramStr}){returnPart}";
    }

    static string FormatAccessibility(Accessibility accessibility)
    {
        return accessibility switch
        {
            Accessibility.Public => "Public ",
            Accessibility.Private => "Private ",
            Accessibility.Protected => "Protected ",
            Accessibility.Friend => "Friend ",
            Accessibility.ProtectedOrFriend => "Protected Friend ",
            _ => ""
        };
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
