using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class ReferenceAnalyzer
{
    public static List<ReferenceEntry> Analyze(VisualBasicCompilation compilation, string formName, string projectRoot)
    {
        var results = new List<ReferenceEntry>();

        foreach (var tree in compilation.SyntaxTrees)
        {
            var model = compilation.GetSemanticModel(tree);
            var root = tree.GetRoot();
            var relFile = RelPath(tree.FilePath, projectRoot);

            foreach (var classBlock in root.DescendantNodes().OfType<ClassBlockSyntax>())
            {
                var classSymbol = model.GetDeclaredSymbol(classBlock);
                if (classSymbol == null) continue;
                if (!string.Equals(classSymbol.Name, formName, StringComparison.OrdinalIgnoreCase))
                    continue;

                AnalyzeInvocations(classBlock, model, tree, relFile, projectRoot, results);
                AnalyzeMemberAccess(classBlock, model, tree, relFile, projectRoot, results);
            }
        }

        return results.OrderBy(r => r.File).ThenBy(r => r.Line).ToList();
    }

    static void AnalyzeInvocations(ClassBlockSyntax classBlock, SemanticModel model,
        SyntaxTree tree, string relFile, string projectRoot, List<ReferenceEntry> results)
    {
        foreach (var invocation in classBlock.DescendantNodes().OfType<InvocationExpressionSyntax>())
        {
            var lineSpan = invocation.GetLocation().GetLineSpan();
            var line = lineSpan.StartLinePosition.Line + 1;
            var callerMethod = GetContainingMethodName(invocation);
            var context = invocation.ToString();
            // 截斷過長的 context
            if (context.Length > 120) context = context[..120] + "...";

            var symbolInfo = model.GetSymbolInfo(invocation);

            if (symbolInfo.Symbol is IMethodSymbol target)
            {
                var targetLoc = target.Locations.FirstOrDefault();
                string? resolvedTo = null;

                if (targetLoc != null && targetLoc.IsInSource)
                {
                    var targetSpan = targetLoc.GetLineSpan();
                    resolvedTo = $"{RelPath(targetSpan.Path, projectRoot)}:{targetSpan.StartLinePosition.Line + 1}";
                }

                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod,
                    Target = $"{target.ContainingType?.Name ?? "?"}.{target.Name}",
                    File = relFile,
                    Line = line,
                    RefType = ClassifyMethodCall(target),
                    Context = context,
                    ResolvedTo = resolvedTo
                });
            }
            else
            {
                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod,
                    Target = invocation.Expression.ToString(),
                    File = relFile,
                    Line = line,
                    RefType = "unresolved",
                    Context = context,
                    ResolvedTo = null
                });
            }
        }
    }

    static void AnalyzeMemberAccess(ClassBlockSyntax classBlock, SemanticModel model,
        SyntaxTree tree, string relFile, string projectRoot, List<ReferenceEntry> results)
    {
        foreach (var memberAccess in classBlock.DescendantNodes().OfType<MemberAccessExpressionSyntax>())
        {
            if (memberAccess.Parent is InvocationExpressionSyntax) continue;

            var symbolInfo = model.GetSymbolInfo(memberAccess);
            if (symbolInfo.Symbol is not IPropertySymbol prop) continue;

            var ownerType = prop.ContainingType;
            if (ownerType == null) continue;
            if (!IsControlType(ownerType)) continue;

            var isWrite = IsLeftSideOfAssignment(memberAccess);
            var callerMethod = GetContainingMethodName(memberAccess);
            var line = memberAccess.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            var context = memberAccess.Parent?.ToString() ?? memberAccess.ToString();
            if (context.Length > 120) context = context[..120] + "...";

            results.Add(new ReferenceEntry
            {
                Caller = callerMethod,
                Target = memberAccess.ToString(),
                File = relFile,
                Line = line,
                RefType = isWrite ? "control-write" : "control-read",
                Context = context,
                ResolvedTo = null  // control property 不指向原始碼位置
            });
        }
    }

    static string ClassifyMethodCall(IMethodSymbol target)
    {
        var typeName = target.ContainingType?.Name ?? "";

        if (typeName.EndsWith("Service") || typeName.EndsWith("Client")
            || typeName.EndsWith("Proxy") || typeName.EndsWith("SoapClient"))
            return "remote-call";

        if (target.Name is "Show" or "ShowDialog"
            && InheritsFrom(target.ContainingType, "Form"))
            return "dialog-navigation";

        if (typeName.Contains("Spread") || typeName.Contains("FpSpread")
            || typeName.Contains("SheetView"))
        {
            if (target.Name.Contains("Sort", StringComparison.OrdinalIgnoreCase))
                return "spread-sort";
            if (target.Name.Contains("Export", StringComparison.OrdinalIgnoreCase)
                || target.Name.Contains("Save", StringComparison.OrdinalIgnoreCase))
                return "spread-export";
        }

        return "method-call";
    }

    static bool IsControlType(INamedTypeSymbol type)
    {
        var current = type;
        while (current != null)
        {
            if (current.Name == "Control"
                && current.ContainingNamespace?.ToDisplayString() == "System.Windows.Forms")
                return true;
            if (current.Name == "AxHost") return true;
            current = current.BaseType;
        }
        return false;
    }

    static bool InheritsFrom(INamedTypeSymbol? type, string baseName)
    {
        var current = type;
        while (current != null)
        {
            if (current.Name == baseName) return true;
            current = current.BaseType;
        }
        return false;
    }

    static bool IsLeftSideOfAssignment(SyntaxNode node)
    {
        if (node.Parent is AssignmentStatementSyntax assignment)
            return assignment.Left == node;
        return false;
    }

    static string GetContainingMethodName(SyntaxNode node)
    {
        var current = node.Parent;
        while (current != null)
        {
            if (current is MethodBlockSyntax method)
                return method.SubOrFunctionStatement.Identifier.Text;
            if (current is PropertyBlockSyntax prop)
                return prop.PropertyStatement.Identifier.Text;
            current = current.Parent;
        }
        return "(class-level)";
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
