using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class ControlAnalyzer
{
    public static List<ControlEntry> Analyze(VisualBasicCompilation compilation, string formName, string projectRoot)
    {
        var results = new List<ControlEntry>();

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

                foreach (var field in classBlock.DescendantNodes().OfType<FieldDeclarationSyntax>())
                {
                    var hasWithEvents = field.Modifiers.Any(m => m.IsKind(SyntaxKind.WithEventsKeyword));
                    if (!hasWithEvents) continue;

                    foreach (var declarator in field.Declarators)
                    {
                        foreach (var name in declarator.Names)
                        {
                            var controlName = name.Identifier.Text;
                            var typeText = "Unknown";

                            if (declarator.AsClause is SimpleAsClauseSyntax asClause)
                            {
                                var typeInfo = model.GetTypeInfo(asClause.Type);
                                typeText = typeInfo.Type?.ToDisplayString() ?? asClause.Type.ToString();
                            }

                            var declLine = field.GetLocation().GetLineSpan();
                            var declStr = $"{RelPath(tree.FilePath, projectRoot)}:{declLine.StartLinePosition.Line + 1}";

                            var initLoc = FindInitLocation(classBlock, controlName, tree, projectRoot);
                            var parentName = FindParent(classBlock, controlName);

                            var isSpread = typeText.Contains("FpSpread", StringComparison.OrdinalIgnoreCase)
                                        || typeText.Contains("AxFPSpread", StringComparison.OrdinalIgnoreCase)
                                        || typeText.Contains("FarPoint", StringComparison.OrdinalIgnoreCase);

                            results.Add(new ControlEntry
                            {
                                Name = controlName,
                                ControlType = typeText,
                                Declaration = declStr,
                                Initialization = initLoc,
                                Parent = parentName,
                                IsWithEvents = true,
                                IsAxFpSpread = isSpread,
                            });
                        }
                    }
                }
            }
        }

        return results.OrderBy(c => c.Name).ToList();
    }

    static string? FindInitLocation(ClassBlockSyntax classBlock, string controlName, SyntaxTree tree, string projectRoot)
    {
        var initMethod = classBlock.DescendantNodes().OfType<MethodBlockSyntax>()
            .FirstOrDefault(m => string.Equals(
                m.SubOrFunctionStatement.Identifier.Text, "InitializeComponent", StringComparison.OrdinalIgnoreCase));

        if (initMethod == null) return null;

        foreach (var assignment in initMethod.DescendantNodes().OfType<AssignmentStatementSyntax>())
        {
            var leftText = assignment.Left.ToString();
            if ((leftText.Equals($"Me.{controlName}", StringComparison.OrdinalIgnoreCase)
                || leftText.Equals(controlName, StringComparison.OrdinalIgnoreCase))
                && assignment.Right is ObjectCreationExpressionSyntax)
            {
                var line = assignment.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
                return $"{RelPath(tree.FilePath, projectRoot)}:{line}";
            }
        }

        return null;
    }

    static string? FindParent(ClassBlockSyntax classBlock, string controlName)
    {
        var initMethod = classBlock.DescendantNodes().OfType<MethodBlockSyntax>()
            .FirstOrDefault(m => string.Equals(
                m.SubOrFunctionStatement.Identifier.Text, "InitializeComponent", StringComparison.OrdinalIgnoreCase));

        if (initMethod == null) return null;

        foreach (var invocation in initMethod.DescendantNodes().OfType<InvocationExpressionSyntax>())
        {
            var text = invocation.ToString();
            if (text.Contains(".Controls.Add", StringComparison.OrdinalIgnoreCase)
                && text.Contains(controlName, StringComparison.OrdinalIgnoreCase))
            {
                if (invocation.Expression is MemberAccessExpressionSyntax memberAccess
                    && memberAccess.Expression is MemberAccessExpressionSyntax controlsAccess
                    && controlsAccess.Expression is MemberAccessExpressionSyntax parentAccess)
                {
                    return parentAccess.Name.Identifier.Text;
                }
            }
        }

        return null;
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
