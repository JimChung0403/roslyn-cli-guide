using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class EventAnalyzer
{
    public static List<EventEntry> Analyze(VisualBasicCompilation compilation, string formName, string projectRoot)
    {
        var results = new List<EventEntry>();

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

                AnalyzeHandlesClauses(classBlock, tree, projectRoot, results);
                AnalyzeAddHandlers(classBlock, tree, projectRoot, results);
            }
        }

        return results.OrderBy(e => e.Control).ThenBy(e => e.EventType).ToList();
    }

    static void AnalyzeHandlesClauses(ClassBlockSyntax classBlock, SyntaxTree tree,
        string projectRoot, List<EventEntry> results)
    {
        foreach (var methodBlock in classBlock.DescendantNodes().OfType<MethodBlockSyntax>())
        {
            var methodStmt = methodBlock.SubOrFunctionStatement;
            var handlesClause = methodStmt.HandlesClause;
            if (handlesClause == null) continue;

            var handlerName = methodStmt.Identifier.Text;
            var handlerLine = methodBlock.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            var defStr = $"{RelPath(tree.FilePath, projectRoot)}:{handlerLine}";

            foreach (var eventItem in handlesClause.Events)
            {
                var controlName = eventItem.EventContainer?.ToString() ?? "Unknown";
                var eventName = eventItem.EventMember.Identifier.Text;

                if (controlName.StartsWith("Me.", StringComparison.OrdinalIgnoreCase))
                    controlName = controlName[3..];

                var wireupLine = handlesClause.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
                var wireupStr = $"{RelPath(tree.FilePath, projectRoot)}:{wireupLine}";

                results.Add(new EventEntry
                {
                    Handler = handlerName,
                    Control = controlName,
                    EventType = eventName,
                    Definition = defStr,
                    Wireups = [wireupStr]
                });
            }
        }
    }

    static void AnalyzeAddHandlers(ClassBlockSyntax classBlock, SyntaxTree tree,
        string projectRoot, List<EventEntry> results)
    {
        foreach (var addHandler in classBlock.DescendantNodes().OfType<AddRemoveHandlerStatementSyntax>())
        {
            if (!addHandler.IsKind(SyntaxKind.AddHandlerStatement)) continue;

            var line = addHandler.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            var wireupStr = $"{RelPath(tree.FilePath, projectRoot)}:{line}";

            var controlName = "Unknown";
            var eventName = "Unknown";

            if (addHandler.EventExpression is MemberAccessExpressionSyntax memberAccess)
            {
                controlName = memberAccess.Expression.ToString();
                eventName = memberAccess.Name.Identifier.Text;
                if (controlName.StartsWith("Me.", StringComparison.OrdinalIgnoreCase))
                    controlName = controlName[3..];
            }

            var handlerName = addHandler.DelegateExpression.ToString();
            if (handlerName.StartsWith("AddressOf ", StringComparison.OrdinalIgnoreCase))
                handlerName = handlerName[10..].Trim();

            results.Add(new EventEntry
            {
                Handler = handlerName,
                Control = controlName,
                EventType = eventName,
                Definition = wireupStr,
                Wireups = [wireupStr]
            });
        }
    }

    static string RelPath(string fullPath, string root)
    {
        if (string.IsNullOrEmpty(root)) return fullPath;
        try { return System.IO.Path.GetRelativePath(root, fullPath); }
        catch { return fullPath; }
    }
}
