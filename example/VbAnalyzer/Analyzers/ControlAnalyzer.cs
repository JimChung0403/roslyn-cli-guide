using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class ControlAnalyzer
{
    // 與 Python 相同的 regex，從 InitializeComponent 提取 Location/Size/Text
    static readonly Regex LocationRe = new(@"\bMe\.([A-Za-z_]\w*)\.Location\s*=\s*New\s+System\.Drawing\.Point\((\d+),\s*(\d+)\)", RegexOptions.IgnoreCase);
    static readonly Regex SizeRe = new(@"\bMe\.([A-Za-z_]\w*)\.Size\s*=\s*New\s+System\.Drawing\.Size\((\d+),\s*(\d+)\)", RegexOptions.IgnoreCase);
    static readonly Regex TextRe = new(@"\bMe\.([A-Za-z_]\w*)\.(Text|Caption)\s*=\s*""([^""]*)""", RegexOptions.IgnoreCase);
    static readonly string[] DefaultProps = ["Text", "Visible", "Enabled", "ReadOnly", "Checked", "SelectedIndex", "SelectedValue", "Dock", "Anchor", "TabIndex", "TabStop"];

    public static List<ControlEntry> Analyze(VisualBasicCompilation compilation, string formName, string projectRoot)
    {
        var results = new List<ControlEntry>();
        // 從 InitializeComponent 的原始碼文字提取 layout 資訊
        var locations = new Dictionary<string, (int x, int y)>(StringComparer.OrdinalIgnoreCase);
        var sizes = new Dictionary<string, (int w, int h)>(StringComparer.OrdinalIgnoreCase);
        var displayTexts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var defaultProperties = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

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

                // 從 InitializeComponent 原始碼提取 layout
                ExtractLayoutFromInitializeComponent(classBlock, tree, projectRoot, locations, sizes, displayTexts, defaultProperties);

                // 找 WithEvents 欄位宣告
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

                            locations.TryGetValue(controlName, out var loc);
                            sizes.TryGetValue(controlName, out var sz);
                            displayTexts.TryGetValue(controlName, out var text);
                            defaultProperties.TryGetValue(controlName, out var props);

                            results.Add(new ControlEntry
                            {
                                Name = controlName,
                                ControlType = typeText,
                                Declaration = declStr,
                                Initialization = initLoc,
                                Parent = parentName,
                                DefaultProperties = props ?? [],
                                DisplayText = text,
                                LocationX = locations.ContainsKey(controlName) ? loc.x : null,
                                LocationY = locations.ContainsKey(controlName) ? loc.y : null,
                                SizeW = sizes.ContainsKey(controlName) ? sz.w : null,
                                SizeH = sizes.ContainsKey(controlName) ? sz.h : null,
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

    static void ExtractLayoutFromInitializeComponent(ClassBlockSyntax classBlock, SyntaxTree tree,
        string projectRoot,
        Dictionary<string, (int x, int y)> locations,
        Dictionary<string, (int w, int h)> sizes,
        Dictionary<string, string> displayTexts,
        Dictionary<string, List<string>> defaultProperties)
    {
        var initMethod = classBlock.DescendantNodes().OfType<MethodBlockSyntax>()
            .FirstOrDefault(m => string.Equals(
                m.SubOrFunctionStatement.Identifier.Text, "InitializeComponent", StringComparison.OrdinalIgnoreCase));

        if (initMethod == null) return;

        var relFile = RelPath(tree.FilePath, projectRoot);
        var lines = initMethod.ToString().Split('\n');
        var baseLineNumber = initMethod.GetLocation().GetLineSpan().StartLinePosition.Line;

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            var lineNum = baseLineNumber + i + 1;

            var locMatch = LocationRe.Match(line);
            if (locMatch.Success)
                locations[locMatch.Groups[1].Value] = (int.Parse(locMatch.Groups[2].Value), int.Parse(locMatch.Groups[3].Value));

            var sizeMatch = SizeRe.Match(line);
            if (sizeMatch.Success)
                sizes[sizeMatch.Groups[1].Value] = (int.Parse(sizeMatch.Groups[2].Value), int.Parse(sizeMatch.Groups[3].Value));

            var textMatch = TextRe.Match(line);
            if (textMatch.Success && !displayTexts.ContainsKey(textMatch.Groups[1].Value))
                displayTexts[textMatch.Groups[1].Value] = textMatch.Groups[3].Value;

            foreach (var prop in DefaultProps)
            {
                var propMatch = Regex.Match(line, $@"\bMe\.([A-Za-z_]\w*)\.{prop}\s*=", RegexOptions.IgnoreCase);
                if (propMatch.Success)
                {
                    var ctrlName = propMatch.Groups[1].Value;
                    if (!defaultProperties.ContainsKey(ctrlName))
                        defaultProperties[ctrlName] = [];
                    defaultProperties[ctrlName].Add($"{relFile}:{lineNum}");
                }
            }
        }
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
            if (!text.Contains(".Controls.Add", StringComparison.OrdinalIgnoreCase))
                continue;

            // 精確 match：argument 裡的控制項名稱必須是 Me.{controlName}
            var args = invocation.ArgumentList?.Arguments;
            if (args == null || args.Value.Count == 0) continue;
            var argText = args.Value[0].ToString().Trim();
            if (!argText.Equals($"Me.{controlName}", StringComparison.OrdinalIgnoreCase)
                && !argText.Equals(controlName, StringComparison.OrdinalIgnoreCase))
                continue;

            // 從 Me.{parent}.Controls.Add(...) 提取 parent
            if (invocation.Expression is MemberAccessExpressionSyntax memberAccess
                && memberAccess.Expression is MemberAccessExpressionSyntax controlsAccess
                && controlsAccess.Expression is MemberAccessExpressionSyntax parentAccess)
            {
                var parentName = parentAccess.Name.Identifier.Text;
                // 排除 parent 是自己的情況
                if (!string.Equals(parentName, controlName, StringComparison.OrdinalIgnoreCase))
                    return parentName;
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
