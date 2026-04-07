using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class ReferenceAnalyzer
{
    const int MaxExpandRounds = 5;

    /// <summary>
    /// Iterative expansion：先掃 Form class 自身，再跟著 resolved_to 追進 helper class，
    /// 最多 MaxExpandRounds 輪，遇到 Framework/.NET 內建型別或不在 source 中的 symbol 就停止。
    /// 邏輯與 Python build_vb_form_index.py 的 collect_references + iterative expansion 對齊。
    /// </summary>
    public static (List<ReferenceEntry> refs, HashSet<string> discoveredTypes) Analyze(VisualBasicCompilation compilation, string formName, string projectRoot)
    {
        var results = new List<ReferenceEntry>();
        var scannedTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 第一輪：掃 Form class 自身
        var typesToScan = new List<string> { formName };
        scannedTypes.Add(formName);

        for (int round = 0; round <= MaxExpandRounds; round++)
        {
            if (typesToScan.Count == 0) break;

            var newRefs = new List<ReferenceEntry>();
            var discoveredTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var tree in compilation.SyntaxTrees)
            {
                var model = compilation.GetSemanticModel(tree);
                var root = tree.GetRoot();
                var relFile = RelPath(tree.FilePath, projectRoot);

                // 掃 class
                foreach (var classBlock in root.DescendantNodes().OfType<ClassBlockSyntax>())
                {
                    var classSymbol = model.GetDeclaredSymbol(classBlock);
                    if (classSymbol == null) continue;
                    if (!typesToScan.Any(t => string.Equals(classSymbol.Name, t, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var ownerName = classSymbol.Name;
                    ScanBlock(classBlock, model, tree, relFile, ownerName, projectRoot, newRefs, discoveredTypes);
                }

                // 掃 module
                foreach (var moduleBlock in root.DescendantNodes().OfType<ModuleBlockSyntax>())
                {
                    var moduleSymbol = model.GetDeclaredSymbol(moduleBlock);
                    if (moduleSymbol == null) continue;
                    if (!typesToScan.Any(t => string.Equals(moduleSymbol.Name, t, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var ownerName = moduleSymbol.Name;
                    // Module 沒有 ClassBlockSyntax，但內部結構一樣有 MethodBlock、MemberAccess
                    ScanBlock(moduleBlock, model, tree, relFile, ownerName, projectRoot, newRefs, discoveredTypes);
                }
            }

            results.AddRange(newRefs);

            // 從新發現的 ref 中找出還沒掃過的、在 source 中的型別
            typesToScan = discoveredTypes
                .Where(t => !scannedTypes.Contains(t))
                .ToList();

            foreach (var t in typesToScan)
                scannedTypes.Add(t);

            if (round < MaxExpandRounds && typesToScan.Count > 0)
                Console.Error.WriteLine($"       Expansion round {round + 1}: discovered {typesToScan.Count} new types ({string.Join(", ", typesToScan.Take(5))})");
        }

        return (results.OrderBy(r => r.File).ThenBy(r => r.Line).ToList(), scannedTypes);
    }

    /// <summary>
    /// 掃描一個 type block（ClassBlock 或 ModuleBlock）內的所有 invocation 和 member access。
    /// </summary>
    static void ScanBlock(SyntaxNode block, SemanticModel model, SyntaxTree tree,
        string relFile, string ownerName, string projectRoot,
        List<ReferenceEntry> results, HashSet<string> discoveredTypes)
    {
        // Invocations
        foreach (var invocation in block.DescendantNodes().OfType<InvocationExpressionSyntax>())
        {
            var line = invocation.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
            var callerMethod = $"{ownerName}.{GetContainingMethodName(invocation)}";
            var context = invocation.ToString();
            if (context.Length > 120) context = context[..120] + "...";

            var symbolInfo = model.GetSymbolInfo(invocation);

            if (symbolInfo.Symbol is IMethodSymbol target)
            {
                var targetLoc = target.Locations.FirstOrDefault();
                string? resolvedTo = null;

                if (targetLoc != null && targetLoc.IsInSource)
                {
                    var span = targetLoc.GetLineSpan();
                    resolvedTo = $"{RelPath(span.Path, projectRoot)}:{span.StartLinePosition.Line + 1}";

                    // 發現新的、在 source 中的型別 → 加入待展開清單
                    var targetTypeName = target.ContainingType?.Name;
                    if (targetTypeName != null && !IsExcludedType(target.ContainingType!))
                        discoveredTypes.Add(targetTypeName);
                }
                else if (target.ContainingAssembly != null)
                {
                    // 不在 source 但 Roslyn 成功解析
                    // 只有 .NET Framework 原生的才算 resolved（System.*, Microsoft.VisualBasic.*）
                    // 第三方 DLL 不算
                    var assemblyName = target.ContainingAssembly.Name ?? "";
                    var ns = target.ContainingType?.ContainingNamespace?.ToDisplayString() ?? "";
                    if (ns.StartsWith("System") || ns.StartsWith("Microsoft.VisualBasic")
                        || assemblyName.StartsWith("System") || assemblyName == "mscorlib"
                        || assemblyName.StartsWith("Microsoft.VisualBasic"))
                    {
                        resolvedTo = $"Framework:{target.ContainingType?.ToDisplayString() ?? assemblyName}";
                    }
                }

                var refType = ClassifyMethodCall(target);
                // binding 偵測：覆寫 ref_type（即使 Roslyn 解析成功，binding 優先）
                if (IsBindingCall(invocation.Expression.ToString()))
                    refType = "binding";

                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod,
                    Target = $"{target.ContainingType?.Name ?? "?"}.{target.Name}",
                    File = relFile, Line = line,
                    RefType = refType,
                    Context = context, ResolvedTo = resolvedTo
                });
            }
            else
            {
                // unresolved 也做 binding 文字比對（缺 DLL 時 DataBindings.Add 會是 unresolved）
                var refType = IsBindingCall(invocation.Expression.ToString()) ? "binding" : "unresolved";

                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod,
                    Target = invocation.Expression.ToString(),
                    File = relFile, Line = line,
                    RefType = refType,
                    Context = context, ResolvedTo = null
                });
            }
        }

        // MemberAccess（控制項讀寫 + binding assignment）
        foreach (var memberAccess in block.DescendantNodes().OfType<MemberAccessExpressionSyntax>())
        {
            if (memberAccess.Parent is InvocationExpressionSyntax) continue;

            var maText = memberAccess.ToString();
            var callerMethod = $"{ownerName}.{GetContainingMethodName(memberAccess)}";
            var maLine = memberAccess.GetLocation().GetLineSpan().StartLinePosition.Line + 1;

            // binding assignment 偵測（如 grdListagem.DataSource = dt）
            // 用文字比對，不依賴型別解析（缺 DLL 也能抓到）
            if (memberAccess.Parent is AssignmentStatementSyntax assignment
                && assignment.Left == memberAccess
                && maText.Contains(".DataSource", StringComparison.OrdinalIgnoreCase))
            {
                var context = assignment.ToString();
                if (context.Length > 120) context = context[..120] + "...";
                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod, Target = maText,
                    File = relFile, Line = maLine,
                    RefType = "binding", Context = context, ResolvedTo = null
                });
                continue;
            }

            // 控制項讀寫（需要型別解析）
            var symbolInfo = model.GetSymbolInfo(memberAccess);
            if (symbolInfo.Symbol is not IPropertySymbol prop) continue;
            if (prop.ContainingType == null || !IsControlType(prop.ContainingType)) continue;

            var isWrite = memberAccess.Parent is AssignmentStatementSyntax a2 && a2.Left == memberAccess;
            var ctx = memberAccess.Parent?.ToString() ?? maText;
            if (ctx.Length > 120) ctx = ctx[..120] + "...";

            results.Add(new ReferenceEntry
            {
                Caller = callerMethod, Target = maText,
                File = relFile, Line = maLine,
                RefType = isWrite ? "control-write" : "control-read",
                Context = ctx, ResolvedTo = null
            });
        }
    }

    /// <summary>
    /// 偵測 binding 呼叫，對齊 Python 的 BINDING_RE：
    /// DataSource, DataBindings.Add, Rows.Add, Rows.Remove, Rows.Clear, Columns.Add, Columns.Clear
    /// </summary>
    static bool IsBindingCall(string expressionText)
    {
        var lower = expressionText.ToLowerInvariant();
        return lower.Contains(".datasource")
            || lower.Contains(".databindings.add")
            || lower.Contains(".rows.add")
            || lower.Contains(".rows.remove")
            || lower.Contains(".rows.clear")
            || lower.Contains(".columns.add")
            || lower.Contains(".columns.clear");
    }

    static string ClassifyMethodCall(IMethodSymbol target)
    {
        var typeName = target.ContainingType?.Name ?? "";
        if (typeName.EndsWith("Service") || typeName.EndsWith("Client") || typeName.EndsWith("Proxy") || typeName.EndsWith("SoapClient"))
            return "remote-call";
        if (target.Name is "Show" or "ShowDialog" && InheritsFrom(target.ContainingType, "Form"))
            return "dialog-navigation";
        if (typeName.Contains("Spread") || typeName.Contains("FpSpread") || typeName.Contains("SheetView"))
        {
            if (target.Name.Contains("Sort", StringComparison.OrdinalIgnoreCase)) return "spread-sort";
            if (target.Name.Contains("Export", StringComparison.OrdinalIgnoreCase) || target.Name.Contains("Save", StringComparison.OrdinalIgnoreCase)) return "spread-export";
        }
        return "method-call";
    }

    /// <summary>
    /// 排除 Framework/.NET 內建型別和 Resources 等非程式邏輯型別。
    /// 這些型別不展開——跟 Python 的 is_framework_builtin + "Framework" 停止條件對齊。
    /// </summary>
    static bool IsExcludedType(INamedTypeSymbol type)
    {
        var name = type.Name;
        var ns = type.ContainingNamespace?.ToDisplayString() ?? "";
        if (name is "Resources" or "MySettings" or "MyProject") return true;
        if (ns.Contains("My.Resources") || ns.Contains("My.")) return true;
        if (ns.StartsWith("System.") || ns == "System" || ns.StartsWith("Microsoft.")) return true;
        return false;
    }

    static bool IsControlType(INamedTypeSymbol type)
    {
        var current = type;
        while (current != null)
        {
            if (current.Name == "Control" && current.ContainingNamespace?.ToDisplayString() == "System.Windows.Forms") return true;
            if (current.Name == "AxHost") return true;
            current = current.BaseType;
        }
        return false;
    }

    static bool InheritsFrom(INamedTypeSymbol? type, string baseName)
    {
        var c = type;
        while (c != null) { if (c.Name == baseName) return true; c = c.BaseType; }
        return false;
    }

    static string GetContainingMethodName(SyntaxNode node)
    {
        var c = node.Parent;
        while (c != null)
        {
            if (c is MethodBlockSyntax m) return m.SubOrFunctionStatement.Identifier.Text;
            if (c is PropertyBlockSyntax p) return p.PropertyStatement.Identifier.Text;
            c = c.Parent;
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
