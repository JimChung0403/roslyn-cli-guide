using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace VbAnalyzer.Analyzers;

public static class ReferenceAnalyzer
{
    const int MaxExpandRounds = 5;

    /// <summary>
    /// 已知的 control-state 屬性：讀取時應標記為 control-read（即使在鏈式存取中）。
    /// 對齊 Python CONTROL_READ_RE 的屬性清單。
    /// </summary>
    static readonly HashSet<string> ControlStateReadProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "Text", "Value", "Checked", "SelectedIndex", "SelectedValue",
        "CurrentRow", "CurrentCell", "SelectedRows", "SelectedCells",
        "SelectedItem", "SelectedItems", "EditValue", "Items", "Rows"
    };

    /// <summary>
    /// 已知的 control-state 屬性：寫入時應標記為 control-write。
    /// 對齊 Python CONTROL_WRITE_RE 的屬性清單。
    /// </summary>
    static readonly HashSet<string> ControlStateWriteProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "Enabled", "Visible", "ReadOnly", "Text", "Checked",
        "SelectedIndex", "SelectedValue", "BackColor", "ForeColor",
        "EditValue"
    };

    /// <summary>
    /// Iterative expansion：先掃 Form class 自身，再跟著 resolved_to 追進 helper class，
    /// 最多 MaxExpandRounds 輪，遇到 Framework/.NET 內建型別或不在 source 中的 symbol 就停止。
    /// 邏輯與 Python build_vb_form_index.py 的 collect_references + iterative expansion 對齊。
    /// </summary>
    public static (List<ReferenceEntry> refs, HashSet<string> discoveredTypes) Analyze(
        VisualBasicCompilation compilation, string formName, string projectRoot,
        HashSet<string>? controlNames = null)
    {
        var results = new List<ReferenceEntry>();
        var scannedTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 預建 syntax-level fallback 查找表：當 GetSymbolInfo() 因 compilation errors 回傳 null 時，
        // 用純語法層的方法名/屬性名 → 原始碼位置映射來做 fallback 解析
        var fallbackTable = BuildFallbackTable(compilation, projectRoot);
        Console.Error.WriteLine($"       Fallback table: {fallbackTable.Count} method/property names indexed");
        var knownControls = controlNames ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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
                    ScanBlock(classBlock, model, tree, relFile, ownerName, projectRoot, newRefs, discoveredTypes, fallbackTable, knownControls);
                }

                // 掃 module
                foreach (var moduleBlock in root.DescendantNodes().OfType<ModuleBlockSyntax>())
                {
                    var moduleSymbol = model.GetDeclaredSymbol(moduleBlock);
                    if (moduleSymbol == null) continue;
                    if (!typesToScan.Any(t => string.Equals(moduleSymbol.Name, t, StringComparison.OrdinalIgnoreCase)))
                        continue;

                    var ownerName = moduleSymbol.Name;
                    ScanBlock(moduleBlock, model, tree, relFile, ownerName, projectRoot, newRefs, discoveredTypes, fallbackTable, knownControls);
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
        List<ReferenceEntry> results, HashSet<string> discoveredTypes,
        Dictionary<string, List<FallbackEntry>> fallbackTable,
        HashSet<string> knownControls)
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
                    var symbolName = $"{target.ContainingType?.Name ?? "?"}.{target.Name}";
                    var fileLoc = $"{RelPath(span.Path, projectRoot)}:{span.StartLinePosition.Line + 1}";
                    resolvedTo = $"{symbolName} ({fileLoc})";

                    // 發現新的、在 source 中的型別 → 加入待展開清單
                    var targetTypeName = target.ContainingType?.Name;
                    if (targetTypeName != null && !IsExcludedType(target.ContainingType!))
                        discoveredTypes.Add(targetTypeName);
                }
                else if (target.ContainingAssembly != null)
                {
                    var assemblyName = target.ContainingAssembly.Name ?? "";
                    var ns = target.ContainingType?.ContainingNamespace?.ToDisplayString() ?? "";
                    if (ns.StartsWith("System") || ns.StartsWith("Microsoft.VisualBasic")
                        || assemblyName.StartsWith("System") || assemblyName == "mscorlib"
                        || assemblyName.StartsWith("Microsoft.VisualBasic"))
                    {
                        resolvedTo = $"Framework:{target.ContainingType?.ToDisplayString() ?? assemblyName}";
                    }
                    else
                    {
                        // 第三方 DLL：Roslyn 認得型別，追蹤到此為終點
                        resolvedTo = $"DLL:{target.ContainingType?.ToDisplayString() ?? assemblyName}";
                    }
                }

                var refType = ClassifyMethodCall(target);
                if (IsBindingCall(invocation.Expression.ToString()))
                    refType = "binding";

                // Target 用 FormatTarget 保留原始碼中的實例名（如 conexao.Close）
                // 類型資訊已在 ResolvedTo 中保存（如 Framework:System.Data.SqlClient.SqlConnection）
                // 這讓 shared-state-summary 能區分同類型的不同控件/變數
                var syntaxTarget = FormatTarget(invocation.Expression);

                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod,
                    Target = syntaxTarget,
                    File = relFile, Line = line,
                    RefType = refType,
                    Context = context, ResolvedTo = resolvedTo
                });
            }
            else
            {
                // Roslyn semantic model 失敗 → 嘗試 syntax-level fallback
                var refType = IsBindingCall(invocation.Expression.ToString()) ? "binding" : "unresolved";
                string? fallbackResolvedTo = null;
                var targetText = FormatTarget(invocation.Expression);

                var methodName = ExtractSimpleName(invocation.Expression);
                if (methodName != null && fallbackTable.TryGetValue(methodName, out var candidates))
                {
                    // 優先同 class 的 match
                    var match = candidates.FirstOrDefault(c =>
                        string.Equals(c.OwnerType, ownerName, StringComparison.OrdinalIgnoreCase));
                    if (match == null && candidates.Count > 0)
                        match = candidates[0];

                    if (match != null)
                    {
                        fallbackResolvedTo = match.ResolvedTo;
                        targetText = $"{match.OwnerType}.{methodName}";
                        if (refType == "unresolved") refType = "method-call";
                        // 不從 fallback 觸發 type expansion — expansion 只靠 semantic model 驅動
                        // 避免因 syntax-level 名稱碰撞導致 expansion 爆炸
                    }
                }

                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod,
                    Target = targetText,
                    File = relFile, Line = line,
                    RefType = refType,
                    Context = context, ResolvedTo = fallbackResolvedTo
                });
            }
        }

        // MemberAccess（控制項讀寫 + binding assignment + dialog-navigation without parens + fallback for unresolved）
        foreach (var memberAccess in block.DescendantNodes().OfType<MemberAccessExpressionSyntax>())
        {
            if (memberAccess.Parent is InvocationExpressionSyntax) continue;

            // VB.NET 允許省略括號呼叫方法（如 findCustomerForm.ShowDialog 等同 .ShowDialog()）
            // 這些在 AST 中是 MemberAccessExpression 而非 InvocationExpression
            // 先偵測 Show/ShowDialog 無括號呼叫，歸類為 dialog-navigation
            var maName = memberAccess.Name.Identifier.Text;
            if (maName is "Show" or "ShowDialog")
            {
                var navSymbolInfo = model.GetSymbolInfo(memberAccess);
                if (navSymbolInfo.Symbol is IMethodSymbol methodSym && InheritsFrom(methodSym.ContainingType, "Form"))
                {
                    var callerM = $"{ownerName}.{GetContainingMethodName(memberAccess)}";
                    var ctx = memberAccess.Parent?.ToString() ?? memberAccess.ToString();
                    if (ctx.Length > 120) ctx = ctx[..120] + "...";
                    var syntaxTarget = FormatTarget(memberAccess);
                    string? resolvedTo = null;
                    var targetLoc = methodSym.Locations.FirstOrDefault();
                    if (targetLoc != null && !targetLoc.IsInSource)
                    {
                        var ns = methodSym.ContainingType?.ContainingNamespace?.ToDisplayString() ?? "";
                        resolvedTo = ns.StartsWith("System") ? $"Framework:{methodSym.ContainingType?.ToDisplayString()}" : $"DLL:{methodSym.ContainingType?.ToDisplayString()}";
                    }
                    results.Add(new ReferenceEntry
                    {
                        Caller = callerM, Target = syntaxTarget,
                        File = relFile, Line = memberAccess.GetLocation().GetLineSpan().StartLinePosition.Line + 1,
                        RefType = "dialog-navigation", Context = ctx, ResolvedTo = resolvedTo
                    });
                    continue;
                }
            }

            // 跳過中間層 MemberAccess，只捕獲最外層的屬性存取
            // Me.txtName.Text → 內層 "Me.txtName" 跳過，外層 "Me.txtName.Text" 捕獲
            // 但例外：如果當前節點存取的是已知 control-state 屬性（Text, Checked, Enabled 等），
            // 即使 parent 也是 MemberAccess（如 txtName.Text.Trim()），仍然捕獲
            if (memberAccess.Parent is MemberAccessExpressionSyntax)
            {
                var propName = memberAccess.Name.Identifier.Text;
                if (!ControlStateReadProperties.Contains(propName) && !ControlStateWriteProperties.Contains(propName))
                    continue;
            }

            var maTarget = FormatTarget(memberAccess);
            var callerMethod = $"{ownerName}.{GetContainingMethodName(memberAccess)}";
            var maLine = memberAccess.GetLocation().GetLineSpan().StartLinePosition.Line + 1;

            // binding assignment 偵測（如 grdListagem.DataSource = dt）
            if (memberAccess.Parent is AssignmentStatementSyntax assignment
                && assignment.Left == memberAccess
                && maTarget.Contains(".DataSource", StringComparison.OrdinalIgnoreCase))
            {
                var context = assignment.ToString();
                if (context.Length > 120) context = context[..120] + "...";
                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod, Target = maTarget,
                    File = relFile, Line = maLine,
                    RefType = "binding", Context = context, ResolvedTo = null
                });
                continue;
            }

            // 控制項讀寫（需要型別解析）
            var symbolInfo = model.GetSymbolInfo(memberAccess);
            if (symbolInfo.Symbol is IPropertySymbol prop)
            {
                if (prop.ContainingType == null || !IsControlType(prop.ContainingType)) continue;
                // 只捕獲 controlName.Property 格式（有 . 分隔）
                // 跳過純控制項欄位存取（如 txtName → 沒有屬性名）
                if (!maTarget.Contains('.')) continue;
                // 過濾純 layout/visual 屬性（已在 layout.json，不是動態行為）
                if (LayoutOnlyProperties.Contains(prop.Name)) continue;

                var isWrite = memberAccess.Parent is AssignmentStatementSyntax a2 && a2.Left == memberAccess;
                var ctx = memberAccess.Parent?.ToString() ?? memberAccess.ToString();
                if (ctx.Length > 120) ctx = ctx[..120] + "...";

                // 控制項屬性追蹤到此為終點（定義在 DLL 裡）
                string? maResolvedTo = null;
                var propAsm = prop.ContainingAssembly?.Name ?? "";
                var propNs = prop.ContainingType.ContainingNamespace?.ToDisplayString() ?? "";
                if (propNs.StartsWith("System") || propNs.StartsWith("Microsoft.VisualBasic")
                    || propAsm.StartsWith("System") || propAsm == "mscorlib")
                    maResolvedTo = $"Framework:{prop.ContainingType.ToDisplayString()}";
                else
                    maResolvedTo = $"DLL:{prop.ContainingType.ToDisplayString()}";

                results.Add(new ReferenceEntry
                {
                    Caller = callerMethod, Target = maTarget,
                    File = relFile, Line = maLine,
                    RefType = isWrite ? "control-write" : "control-read",
                    Context = ctx, ResolvedTo = maResolvedTo
                });
            }
            else if (symbolInfo.Symbol == null)
            {
                // Fallback：symbol 為 null（compilation errors 導致）
                var memberName = memberAccess.Name.Identifier.Text;
                var receiverName = ExtractSimpleName(memberAccess.Expression);
                var handled = false;

                // Fallback A: 用已知 control 名稱 + control-state 屬性名偵測 control-read/write
                if (receiverName != null && knownControls.Contains(receiverName)
                    && (ControlStateReadProperties.Contains(memberName) || ControlStateWriteProperties.Contains(memberName)))
                {
                    var isWrite = memberAccess.Parent is AssignmentStatementSyntax a4 && a4.Left == memberAccess;
                    if ((isWrite && ControlStateWriteProperties.Contains(memberName)) ||
                        (!isWrite && ControlStateReadProperties.Contains(memberName)))
                    {
                        var ctx = memberAccess.Parent?.ToString() ?? memberAccess.ToString();
                        if (ctx.Length > 120) ctx = ctx[..120] + "...";
                        results.Add(new ReferenceEntry
                        {
                            Caller = callerMethod, Target = maTarget,
                            File = relFile, Line = maLine,
                            RefType = isWrite ? "control-write" : "control-read",
                            Context = ctx, ResolvedTo = null
                        });
                        handled = true;
                    }
                }

                // Fallback B: 用 fallback table 解其他 unresolved member access
                if (!handled && fallbackTable.TryGetValue(memberName, out var propCandidates))
                {
                    var match = propCandidates.FirstOrDefault(c =>
                        string.Equals(c.OwnerType, ownerName, StringComparison.OrdinalIgnoreCase));
                    if (match == null && receiverName != null)
                        match = propCandidates.FirstOrDefault(c =>
                            string.Equals(c.OwnerType, receiverName, StringComparison.OrdinalIgnoreCase));
                    if (match == null && propCandidates.Count == 1)
                        match = propCandidates[0];

                    if (match != null)
                    {
                        var ctx = memberAccess.Parent?.ToString() ?? memberAccess.ToString();
                        if (ctx.Length > 120) ctx = ctx[..120] + "...";
                        var isWrite = memberAccess.Parent is AssignmentStatementSyntax a3 && a3.Left == memberAccess;
                        results.Add(new ReferenceEntry
                        {
                            Caller = callerMethod, Target = maTarget,
                            File = relFile, Line = maLine,
                            RefType = isWrite ? "control-write" : "control-read",
                            Context = ctx, ResolvedTo = match.ResolvedTo
                        });
                    }
                }
                // fallback 都找不到 → drop（避免噪音）
            }
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

    /// <summary>
    /// 純 layout / visual 屬性 — 只在 Designer InitializeComponent 裡設定，不是動態行為。
    /// layout 資訊已在 layout.json，不需要重複出現在 references 裡。
    /// </summary>
    /// <summary>
    /// 純 layout / visual 屬性 — 只在 Designer InitializeComponent 裡設定，不是動態行為。
    /// layout 資訊已在 layout.json，不需要重複出現在 references 裡。
    /// 注意：BackColor, ForeColor 已移除，因為 event handler 中動態設定這些屬性是 UI state 變更
    ///（例如驗證失敗時按鈕變紅），對 React 轉換有意義。
    /// </summary>
    static readonly HashSet<string> LayoutOnlyProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        // Position and sizing（已在 layout.json）
        "Location", "Size", "ClientSize", "MinimumSize", "MaximumSize",
        "AutoSize", "AutoSizeMode", "AutoScaleDimensions", "AutoScaleMode",
        "Anchor", "Dock", "Margin", "Padding",
        // Identity
        "Name",
        // Visual appearance (static-only, not dynamic state)
        "UseVisualStyleBackColor", "Font",
        "BackgroundImage", "BackgroundImageLayout", "Cursor", "RightToLeft",
        "FlatStyle", "Image", "ImageAlign", "TextAlign", "TextImageRelation",
        // Layout metadata
        "ColumnHeadersHeightSizeMode", "RowHeadersWidth", "RowHeadersVisible",
        "ScrollBars", "BorderStyle",
        // Form designer
        "FormBorderStyle", "StartPosition", "WindowState",
        "Icon", "MaximizeBox", "MinimizeBox", "ShowIcon", "ShowInTaskbar",
        "SizeGripStyle", "ImeMode", "KeyPreview",
    };

    /// <summary>
    /// UI component 基底類型名稱 — 這些類型的子類擁有 .Visible, .Enabled 等 UI 屬性，
    /// 即使不繼承 System.Windows.Forms.Control（例如 DevExpress GridColumn, BarItem,
    /// ToolStripItem, DataGridViewColumn 等）。
    /// </summary>
    static readonly HashSet<string> UiComponentBaseNames = new(StringComparer.OrdinalIgnoreCase)
    {
        // WinForms
        "Control", "AxHost", "ToolStripItem", "DataGridViewColumn", "DataGridViewCell",
        // DevExpress common bases
        "GridColumn", "BarItem", "RepositoryItem",
        // ComponentModel (broadest — catches anything registered as a component)
        "Component",
    };

    static bool IsControlType(INamedTypeSymbol type)
    {
        var current = type;
        while (current != null)
        {
            if (UiComponentBaseNames.Contains(current.Name))
            {
                // For "Component" base, require it to be System.ComponentModel.Component
                // to avoid false positives on unrelated classes named Component
                if (current.Name == "Component")
                {
                    var ns = current.ContainingNamespace?.ToDisplayString() ?? "";
                    if (ns == "System.ComponentModel") return true;
                }
                else
                {
                    return true;
                }
            }
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

    // Fallback entry：syntax-level 的方法/屬性定義位置
    record FallbackEntry(string OwnerType, string ResolvedTo);

    /// <summary>
    /// 預掃所有 syntax tree，建立 method/property name → source location 的映射表。
    /// 當 Roslyn semantic model 因 compilation errors 無法解析時，用這張表做 fallback。
    /// </summary>
    static Dictionary<string, List<FallbackEntry>> BuildFallbackTable(
        VisualBasicCompilation compilation, string projectRoot)
    {
        var table = new Dictionary<string, List<FallbackEntry>>(StringComparer.OrdinalIgnoreCase);

        foreach (var tree in compilation.SyntaxTrees)
        {
            var root = tree.GetRoot();
            var file = RelPath(tree.FilePath, projectRoot);

            foreach (var typeBlock in root.DescendantNodes()
                .Where(n => n is ClassBlockSyntax or ModuleBlockSyntax))
            {
                var typeName = typeBlock switch
                {
                    ClassBlockSyntax cb => cb.ClassStatement.Identifier.Text,
                    ModuleBlockSyntax mb => mb.ModuleStatement.Identifier.Text,
                    _ => "?"
                };

                // Methods (Sub / Function)
                foreach (var method in typeBlock.DescendantNodes().OfType<MethodBlockSyntax>())
                {
                    var name = method.SubOrFunctionStatement.Identifier.Text;
                    var loc = method.SubOrFunctionStatement.GetLocation().GetLineSpan();
                    var entry = new FallbackEntry(typeName, $"{typeName}.{name} ({file}:{loc.StartLinePosition.Line + 1})");
                    if (!table.ContainsKey(name)) table[name] = new();
                    table[name].Add(entry);
                }

                // Properties
                foreach (var prop in typeBlock.DescendantNodes().OfType<PropertyBlockSyntax>())
                {
                    var name = prop.PropertyStatement.Identifier.Text;
                    var loc = prop.PropertyStatement.GetLocation().GetLineSpan();
                    var entry = new FallbackEntry(typeName, $"{typeName}.{name} ({file}:{loc.StartLinePosition.Line + 1})");
                    if (!table.ContainsKey(name)) table[name] = new();
                    table[name].Add(entry);
                }
            }
        }

        return table;
    }

    /// <summary>
    /// 從 expression 中取出最末端的簡單名稱。
    /// isUpgradeByWO → "isUpgradeByWO"
    /// Me.Get_Grade_Main_Info → "Get_Grade_Main_Info"
    /// obj.Method → "Method"
    /// </summary>
    static string? ExtractSimpleName(ExpressionSyntax expression)
    {
        return expression switch
        {
            IdentifierNameSyntax id => id.Identifier.Text,
            MemberAccessExpressionSyntax ma => ma.Name.Identifier.Text,
            _ => null
        };
    }

    /// <summary>
    /// 從 syntax tree 遞迴建構 target 名稱，自動去掉 VB.NET 的自我引用前綴。
    /// Me.txtName.Text → txtName.Text
    /// MyBase.OnLoad → OnLoad
    /// MyClass.Method → Method
    /// SomeModule.Method → SomeModule.Method（保留非自我引用的 receiver）
    /// </summary>
    static string FormatTarget(ExpressionSyntax? expr)
    {
        if (expr == null) return "?";
        return expr switch
        {
            MemberAccessExpressionSyntax ma when IsSelfReference(ma.Expression)
                => ma.Name.Identifier.Text,
            MemberAccessExpressionSyntax ma
                => $"{FormatTarget(ma.Expression)}.{ma.Name.Identifier.Text}",
            InvocationExpressionSyntax inv
                => FormatTarget(inv.Expression),
            IdentifierNameSyntax id => id.Identifier.Text,
            _ => expr.ToString()
        };
    }

    static bool IsSelfReference(ExpressionSyntax expr)
    {
        // VB.NET 的 Me / MyBase / MyClass 是專用 syntax node，不是 IdentifierNameSyntax
        return expr is MeExpressionSyntax
            or MyBaseExpressionSyntax
            or MyClassExpressionSyntax;
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
