using System.Text;

namespace VbAnalyzer;

public static class MdWriter
{
    public static void WriteFilesMd(string path, List<FileEntry> entries, string formName)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"<!-- 用途：{formName} 相關檔案清單（主檔、Designer、partial、helper） -->");
        sb.AppendLine($"<!-- 使用者：02-code-index-py、02b-lsp-index -->");
        sb.AppendLine();
        sb.AppendLine($"# {formName} Files Index");
        sb.AppendLine();
        sb.AppendLine("## 檔案清單");
        sb.AppendLine(entries.Count == 0
            ? MdTable(["檔案", "角色", "關聯原因"], [["⚠ 待確認", "", "沒有找到相關檔案"]])
            : MdTable(["檔案", "角色", "關聯原因"], entries.Select(e => new[] { e.Path, e.Role, e.Reason }).ToList()));

        if (entries.Count >= 5)
        {
            sb.AppendLine(); sb.AppendLine("## Mermaid 圖表"); sb.AppendLine("```mermaid"); sb.AppendLine("graph TD");
            for (int i = 0; i < entries.Count; i++)
            {
                sb.AppendLine($"    N{i}[{System.IO.Path.GetFileName(entries[i].Path)}]");
                if (i > 0) sb.AppendLine($"    N0 --> N{i}");
            }
            sb.AppendLine("```");
        }
        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        Console.Error.WriteLine($"       Wrote {path}");
    }

    public static void WriteControlsMd(string path, List<ControlEntry> entries, string formName)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"<!-- 用途：{formName} 控制項清單（型別、顯示文字、父容器、FarPoint 標記） -->");
        sb.AppendLine($"<!-- 使用者：03-reference-scanner（分類到 17 類）、07-rewrite-prep（控制項語義映射） -->");
        sb.AppendLine();
        sb.AppendLine($"# {formName} Controls Index");
        sb.AppendLine();

        var rows = entries.Select(e =>
        {
            var notes = new List<string>();
            if (e.IsWithEvents) notes.Add("WithEvents");
            if (e.IsAxFpSpread) notes.Add("FarPoint ActiveX Spread");
            if (e.SpreadSortRefs.Count > 0) notes.Add("sort:" + string.Join(", ", e.SpreadSortRefs.Take(3)));
            if (e.SpreadExportRefs.Count > 0) notes.Add("export:" + string.Join(", ", e.SpreadExportRefs.Take(3)));
            return new[] { e.Name, e.ControlType, e.DisplayText ?? "", e.Declaration, e.Initialization ?? "",
                e.Parent ?? "", string.Join("<br/>", e.DefaultProperties.Take(5)), string.Join("<br/>", notes) };
        }).ToList();

        sb.AppendLine(rows.Count == 0
            ? MdTable(["Control", "Type", "顯示文字", "宣告位置", "初始化位置", "父容器", "預設屬性", "備註"],
                [["⚠ 待確認", "", "", "", "", "", "", "沒有找到控制項"]])
            : MdTable(["Control", "Type", "顯示文字", "宣告位置", "初始化位置", "父容器", "預設屬性", "備註"], rows));

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        Console.Error.WriteLine($"       Wrote {path}");
    }

    public static void WriteEventsMd(string path, List<EventEntry> entries, string formName)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"<!-- 用途：{formName} 事件 handler 清單（handler 名稱、控制項、事件型別、定義位置） -->");
        sb.AppendLine($"<!-- 使用者：03-reference-scanner（分類 entry point）、04-item-dfs-tracker（DFS 進入點，整份載入） -->");
        sb.AppendLine();
        sb.AppendLine($"# {formName} Events Index");
        sb.AppendLine();

        var rows = entries.Select(e => new[] { e.Handler, e.Control, e.EventType, e.Definition, string.Join("<br/>", e.Wireups) }).ToList();
        sb.AppendLine(rows.Count == 0
            ? MdTable(["Handler", "Control", "Event", "定義位置", "掛載位置"], [["⚠ 待確認", "", "", "", ""]])
            : MdTable(["Handler", "Control", "Event", "定義位置", "掛載位置"], rows));

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        Console.Error.WriteLine($"       Wrote {path}");
    }

    public static void WriteMethodsMd(string path, List<MethodEntry> entries, string formName)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"<!-- 用途：{formName} 方法定義清單（方法名稱、Owner、Start/End 行號） -->");
        sb.AppendLine($"<!-- 使用者：04-item-dfs-tracker（DFS 方法定位，整份載入）、dfs_lookup.py（查目標範圍） -->");
        sb.AppendLine();
        sb.AppendLine($"# {formName} Methods Index");
        sb.AppendLine();

        var rows = entries.Select(e => new[] { e.Name, e.Owner, $"{e.File}:{e.StartLine}",
            e.StartLine.ToString(), e.EndLine.ToString(), string.Join(", ", e.Callers.Distinct().OrderBy(c => c)) }).ToList();
        sb.AppendLine(rows.Count == 0
            ? MdTable(["Method", "Owner", "定義位置", "Start", "End", "呼叫來源"], [["⚠ 待確認", "", "", "", "", ""]])
            : MdTable(["Method", "Owner", "定義位置", "Start", "End", "呼叫來源"], rows));

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        Console.Error.WriteLine($"       Wrote {path}");
    }

    public static void WriteRefsMd(string path, List<ReferenceEntry> entries, string formName)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"<!-- 用途：{formName} 所有 reference 關係（method-call、control-read/write、binding 等） -->");
        sb.AppendLine($"<!-- 使用者：04-item-dfs-tracker（透過 dfs_lookup.py 按需查詢，不整份載入）、build_shared_state_summary.py、build_reference_scan.py -->");
        sb.AppendLine();
        sb.AppendLine($"# {formName} References Index");
        sb.AppendLine();

        var rows = entries.Select(e => new[] { e.Caller, e.Target, e.RefType, $"{e.File}:{e.Line}",
            e.ResolvedTo ?? "", e.Context.Replace("|", "\\|") }).ToList();
        sb.AppendLine(rows.Count == 0
            ? MdTable(["Caller", "Target", "類型", "位置", "Resolved", "Context"], [["⚠ 待確認", "", "", "", "", ""]])
            : MdTable(["Caller", "Target", "類型", "位置", "Resolved", "Context"], rows));

        if (entries.Count >= 5)
        {
            sb.AppendLine(); sb.AppendLine("## Mermaid 圖表"); sb.AppendLine("```mermaid"); sb.AppendLine("graph TD");
            foreach (var (e, i) in entries.Take(12).Select((e, i) => (e, i)))
                sb.AppendLine($"    C{i}[{e.Caller}] --> T{i}[{e.Target}]");
            sb.AppendLine("```");
        }
        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        Console.Error.WriteLine($"       Wrote {path}");
    }

    public static void WriteLayoutMd(string path, LayoutData layout)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"<!-- 用途：{layout.Form} UI 容器層級與控制項位置（按容器分組、按 y 分排、按 x 排序） -->");
        sb.AppendLine($"<!-- 使用者：07-rewrite-prep（映射 React component 層級）、未來 React Agent Team -->");
        sb.AppendLine();
        sb.AppendLine($"# {layout.Form} Layout");
        sb.AppendLine();

        foreach (var container in layout.Containers)
        {
            sb.AppendLine($"## {container.Container}");
            sb.AppendLine();
            for (int i = 0; i < container.Rows.Count; i++)
            {
                var row = container.Rows[i];
                sb.AppendLine($"### Row {i + 1} (y≈{row.Y})");
                sb.AppendLine();
                sb.AppendLine("| 順序 | Control | 顯示文字 | Type | x | y | w | h |");
                sb.AppendLine("|---|---|---|---|---|---|---|---|");
                for (int j = 0; j < row.Controls.Count; j++)
                {
                    var c = row.Controls[j];
                    sb.AppendLine($"| {j + 1} | {c.Name} | {c.Text} | {c.Type} | {c.X} | {c.Y} | {c.W} | {c.H} |");
                }
                sb.AppendLine();
            }
        }

        var totalContainers = layout.Containers.Count;
        var totalControls = layout.Containers.Sum(c => c.Rows.Sum(r => r.Controls.Count));
        sb.AppendLine($"**統計**: {totalContainers} 個容器, {totalControls} 個有座標的控制項");

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        Console.Error.WriteLine($"       Wrote {path}");
    }

    static string MdTable(string[] headers, List<string[]> rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine("| " + string.Join(" | ", headers) + " |");
        sb.AppendLine(string.Concat(Enumerable.Repeat("|---", headers.Length)) + "|");
        foreach (var row in rows)
            sb.AppendLine("| " + string.Join(" | ", row) + " |");
        return sb.ToString().TrimEnd();
    }
}
