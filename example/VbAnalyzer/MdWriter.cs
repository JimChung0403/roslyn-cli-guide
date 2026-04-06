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

        if (entries.Count == 0)
        {
            sb.AppendLine(MdTable(["檔案", "角色", "關聯原因"], [["⚠ 待確認", "", "沒有找到相關檔案"]]));
        }
        else
        {
            var rows = entries.Select(e => new[] { e.Path, e.Role, e.Reason }).ToList();
            sb.AppendLine(MdTable(["檔案", "角色", "關聯原因"], rows));
        }

        if (entries.Count >= 5)
        {
            sb.AppendLine();
            sb.AppendLine("## Mermaid 圖表");
            sb.AppendLine("```mermaid");
            sb.AppendLine("graph TD");
            for (int i = 0; i < entries.Count; i++)
            {
                var fileName = System.IO.Path.GetFileName(entries[i].Path);
                sb.AppendLine($"    N{i}[{fileName}]");
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

        if (entries.Count == 0)
        {
            sb.AppendLine(MdTable(
                ["Control", "Type", "顯示文字", "宣告位置", "初始化位置", "父容器", "預設屬性", "備註"],
                [["⚠ 待確認", "", "", "", "", "", "", "沒有找到控制項"]]));
        }
        else
        {
            var rows = entries.Select(e =>
            {
                var notes = new List<string>();
                if (e.IsWithEvents) notes.Add("WithEvents");
                if (e.IsAxFpSpread) notes.Add("FarPoint ActiveX Spread");
                if (e.SpreadSortRefs.Count > 0) notes.Add("sort:" + string.Join(", ", e.SpreadSortRefs.Take(3)));
                if (e.SpreadExportRefs.Count > 0) notes.Add("export:" + string.Join(", ", e.SpreadExportRefs.Take(3)));

                return new[]
                {
                    e.Name,
                    e.ControlType,
                    e.DisplayText ?? "",
                    e.Declaration,
                    e.Initialization ?? "",
                    e.Parent ?? "",
                    string.Join("<br/>", e.DefaultProperties.Take(5)),
                    string.Join("<br/>", notes)
                };
            }).ToList();

            sb.AppendLine(MdTable(
                ["Control", "Type", "顯示文字", "宣告位置", "初始化位置", "父容器", "預設屬性", "備註"],
                rows));
        }

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

        if (entries.Count == 0)
        {
            sb.AppendLine(MdTable(["Handler", "Control", "Event", "定義位置", "掛載位置"],
                [["⚠ 待確認", "", "", "", ""]]));
        }
        else
        {
            var rows = entries.Select(e => new[]
            {
                e.Handler, e.Control, e.EventType, e.Definition,
                string.Join("<br/>", e.Wireups)
            }).ToList();

            sb.AppendLine(MdTable(["Handler", "Control", "Event", "定義位置", "掛載位置"], rows));
        }

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

        if (entries.Count == 0)
        {
            sb.AppendLine(MdTable(["Method", "Owner", "定義位置", "Start", "End", "呼叫來源"],
                [["⚠ 待確認", "", "", "", "", ""]]));
        }
        else
        {
            var rows = entries.Select(e => new[]
            {
                e.Name, e.Owner, $"{e.File}:{e.StartLine}",
                e.StartLine.ToString(), e.EndLine.ToString(),
                string.Join(", ", e.Callers.Distinct().OrderBy(c => c))
            }).ToList();

            sb.AppendLine(MdTable(["Method", "Owner", "定義位置", "Start", "End", "呼叫來源"], rows));
        }

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

        if (entries.Count == 0)
        {
            sb.AppendLine(MdTable(["Caller", "Target", "類型", "位置", "Resolved", "Context"],
                [["⚠ 待確認", "", "", "", "", ""]]));
        }
        else
        {
            var rows = entries.Select(e => new[]
            {
                e.Caller, e.Target, e.RefType,
                $"{e.File}:{e.Line}",
                e.ResolvedTo ?? "",
                e.Context.Replace("|", "\\|")
            }).ToList();

            sb.AppendLine(MdTable(["Caller", "Target", "類型", "位置", "Resolved", "Context"], rows));
        }

        if (entries.Count >= 5)
        {
            sb.AppendLine();
            sb.AppendLine("## Mermaid 圖表");
            sb.AppendLine("```mermaid");
            sb.AppendLine("graph TD");
            foreach (var (entry, idx) in entries.Take(12).Select((e, i) => (e, i)))
            {
                sb.AppendLine($"    C{idx}[{entry.Caller}] --> T{idx}[{entry.Target}]");
            }
            sb.AppendLine("```");
        }

        File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        Console.Error.WriteLine($"       Wrote {path}");
    }

    /// <summary>
    /// 產出 Markdown 表格，格式與 Python md_table() 一致。
    /// </summary>
    static string MdTable(string[] headers, List<string[]> rows)
    {
        var sb = new StringBuilder();
        sb.AppendLine("| " + string.Join(" | ", headers) + " |");
        sb.AppendLine(string.Concat(Enumerable.Repeat("|---", headers.Length)) + "|");
        foreach (var row in rows)
        {
            sb.AppendLine("| " + string.Join(" | ", row) + " |");
        }
        return sb.ToString().TrimEnd();
    }
}
