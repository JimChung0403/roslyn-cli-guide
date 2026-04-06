namespace VbAnalyzer;

/// <summary>
/// 所有 record 的欄位名稱和結構與 Python build_vb_form_index.py 的 dataclass 完全對齊，
/// 讓 lsp-index/*.json 和 code-index/*.json 可被下游 agent 用相同邏輯讀取。
/// C# PascalCase 透過 JsonNamingPolicy.SnakeCaseLower 自動轉為 snake_case。
/// </summary>

// 對應 Python FileEntry
public record FileEntry
{
    public string Path { get; init; } = "";
    public string Role { get; init; } = "";      // "main", "designer", "partial", "helper", "resolved-dependency"
    public string Reason { get; init; } = "";
}

// 對應 Python ControlEntry
public record ControlEntry
{
    public string Name { get; init; } = "";
    public string ControlType { get; init; } = "";
    public string Declaration { get; init; } = "";       // "file:line" 格式
    public string? Initialization { get; init; }          // "file:line" 格式
    public string? Parent { get; init; }
    public List<string> DefaultProperties { get; init; } = [];
    public string? DisplayText { get; init; }
    public int? LocationX { get; init; }
    public int? LocationY { get; init; }
    public int? SizeW { get; init; }
    public int? SizeH { get; init; }
    public bool IsWithEvents { get; init; }
    public bool IsAxFpSpread { get; init; }
    public List<string> SpreadSortRefs { get; init; } = [];
    public List<string> SpreadExportRefs { get; init; } = [];
}

// 對應 Python EventEntry
public record EventEntry
{
    public string Handler { get; init; } = "";
    public string Control { get; init; } = "";
    public string EventType { get; init; } = "";
    public string Definition { get; init; } = "";         // "file:line" 格式
    public List<string> Wireups { get; init; } = [];      // ["file:line", ...] 格式
}

// 對應 Python MethodEntry
public record MethodEntry
{
    public string Name { get; init; } = "";
    public string Owner { get; init; } = "";
    public string File { get; init; } = "";
    public int StartLine { get; init; }
    public int EndLine { get; init; }
    public string Signature { get; init; } = "";          // "Sub btnSave_Click(sender As Object, e As EventArgs)"
    public List<string> Callers { get; init; } = [];
}

// 對應 Python ReferenceEntry
public record ReferenceEntry
{
    public string Caller { get; init; } = "";
    public string Target { get; init; } = "";
    public string File { get; init; } = "";
    public int Line { get; init; }
    public string RefType { get; init; } = "";
    public string Context { get; init; } = "";            // 呼叫所在的程式碼片段
    public string? ResolvedTo { get; init; }              // "file:line" 格式，null = unresolved
}

// JSON 外層包裝（對應 Python 的 { "_purpose", "_consumers", "data" }）
public record JsonWrapper<T>
{
    public string Purpose { get; init; } = "";
    public string Consumers { get; init; } = "";
    public List<T> Data { get; init; } = [];
}

// stats.json（Roslyn 獨有，Python 沒有）
public record AnalysisStats
{
    public int TotalFiles { get; init; }
    public int TotalMethods { get; init; }
    public int TotalReferences { get; init; }
    public int ResolvedReferences { get; init; }
    public int UnresolvedReferences { get; init; }
    public double ResolvedRate { get; init; }
    public int CompilationErrors { get; init; }
    public List<string> MissingTypes { get; init; } = [];
}
