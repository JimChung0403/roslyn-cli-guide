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
    public string Role { get; init; } = "";
    public string Reason { get; init; } = "";
}

// 對應 Python ControlEntry
public record ControlEntry
{
    public string Name { get; init; } = "";
    public string ControlType { get; init; } = "";
    public string Declaration { get; init; } = "";
    public string? Initialization { get; init; }
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
    public string Definition { get; init; } = "";
    public List<string> Wireups { get; init; } = [];
}

// 對應 Python MethodEntry
public record MethodEntry
{
    public string Name { get; init; } = "";
    public string Owner { get; init; } = "";
    public string File { get; init; } = "";
    public int StartLine { get; init; }
    public int EndLine { get; init; }
    public string Signature { get; init; } = "";
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
    public string Context { get; init; } = "";
    public string? ResolvedTo { get; init; }
}

// layout.json 的結構
public record LayoutData
{
    public string Form { get; init; } = "";
    public List<ContainerData> Containers { get; init; } = [];
}

public record ContainerData
{
    public string Container { get; init; } = "";
    public List<RowData> Rows { get; init; } = [];
}

public record RowData
{
    public int Y { get; init; }
    public List<LayoutControl> Controls { get; init; } = [];
}

public record LayoutControl
{
    public string Name { get; init; } = "";
    public string Text { get; init; } = "";
    public string Type { get; init; } = "";
    public int? X { get; init; }
    public int? Y { get; init; }
    public int? W { get; init; }
    public int? H { get; init; }
}

// stats.json
public record AnalysisStats
{
    public int TotalFiles { get; init; }
    public int TotalMethods { get; init; }
    public int TotalReferences { get; init; }

    // resolved_to != null（target 在原始碼中，有精確的 file:line）
    public int ResolvedReferences { get; init; }
    public int UnresolvedReferences { get; init; }
    public double ResolvedRate { get; init; }

    // ref_type != "unresolved"（Roslyn 成功辨識型別，即使 target 在 DLL 不在原始碼）
    public int AnalyzedReferences { get; init; }
    public double AnalyzedRate { get; init; }

    // ref_type 分佈
    public Dictionary<string, int> RefTypeDistribution { get; init; } = new();

    public int CompilationErrors { get; init; }
    public List<string> MissingTypes { get; init; } = [];
}
