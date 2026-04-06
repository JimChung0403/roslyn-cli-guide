using System.Text.Json;
using VbAnalyzer;
using VbAnalyzer.Analyzers;

// ── CLI 參數解析 ──────────────────────────────────────────────

string? slnPath = null;
string? projectPath = null;
string? formName = null;
string? outputDir = null;
string libsDir = "libs/";

for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--sln": slnPath = args[++i]; break;
        case "--project": projectPath = args[++i]; break;
        case "--form": formName = args[++i]; break;
        case "--output": outputDir = args[++i]; break;
        case "--libs": libsDir = args[++i]; break;
        case "--help":
            PrintUsage();
            return 0;
    }
}

if (formName == null || outputDir == null || (slnPath == null && projectPath == null))
{
    PrintUsage();
    return 1;
}

// ── Step 1: 收集 .vb 檔案 ────────────────────────────────────

Console.Error.WriteLine("[1/5] Collecting .vb files...");
var (vbFiles, detectedVbproj) = SolutionLoader.CollectVbFiles(slnPath, projectPath);
Console.Error.WriteLine($"       Found {vbFiles.Count} .vb files");

if (vbFiles.Count == 0)
{
    Console.Error.WriteLine("[ERROR] No .vb files found. Check --sln or --project path.");
    return 1;
}

// 推算 project root（.sln 所在目錄，或 .vbproj 所在目錄）
var projectRoot = slnPath != null
    ? Path.GetDirectoryName(Path.GetFullPath(slnPath))!
    : Path.GetDirectoryName(Path.GetFullPath(projectPath!))!;

// ── Step 2: 建立 Compilation ──────────────────────────────────

Console.Error.WriteLine("[2/5] Building compilation...");
var (compilation, compilationErrors, missingTypes) = CompilationBuilder.Build(vbFiles, libsDir, detectedVbproj);
Console.Error.WriteLine($"       {compilationErrors} compilation errors, {missingTypes.Count} missing types");

// ── Step 3: 執行分析 ─────────────────────────────────────────

Console.Error.WriteLine($"[3/5] Analyzing form '{formName}'...");
var methods = MethodAnalyzer.Analyze(compilation, formName, projectRoot);
var controls = ControlAnalyzer.Analyze(compilation, formName, projectRoot);
var events = EventAnalyzer.Analyze(compilation, formName, projectRoot);
var references = ReferenceAnalyzer.Analyze(compilation, formName, projectRoot);
var files = FileAnalyzer.Analyze(compilation, formName, projectRoot);

Console.Error.WriteLine($"       Methods: {methods.Count}, Controls: {controls.Count}, Events: {events.Count}, References: {references.Count}");

// ── Step 4: 輸出 JSON ──────────────────────────────────────

Console.Error.WriteLine($"[4/5] Writing output to {outputDir}/...");
Directory.CreateDirectory(outputDir);

var jsonOptions = new JsonSerializerOptions
{
    WriteIndented = true,
    PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
    DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
};

OutputWriter.WriteWrapped(Path.Combine(outputDir, "files.json"), files,
    $"{formName} 相關檔案清單",
    "02-code-index-py, 02b-lsp-index", jsonOptions);

OutputWriter.WriteWrapped(Path.Combine(outputDir, "controls.json"), controls,
    $"{formName} 控制項清單（型別、父容器、FarPoint 標記）",
    "03-reference-scanner, 07-rewrite-prep", jsonOptions);

OutputWriter.WriteWrapped(Path.Combine(outputDir, "events.json"), events,
    $"{formName} 事件 handler 清單",
    "03-reference-scanner, 04-item-dfs-tracker", jsonOptions);

OutputWriter.WriteWrapped(Path.Combine(outputDir, "methods.json"), methods,
    $"{formName} 方法定義清單（Start/End 行號）",
    "04-item-dfs-tracker, dfs_lookup.py", jsonOptions);

OutputWriter.WriteWrapped(Path.Combine(outputDir, "references.json"), references,
    $"{formName} 所有 reference 關係（method-call, control-read/write 等）",
    "dfs_lookup.py, build_reference_scan.py, build_shared_state_summary.py", jsonOptions);

// .md 輸出（格式與 Python build_vb_form_index.py 一致）
MdWriter.WriteFilesMd(Path.Combine(outputDir, "files.md"), files, formName);
MdWriter.WriteControlsMd(Path.Combine(outputDir, "controls.md"), controls, formName);
MdWriter.WriteEventsMd(Path.Combine(outputDir, "events.md"), events, formName);
MdWriter.WriteMethodsMd(Path.Combine(outputDir, "methods.md"), methods, formName);
MdWriter.WriteRefsMd(Path.Combine(outputDir, "references.md"), references, formName);

// ── Step 5: 輸出統計 ──────────────────────────────────────

var resolvedCount = references.Count(r => r.ResolvedTo != null);
var unresolvedCount = references.Count(r => r.ResolvedTo == null);
var stats = new AnalysisStats
{
    TotalFiles = vbFiles.Count,
    TotalMethods = methods.Count,
    TotalReferences = references.Count,
    ResolvedReferences = resolvedCount,
    UnresolvedReferences = unresolvedCount,
    ResolvedRate = references.Count > 0 ? Math.Round((double)resolvedCount / references.Count * 100, 1) : 100,
    CompilationErrors = compilationErrors,
    MissingTypes = missingTypes.Take(50).ToList()
};

OutputWriter.WriteJson(Path.Combine(outputDir, "stats.json"), stats, jsonOptions);

Console.Error.WriteLine($"[5/5] Done. Resolved rate: {stats.ResolvedRate}%");
Console.Error.WriteLine($"       Output: {outputDir}/");

return 0;

// ── Helper ──────────────────────────────────────────────────

static void PrintUsage()
{
    Console.Error.WriteLine(@"
VbAnalyzer — Roslyn-based VB.NET semantic analyzer

Usage:
  dotnet run -- --sln <path.sln> --form <FormName> --output <dir> [--libs <dir>]
  dotnet run -- --project <path.vbproj> --form <FormName> --output <dir> [--libs <dir>]

Options:
  --sln <path>       Path to .sln file (scans all .vbproj in solution)
  --project <path>   Path to single .vbproj (alternative to --sln)
  --form <name>      Form name to analyze (e.g. frmOrder)
  --output <dir>     Output directory for JSON files
  --libs <dir>       Directory containing third-party DLLs (default: libs/)
  --help             Show this help
");
}
