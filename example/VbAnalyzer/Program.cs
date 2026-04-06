using System.Text.Json;
using VbAnalyzer;
using VbAnalyzer.Analyzers;

string? slnPath = null, projectPath = null, formName = null, outputDir = null;
string libsDir = "libs/";

for (int i = 0; i < args.Length; i++)
    switch (args[i])
    {
        case "--sln": slnPath = args[++i]; break;
        case "--project": projectPath = args[++i]; break;
        case "--form": formName = args[++i]; break;
        case "--output": outputDir = args[++i]; break;
        case "--libs": libsDir = args[++i]; break;
        case "--help": PrintUsage(); return 0;
    }

if (formName == null || outputDir == null || (slnPath == null && projectPath == null))
{ PrintUsage(); return 1; }

// ── Step 1: 收集 .vb ──
Console.Error.WriteLine("[1/5] Collecting .vb files...");
var (vbFiles, detectedVbproj) = SolutionLoader.CollectVbFiles(slnPath, projectPath);
Console.Error.WriteLine($"       Found {vbFiles.Count} .vb files");
if (vbFiles.Count == 0) { Console.Error.WriteLine("[ERROR] No .vb files found."); return 1; }

var projectRoot = slnPath != null
    ? Path.GetDirectoryName(Path.GetFullPath(slnPath))!
    : Path.GetDirectoryName(Path.GetFullPath(projectPath!))!;

// ── Step 2: Compilation ──
Console.Error.WriteLine("[2/5] Building compilation...");
var (compilation, compilationErrors, missingTypes) = CompilationBuilder.Build(vbFiles, libsDir, detectedVbproj);

// ── Step 3: 分析 ──
Console.Error.WriteLine($"[3/5] Analyzing form '{formName}'...");
var methods = MethodAnalyzer.Analyze(compilation, formName, projectRoot);
var controls = ControlAnalyzer.Analyze(compilation, formName, projectRoot);
var events = EventAnalyzer.Analyze(compilation, formName, projectRoot);
var references = ReferenceAnalyzer.Analyze(compilation, formName, projectRoot);
var files = FileAnalyzer.Analyze(compilation, formName, projectRoot);
var layout = LayoutAnalyzer.Build(controls, formName);

Console.Error.WriteLine($"       Methods: {methods.Count}, Controls: {controls.Count}, Events: {events.Count}, References: {references.Count}");

// ── Step 4: 輸出 ──
Console.Error.WriteLine($"[4/5] Writing output to {outputDir}/...");
Directory.CreateDirectory(outputDir);

var jsonOpt = new JsonSerializerOptions
{
    WriteIndented = true,
    PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
    DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
};

// JSON
OutputWriter.WriteWrapped(Path.Combine(outputDir, "files.json"), files,
    $"{formName} 相關檔案清單", "02-code-index-py, 02b-lsp-index", jsonOpt);
OutputWriter.WriteWrapped(Path.Combine(outputDir, "controls.json"), controls,
    $"{formName} 控制項清單（型別、顯示文字、父容器、FarPoint 標記）", "03-reference-scanner, 07-rewrite-prep", jsonOpt);
OutputWriter.WriteWrapped(Path.Combine(outputDir, "events.json"), events,
    $"{formName} 事件 handler 清單", "03-reference-scanner, 04-item-dfs-tracker", jsonOpt);
OutputWriter.WriteWrapped(Path.Combine(outputDir, "methods.json"), methods,
    $"{formName} 方法定義清單（Start/End 行號）", "04-item-dfs-tracker, dfs_lookup.py", jsonOpt);
OutputWriter.WriteWrapped(Path.Combine(outputDir, "references.json"), references,
    $"{formName} 所有 reference 關係（method-call, control-read/write 等）",
    "dfs_lookup.py, build_reference_scan.py, build_shared_state_summary.py", jsonOpt);
OutputWriter.WriteLayout(Path.Combine(outputDir, "layout.json"), layout, formName, jsonOpt);

// Markdown
MdWriter.WriteFilesMd(Path.Combine(outputDir, "files.md"), files, formName);
MdWriter.WriteControlsMd(Path.Combine(outputDir, "controls.md"), controls, formName);
MdWriter.WriteEventsMd(Path.Combine(outputDir, "events.md"), events, formName);
MdWriter.WriteMethodsMd(Path.Combine(outputDir, "methods.md"), methods, formName);
MdWriter.WriteRefsMd(Path.Combine(outputDir, "references.md"), references, formName);
MdWriter.WriteLayoutMd(Path.Combine(outputDir, "layout.md"), layout);

// ── Step 5: 統計 ──
var resolvedCount = references.Count(r => r.ResolvedTo != null);
var stats = new AnalysisStats
{
    TotalFiles = vbFiles.Count, TotalMethods = methods.Count,
    TotalReferences = references.Count,
    ResolvedReferences = resolvedCount,
    UnresolvedReferences = references.Count - resolvedCount,
    ResolvedRate = references.Count > 0 ? Math.Round((double)resolvedCount / references.Count * 100, 1) : 100,
    CompilationErrors = compilationErrors,
    MissingTypes = missingTypes.Take(50).ToList()
};
OutputWriter.WriteJson(Path.Combine(outputDir, "stats.json"), stats, jsonOpt);

Console.Error.WriteLine($"[5/5] Done. Resolved rate: {stats.ResolvedRate}%");
return 0;

static void PrintUsage() => Console.Error.WriteLine(@"
VbAnalyzer — Roslyn-based VB.NET semantic analyzer

Usage:
  dotnet run --project <VbAnalyzer-dir> -- --sln <path.sln> --form <FormName> --output <dir> [--libs <dir>]

Options:
  --sln <path>       .sln 檔案路徑（掃描 solution 內所有 .vbproj）
  --project <path>   單一 .vbproj 路徑（跟 --sln 擇一）
  --form <name>      要分析的 Form 名稱（如 frmOrder）
  --output <dir>     JSON + MD 輸出目錄
  --libs <dir>       第三方 DLL 目錄（預設 libs/）
  --help             顯示用法
");
