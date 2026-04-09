using System.Text.Json;
using VbAnalyzer;
using VbAnalyzer.Analyzers;

string? slnPath = null, projectPath = null, formName = null, outputDir = null, projectRootOverride = null;
string libsDir = "libs/";
bool debugMode = false;

for (int i = 0; i < args.Length; i++)
    switch (args[i])
    {
        case "--sln": slnPath = args[++i]; break;
        case "--project": projectPath = args[++i]; break;
        case "--form": formName = args[++i]; break;
        case "--output": outputDir = args[++i]; break;
        case "--libs": libsDir = args[++i]; break;
        case "--project-root": projectRootOverride = args[++i]; break;
        case "--debug": debugMode = true; break;
        case "--help": PrintUsage(); return 0;
    }

if (formName == null || outputDir == null || (slnPath == null && projectPath == null) || projectRootOverride == null)
{ PrintUsage(); return 1; }

// ── Step 1: 收集 projects ──
Console.Error.WriteLine("[1/5] Collecting projects...");
var projects = SolutionLoader.CollectProjects(slnPath, projectPath);
var totalFiles = projects.Sum(p => p.VbFiles.Count);
Console.Error.WriteLine($"       {projects.Count} projects, {totalFiles} .vb files total");
if (totalFiles == 0) { Console.Error.WriteLine("[ERROR] No .vb files found."); return 1; }

var projectRoot = Path.GetFullPath(projectRootOverride);

// ── Step 2: Multi-project Compilation ──
Console.Error.WriteLine("[2/5] Building compilation (multi-project)...");
var (compilation, compilationErrors, missingTypes) = CompilationBuilder.Build(projects, libsDir, formName);

// ── Step 3: 分析 ──
Console.Error.WriteLine($"[3/5] Analyzing form '{formName}'...");
var (references, discoveredTypes) = ReferenceAnalyzer.Analyze(compilation, formName, projectRoot);
var methods = MethodAnalyzer.Analyze(compilation, formName, projectRoot, discoveredTypes);
var controls = ControlAnalyzer.Analyze(compilation, formName, projectRoot);
var events = EventAnalyzer.Analyze(compilation, formName, projectRoot);
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

MdWriter.WriteFilesMd(Path.Combine(outputDir, "files.md"), files, formName);
MdWriter.WriteControlsMd(Path.Combine(outputDir, "controls.md"), controls, formName);
MdWriter.WriteEventsMd(Path.Combine(outputDir, "events.md"), events, formName);
MdWriter.WriteMethodsMd(Path.Combine(outputDir, "methods.md"), methods, formName);
MdWriter.WriteRefsMd(Path.Combine(outputDir, "references.md"), references, formName);
MdWriter.WriteLayoutMd(Path.Combine(outputDir, "layout.md"), layout);

// ── Step 5: 統計 ──
var resolvedCount = references.Count(r => r.ResolvedTo != null);
var analyzedCount = references.Count(r => r.RefType != "unresolved");
var refTypeDist = references
    .GroupBy(r => r.RefType)
    .ToDictionary(g => g.Key, g => g.Count());

var stats = new AnalysisStats
{
    TotalFiles = totalFiles, TotalMethods = methods.Count,
    TotalReferences = references.Count,
    ResolvedReferences = resolvedCount,
    UnresolvedReferences = references.Count - resolvedCount,
    ResolvedRate = references.Count > 0 ? Math.Round((double)resolvedCount / references.Count * 100, 1) : 100,
    AnalyzedReferences = analyzedCount,
    AnalyzedRate = references.Count > 0 ? Math.Round((double)analyzedCount / references.Count * 100, 1) : 100,
    RefTypeDistribution = refTypeDist,
    CompilationErrors = compilationErrors,
    MissingTypes = missingTypes.Take(50).ToList()
};
OutputWriter.WriteJson(Path.Combine(outputDir, "stats.json"), stats, jsonOpt);

Console.Error.WriteLine($"[5/5] Done. Analyzed: {stats.AnalyzedRate}% ({analyzedCount}/{references.Count}), Resolved to source: {stats.ResolvedRate}%");

// ── Diagnostic log（--debug 模式才產生） ──
if (debugMode)
{
    var diagPath = Path.Combine(outputDir, "diagnostic.log");
    using (var log = new StreamWriter(diagPath, false, System.Text.Encoding.UTF8))
    {
        log.WriteLine($"=== VbAnalyzer Diagnostic Log (DEBUG MODE) ===");
        log.WriteLine($"Form: {formName}");
        log.WriteLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        log.WriteLine();

        // 1. Project 資訊
        log.WriteLine($"--- Projects ({projects.Count}) ---");
        foreach (var p in projects)
        {
            var (ns, imports) = CompilationBuilder.ParseVbproj(p.VbprojPath);
            log.WriteLine($"  {p.ProjectName}: {p.VbFiles.Count} files, RootNamespace='{ns}', Imports={imports.Count}");
        }
        log.WriteLine();

        // 2. Compilation 品質
        log.WriteLine($"--- Compilation ---");
        log.WriteLine($"  Total .vb files: {totalFiles}");
        log.WriteLine($"  Compilation errors: {compilationErrors}");
        log.WriteLine($"  Missing types: {missingTypes.Count}");
        log.WriteLine();

        // 3. Reference 分類
        log.WriteLine($"--- References ({references.Count}) ---");
        log.WriteLine($"  analyzed_rate: {stats.AnalyzedRate}% ({analyzedCount}/{references.Count})");
        log.WriteLine($"  resolved_rate: {stats.ResolvedRate}% ({resolvedCount}/{references.Count})");
        log.WriteLine();
        log.WriteLine($"  ref_type distribution:");
        foreach (var kv in refTypeDist.OrderByDescending(kv => kv.Value))
            log.WriteLine($"    {kv.Key}: {kv.Value}");
        log.WriteLine();

        // 4. resolved_to 分類
        var resolvedToSource = references.Where(r => r.ResolvedTo != null && !r.ResolvedTo.StartsWith("Framework:")).ToList();
        var resolvedToFramework = references.Where(r => r.ResolvedTo != null && r.ResolvedTo.StartsWith("Framework:")).ToList();
        var resolvedToNull = references.Where(r => r.ResolvedTo == null && r.RefType != "unresolved").ToList();
        var unresolvedRefs = references.Where(r => r.RefType == "unresolved").ToList();

        log.WriteLine($"  resolved_to breakdown:");
        log.WriteLine($"    Source code (file:line):  {resolvedToSource.Count}");
        log.WriteLine($"    Framework (.NET native):  {resolvedToFramework.Count}");
        log.WriteLine($"    Third-party DLL (null):   {resolvedToNull.Count}");
        log.WriteLine($"    Unresolved (null):        {unresolvedRefs.Count}");
        log.WriteLine();

        // 5. 每一筆 NOT resolved 的原因（核心 debug 資訊）
        log.WriteLine($"========================================");
        log.WriteLine($"  NOT RESOLVED — 逐筆原因分析");
        log.WriteLine($"========================================");
        log.WriteLine();

        // 5a. Third-party DLL（Roslyn 解析成功但不是 .NET 原生 → resolved_to = null）
        log.WriteLine($"--- [NOT RESOLVED] Third-party DLL — Roslyn 知道型別，但不是 .NET 原生 ({resolvedToNull.Count}) ---");
        log.WriteLine($"    原因：這些 reference 的 target 定義在第三方 DLL（如 DevExpress），不在 .NET Framework 裡");
        log.WriteLine($"    影響：ref_type 正確（control-read/write 等），但 resolved_to = null");
        log.WriteLine();
        foreach (var g in resolvedToNull.GroupBy(r => r.RefType).OrderByDescending(g => g.Count()))
        {
            log.WriteLine($"  ref_type={g.Key}: {g.Count()} 筆");
            foreach (var r in g.Take(5))
                log.WriteLine($"    target={r.Target}, caller={r.Caller}, file={r.File}:{r.Line}");
            if (g.Count() > 5)
                log.WriteLine($"    ... and {g.Count() - 5} more");
            log.WriteLine();
        }

        // 5b. Unresolved（Roslyn 完全解不了 → ref_type = "unresolved"）
        log.WriteLine($"--- [NOT RESOLVED] Unresolved — Roslyn 無法辨識型別 ({unresolvedRefs.Count}) ---");
        log.WriteLine($"    原因：缺少 DLL、namespace 不對、或原始碼有語法錯誤");
        log.WriteLine($"    影響：ref_type = unresolved，resolved_to = null");
        log.WriteLine();
        foreach (var g in unresolvedRefs.GroupBy(r => r.Target).OrderByDescending(g => g.Count()))
        {
            log.WriteLine($"  target={g.Key} ({g.Count()} 筆)");
            foreach (var r in g.Take(3))
                log.WriteLine($"    caller={r.Caller}, file={r.File}:{r.Line}, context={r.Context[..Math.Min(80, r.Context.Length)]}");
            if (g.Count() > 3)
                log.WriteLine($"    ... and {g.Count() - 3} more");
            log.WriteLine();
        }

        // 6. 每一筆 RESOLVED 的明細
        log.WriteLine($"========================================");
        log.WriteLine($"  RESOLVED — 逐筆明細");
        log.WriteLine($"========================================");
        log.WriteLine();

        // 6a. Source code
        log.WriteLine($"--- [RESOLVED] Source code — 有原始碼位置 ({resolvedToSource.Count}) ---");
        foreach (var r in resolvedToSource)
            log.WriteLine($"  {r.Caller} → {r.Target} | {r.RefType} | resolved_to={r.ResolvedTo}");
        log.WriteLine();

        // 6b. Framework
        log.WriteLine($"--- [RESOLVED] Framework — .NET 原生方法 ({resolvedToFramework.Count}) ---");
        var fwGrouped = resolvedToFramework.GroupBy(r => r.ResolvedTo).OrderByDescending(g => g.Count());
        foreach (var g in fwGrouped)
        {
            log.WriteLine($"  {g.Key} ({g.Count()} 筆)");
            foreach (var r in g.Take(2))
                log.WriteLine($"    caller={r.Caller}, target={r.Target}");
            if (g.Count() > 2)
                log.WriteLine($"    ... and {g.Count() - 2} more");
        }
        log.WriteLine();

        // 7. Missing types 完整清單
        log.WriteLine($"--- Missing types ({missingTypes.Count}) ---");
        foreach (var t in missingTypes)
            log.WriteLine($"  {t}");
        log.WriteLine();

        // 8. Methods 按 owner 分佈
        log.WriteLine($"--- Methods ({methods.Count}) by owner ---");
        var methodsByOwner = methods.GroupBy(m => m.Owner).OrderByDescending(g => g.Count());
        foreach (var g in methodsByOwner)
            log.WriteLine($"  {g.Key}: {g.Count()} methods");
        log.WriteLine();

        // 9. Files 清單
        log.WriteLine($"--- Files ({files.Count}) ---");
        foreach (var f in files)
            log.WriteLine($"  [{f.Role}] {f.Path} — {f.Reason}");
    }

    Console.Error.WriteLine($"       Diagnostic log: {diagPath}");
}
return 0;

static void PrintUsage() => Console.Error.WriteLine(@"
VbAnalyzer — Roslyn-based VB.NET semantic analyzer

Usage:
  dotnet run --project <VbAnalyzer-dir> -- --sln <path.sln> --form <FormName> --output <dir> [--libs <dir>] [--project-root <dir>]

Options:
  --sln <path>          .sln 檔案路徑（掃描 solution 內所有 .vbproj）
  --project <path>      單一 .vbproj 路徑（跟 --sln 擇一）
  --form <name>         要分析的 Form 名稱（如 frmOrder）
  --output <dir>        JSON + MD 輸出目錄
  --libs <dir>          第三方 DLL 目錄（預設 libs/）
  --project-root <dir>  相對路徑基準點（必填，與 Python 的 --project-root 對齊）
  --debug               產出 diagnostic.log（逐筆記錄每個 reference 為什麼 resolved 或沒 resolved）
  --help                顯示用法
");
