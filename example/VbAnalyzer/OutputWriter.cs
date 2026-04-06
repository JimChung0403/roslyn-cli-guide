using System.Text.Json;

namespace VbAnalyzer;

public static class OutputWriter
{
    public static void WriteWrapped<T>(string path, List<T> data, string purpose, string consumers, JsonSerializerOptions options)
    {
        var wrapper = new Dictionary<string, object>
        {
            ["_purpose"] = purpose,
            ["_consumers"] = consumers,
            ["data"] = data
        };
        File.WriteAllText(path, JsonSerializer.Serialize(wrapper, options));
        Console.Error.WriteLine($"       Wrote {path} ({data.Count} entries)");
    }

    public static void WriteLayout(string path, LayoutData layout, string formName, JsonSerializerOptions options)
    {
        var wrapper = new Dictionary<string, object>
        {
            ["_purpose"] = $"{formName} UI 容器層級與控制項位置（按容器分組、按 y 分排、按 x 排序）",
            ["_consumers"] = "07-rewrite-prep, React Agent Team",
            ["form"] = layout.Form,
            ["containers"] = layout.Containers
        };
        File.WriteAllText(path, JsonSerializer.Serialize(wrapper, options));
        Console.Error.WriteLine($"       Wrote {path}");
    }

    public static void WriteJson<T>(string path, T data, JsonSerializerOptions options)
    {
        File.WriteAllText(path, JsonSerializer.Serialize(data, options));
        Console.Error.WriteLine($"       Wrote {path}");
    }
}
