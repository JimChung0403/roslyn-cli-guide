using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbAnalyzer;

public static class OutputWriter
{
    /// <summary>
    /// 寫入帶外層包裝的 JSON，格式與 Python build_vb_form_index.py 一致：
    /// { "_purpose": "...", "_consumers": "...", "data": [...] }
    /// </summary>
    public static void WriteWrapped<T>(string path, List<T> data, string purpose, string consumers, JsonSerializerOptions options)
    {
        // 手動組裝，因為 _purpose 開頭是底線，snake_case 轉換無法處理
        var wrapper = new Dictionary<string, object>
        {
            ["_purpose"] = purpose,
            ["_consumers"] = consumers,
            ["data"] = data
        };
        var json = JsonSerializer.Serialize(wrapper, options);
        File.WriteAllText(path, json);
        Console.Error.WriteLine($"       Wrote {path} ({data.Count} entries)");
    }

    public static void WriteJson<T>(string path, T data, JsonSerializerOptions options)
    {
        var json = JsonSerializer.Serialize(data, options);
        File.WriteAllText(path, json);
        Console.Error.WriteLine($"       Wrote {path}");
    }
}
