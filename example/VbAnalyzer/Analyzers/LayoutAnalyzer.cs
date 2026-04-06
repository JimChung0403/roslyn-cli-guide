namespace VbAnalyzer.Analyzers;

public static class LayoutAnalyzer
{
    /// <summary>
    /// 從 controls 的 location/size 資訊建立 layout tree。
    /// 邏輯與 Python build_layout() 完全一致：按 parent 分組 → 按 y 分排（10px 容差）→ 按 x 排序。
    /// </summary>
    public static LayoutData Build(List<ControlEntry> controls, string formName)
    {
        // 按 parent 分組（只取有座標的控制項）
        var byParent = new Dictionary<string, List<ControlEntry>>(StringComparer.OrdinalIgnoreCase);

        foreach (var ctrl in controls)
        {
            if (ctrl.LocationX == null || ctrl.LocationY == null) continue;
            var parent = ctrl.Parent ?? formName;
            if (!byParent.ContainsKey(parent))
                byParent[parent] = [];
            byParent[parent].Add(ctrl);
        }

        var containers = new List<ContainerData>();

        foreach (var containerName in byParent.Keys.OrderBy(k => k))
        {
            var entries = byParent[containerName]
                .OrderBy(e => e.LocationY ?? 0)
                .ThenBy(e => e.LocationX ?? 0)
                .ToList();

            // 按 y 分排：y 差距 <= 10px 視為同一排
            var rows = new List<List<ControlEntry>>();
            var currentRow = new List<ControlEntry>();
            int? currentY = null;

            foreach (var entry in entries)
            {
                var y = entry.LocationY ?? 0;
                if (currentY == null || Math.Abs(y - currentY.Value) <= 10)
                {
                    currentRow.Add(entry);
                    currentY ??= y;
                }
                else
                {
                    if (currentRow.Count > 0) rows.Add(currentRow);
                    currentRow = [entry];
                    currentY = y;
                }
            }
            if (currentRow.Count > 0) rows.Add(currentRow);

            // 每排按 x 排序
            var rowDataList = rows.Select(row =>
            {
                var sorted = row.OrderBy(e => e.LocationX ?? 0).ToList();
                return new RowData
                {
                    Y = sorted[0].LocationY ?? 0,
                    Controls = sorted.Select(e => new LayoutControl
                    {
                        Name = e.Name,
                        Text = e.DisplayText ?? "",
                        Type = e.ControlType,
                        X = e.LocationX,
                        Y = e.LocationY,
                        W = e.SizeW,
                        H = e.SizeH
                    }).ToList()
                };
            }).ToList();

            containers.Add(new ContainerData
            {
                Container = containerName,
                Rows = rowDataList
            });
        }

        return new LayoutData { Form = formName, Containers = containers };
    }
}
