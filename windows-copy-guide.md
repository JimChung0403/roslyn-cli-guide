# Windows 機器複製指南

## 為什麼需要從 Windows 複製東西

Roslyn 在 Linux 上做語義分析時，需要知道每個型別的定義。.NET Framework 標準庫（`System.Windows.Forms.Button` 等）已由 NuGet 套件 `Microsoft.NETFramework.ReferenceAssemblies.net48` 提供，不需要 Windows。

但以下型別 **NuGet 套件沒有**，必須從 Windows 複製：

| 類別 | 範例 | 缺少時的影響 |
|------|------|------------|
| 第三方 COM/ActiveX | AxFPSpread、AxMSFlexGrid | 該控制項的所有屬性、方法、事件都解不了 |
| 第三方 .NET 元件 | DevExpress、Telerik、Infragistics | 同上 |
| 專案內的其他 DLL | SharedLibrary.dll、CommonUtils.dll | 跨 project 的呼叫解不了 |
| COM Interop 產生的 DLL | Interop.ADODB.dll、Interop.Excel.dll | COM 呼叫解不了 |

**如果你的專案沒用第三方元件，跳過本文件，直接跑。**

---

## 要複製什麼

### 情況 1：專案有 `packages/` 或 `lib/` 目錄

很多舊 VB.NET 專案把第三方 DLL 直接放在 repo 裡。先檢查：

```
YourSolution/
├── packages/           ← NuGet packages（舊格式 packages.config）
│   ├── FarPoint.Spread.8.0/
│   │   └── lib/
│   │       └── net40/
│   │           └── FarPoint.Spread.dll
│   └── ...
├── lib/                ← 手動放的第三方 DLL
│   ├── AxInterop.FPSpreadADO.dll
│   └── Interop.ADODB.dll
└── YourProject/
    └── YourProject.vbproj
```

如果有，**直接把這些 DLL 複製到 Linux 的 `libs/` 目錄**：

```bash
# 在 Linux 上
mkdir -p ~/roslyn-analyzer/libs

# 從 Windows 用 scp / USB / 網路磁碟複製
scp user@windows-machine:/path/to/packages/FarPoint.Spread.8.0/lib/net40/*.dll ~/roslyn-analyzer/libs/
scp user@windows-machine:/path/to/lib/*.dll ~/roslyn-analyzer/libs/
```

### 情況 2：DLL 在 GAC 或 COM 註冊目錄

如果 DLL 不在 repo 裡，要從 Windows 機器的以下位置找：

#### COM/ActiveX Interop DLL

在 Windows 上開 cmd/PowerShell：

```powershell
# 找 FarPoint 相關
dir /s C:\Windows\assembly\*FarPoint*
dir /s C:\Windows\Microsoft.NET\assembly\*FarPoint*

# 或者看專案的 bin 目錄（編譯過的話 DLL 都在這）
dir /s C:\path\to\YourProject\bin\Debug\*.dll
dir /s C:\path\to\YourProject\bin\Release\*.dll
```

**最快的方式：直接從 `bin/Debug/` 或 `bin/Release/` 複製所有 DLL。** 裡面包含了專案編譯時實際用到的所有 reference。

```powershell
# Windows 上，打包 bin 目錄的所有 DLL
cd C:\path\to\YourProject\bin\Debug
tar -czf project-dlls.tar.gz *.dll
```

```bash
# Linux 上，解壓到 libs/
scp user@windows-machine:/path/to/project-dlls.tar.gz ~/roslyn-analyzer/
cd ~/roslyn-analyzer
tar -xzf project-dlls.tar.gz -C libs/
```

#### 專案引用的其他 project DLL

如果 solution 裡有多個 project 互相引用（如 `SharedLibrary.vbproj`），那些 project 的 DLL 也要複製。

查看方式：在 Windows 上用文字編輯器打開 `.vbproj`，找 `<Reference>` 和 `<ProjectReference>`：

```xml
<!-- 第三方 DLL reference -->
<Reference Include="FarPoint.Spread">
  <HintPath>..\packages\FarPoint.Spread.8.0\lib\net40\FarPoint.Spread.dll</HintPath>
</Reference>

<!-- 專案內 reference -->
<ProjectReference Include="..\SharedLibrary\SharedLibrary.vbproj" />
```

`<Reference>` 的 `HintPath` 告訴你 DLL 在哪。`<ProjectReference>` 的 DLL 在該 project 的 `bin/` 下。

---

## 複製到 Linux 後的目錄結構

```
~/roslyn-analyzer/
├── VbAnalyzer/          ← Roslyn CLI 專案
└── libs/                ← 所有從 Windows 複製來的 DLL
    ├── FarPoint.Spread.dll
    ├── FarPoint.Win.dll
    ├── AxInterop.FPSpreadADO.dll
    ├── Interop.ADODB.dll
    ├── SharedLibrary.dll
    └── ...
```

執行時加 `--libs` 參數：

```bash
dotnet run --project ~/roslyn-analyzer/VbAnalyzer/ -- \
  --sln "/path/to/solution.sln" \
  --form "frmOrder" \
  --output "/path/to/output/" \
  --libs ~/roslyn-analyzer/libs/
```

---

## 驗證複製是否足夠

跑完後看 stderr 和 `stats.json`：

```bash
# 導出 compilation errors
dotnet run --project ~/roslyn-analyzer/VbAnalyzer/ -- ... 2> errors.log

# 看還缺什麼型別
grep "BC30002" errors.log | sort -u
# BC30002: Type 'XxxType' is not defined.
```

如果 `BC30002` 列出的型別都是你不關心的（如 `AxMSChart`），可以忽略。Roslyn 不會因為缺一個型別就整個失敗，只是該型別相關的 symbol 標 `unresolved`。

---

## 不需要複製的東西

| 項目 | 原因 |
|------|------|
| .NET Framework 本身 | NuGet 套件 `Microsoft.NETFramework.ReferenceAssemblies.net48` 已提供 |
| Visual Studio | Roslyn 是獨立套件，不需要 IDE |
| Windows SDK | 不需要 |
| MSBuild | 我們用 AdhocWorkspace，不走 MSBuild |
| NuGet.exe | `dotnet restore` 處理 |
