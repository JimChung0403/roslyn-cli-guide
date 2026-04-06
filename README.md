# Roslyn CLI for VB.NET 語義分析 — Linux SOP

## 目標

在 Linux (Ubuntu) 上用 Roslyn 對 VB.NET WinForms 原始碼做 compiler 級語義分析，輸出 JSON 索引檔。解決 Python regex 無法處理的 5 個盲區：

| 盲區 | 說明 | Roslyn 解法 |
|------|------|------------|
| 跨檔呼叫 | `svc.Save()` → 不知道定義在哪 | `GetSymbolInfo()` 回傳定義位置 |
| Partial class | 分散在多個 .vb 檔 | `Compilation` 自動合併 |
| 繼承鏈 | `MyBase.OnLoad(e)` → 不知道父類 | `GetSymbolInfo()` 解析繼承 |
| With block | `.Columns.Add` → 不知道主體 | 語義模型知道 `.` 前的型別 |
| 事件掛載 | `Handles btnSave.Click` → 靠猜 | `HandlesClause` syntax node |

---

## 快速上手

整個流程分兩階段：**建置**（一次性）和**執行**（每個 Form 跑一次）。

### 階段一：建置（一次性，之後不用再跑）

```bash
# 1. 安裝 .NET SDK（任選一種，Snap 最方便）
sudo snap install dotnet-sdk --classic --channel=9.0

# 2. 複製 VbAnalyzer 到任意位置
#    VbAnalyzer 是獨立的分析工具，跟你的 VB 專案完全無關，放哪裡都行
cp -r /path/to/roslyn-cli-guide/example/VbAnalyzer ~/tools/VbAnalyzer

# 3. 還原套件 + 編譯（需要網路下載 NuGet 套件，約 60MB，只需跑一次）
cd ~/tools/VbAnalyzer
dotnet restore
dotnet build
```

完成後 `~/tools/VbAnalyzer/` 就是一個可用的工具，不需要再動。

### 階段二：執行（每個 Form 跑一次）

```bash
# 分析一個 Form
dotnet run --project ~/tools/VbAnalyzer/ -- \
  --sln "/path/to/YourSolution.sln" \
  --form "frmXXX" \
  --output "./output"

# 看結果
cat ./output/stats.json
ls ./output/
```

如果專案用了第三方元件（DevExpress、FarPoint 等），加 `--libs` 提高解析率：

```bash
dotnet run --project ~/tools/VbAnalyzer/ -- \
  --sln "/path/to/YourSolution.sln" \
  --form "frmXXX" \
  --output "./output" \
  --libs "/path/to/libs/"
```

第三方 DLL 從 Windows 機器複製，詳見 [windows-copy-guide.md](./windows-copy-guide.md)。

### 全部參數

| 參數 | 必填 | 說明 |
|------|------|------|
| `--project` | 是 | VbAnalyzer 工具的位置（`dotnet run` 的參數，不是 VB 專案） |
| `--sln` | 二選一 | .sln 檔案路徑，掃描 solution 內所有 .vbproj |
| `--project`（`--` 後面的） | 二選一 | 單一 .vbproj 路徑（跟 `--sln` 擇一使用） |
| `--form` | 是 | 要分析的 Form 名稱（如 `frmOrder`） |
| `--output` | 是 | JSON 輸出目錄 |
| `--libs` | 否 | 第三方 DLL 目錄，沒有就不加，缺的 symbol 標 unresolved |
| `--help` | 否 | 顯示用法 |

`--project`（`--` 前面）指向 VbAnalyzer 工具，`--sln` 指向你的 VB 專案。兩者完全獨立，可以在不同目錄、不同磁碟。

### 產出檔案

```
output/
├── methods.json        ← 方法定義、起訖行、參數
├── controls.json       ← 控制項名稱、型別、宣告位置
├── events.json         ← 事件 handler、掛載方式、位置
├── references.json     ← 呼叫關係、控制項讀寫、跨檔解析
├── files.json          ← 檔案關聯、partial class
└── stats.json          ← 統計：解析成功率、缺失型別
```

### 驗證

- `stats.json` 的 `resolved_rate` > 80% → 成功
- `resolved_rate` < 50% → 缺第三方 DLL，加 `--libs ./libs/` 重跑。見 [windows-copy-guide.md](./windows-copy-guide.md)

### 分析多個 Form

階段一只跑一次。之後每個 Form 只要重複階段二，換 `--form` 和 `--output`：

```bash
dotnet run --project ~/tools/VbAnalyzer/ -- --sln "/path/to/Solution.sln" --form "frmOrder" --output "./output/frmOrder"
dotnet run --project ~/tools/VbAnalyzer/ -- --sln "/path/to/Solution.sln" --form "frmCustomer" --output "./output/frmCustomer"
dotnet run --project ~/tools/VbAnalyzer/ -- --sln "/path/to/Solution.sln" --form "frmInvoice" --output "./output/frmInvoice"
```

### 實測結果（telausuario 專案，frmBackupAutomatico）

```
$ dotnet run --project ~/tools/VbAnalyzer/ -- --sln "telausuario.sln" --form "frmBackupAutomatico" --output "./output"

[1/5] Collecting .vb files...
       Found 71 .vb files
[2/5] Building compilation...
       .NET Framework 4.8 refs: 237 assemblies
       Root namespace: 'telausuario'
       Global imports: 11 (Microsoft.VisualBasic, System, System.Data, ...)
[3/5] Analyzing form 'frmBackupAutomatico'...
       Methods: 21, Controls: 16, Events: 5, References: 72
[5/5] Done. Resolved rate: 51.4%
```

51.4% 是因為缺 DevExpress DLL。加上後預估 85-95%。

跨檔呼叫解析範例（regex 做不到的）：

```json
{
  "caller": "btnDownload_Click",
  "target": "clsDropBoxFuncao.BackupEspecifico",
  "resolved_to": "telausuario/Classes/clsDropBoxFuncao.vb:646"
}
```

產出的跨檔呼叫解析範例：

```json
{
  "caller": "btnDownload_Click",
  "target": "clsDropBoxFuncao.BackupEspecifico",
  "target_file": "Classes/clsDropBoxFuncao.vb",
  "target_line": 646,
  "resolved": true
}
```

這是 Python regex 做不到的——Roslyn 把 `BackupEspecifico()` 解析到 `clsDropBoxFuncao.vb` 第 646 行。

---

## Step 1：安裝 .NET SDK（一次性）

先確認有沒有裝過：

```bash
dotnet --version
# 有顯示版本號 → 跳到 Step 2
```

以下提供 5 種安裝方式，任選一種成功即可。

### 方式 A：Snap（推薦，公司環境常用）

```bash
sudo snap install dotnet-sdk --classic --channel=9.0
```

Snap 安裝的 dotnet 路徑在 `/snap/dotnet-sdk/current/dotnet`，通常 snap 會自動加到 PATH。如果 `dotnet --version` 找不到指令：

```bash
# 手動加 PATH
echo 'export PATH=$PATH:/snap/dotnet-sdk/current' >> ~/.bashrc
source ~/.bashrc
```

如果 9.0 channel 不可用，改用 8.0（同樣支援 Roslyn）：

```bash
sudo snap install dotnet-sdk --classic --channel=8.0
```

### 方式 B：apt-get

```bash
sudo apt-get update
sudo apt-get install -y dotnet-sdk-9.0
```

如果 apt 找不到套件，手動加 Microsoft 套件源：

```bash
wget https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
sudo apt-get update
sudo apt-get install -y dotnet-sdk-9.0
```

### 方式 C：官方安裝腳本（不需要 sudo）

適合公司機器沒有 root 權限的情況，裝在 home 目錄下：

```bash
wget https://dot.net/v1/dotnet-install.sh -O dotnet-install.sh
chmod +x dotnet-install.sh

# 裝到 ~/.dotnet
./dotnet-install.sh --channel 9.0

# 加到 PATH
echo 'export DOTNET_ROOT=$HOME/.dotnet' >> ~/.bashrc
echo 'export PATH=$PATH:$DOTNET_ROOT' >> ~/.bashrc
source ~/.bashrc
```

### 方式 D：手動下載 tarball（離線安裝）

公司網路受限時，在外面先下載好帶進去：

1. 到 https://dotnet.microsoft.com/download/dotnet/9.0 下載 **Linux x64 SDK** 的 `.tar.gz`
2. 複製到公司機器，解壓：

```bash
mkdir -p ~/.dotnet
tar -xzf dotnet-sdk-9.0.*-linux-x64.tar.gz -C ~/.dotnet

echo 'export DOTNET_ROOT=$HOME/.dotnet' >> ~/.bashrc
echo 'export PATH=$PATH:$DOTNET_ROOT' >> ~/.bashrc
source ~/.bashrc
```

### 方式 E：改用 .NET 8.0（9.0 裝不上時的備案）

.NET 8.0 是 LTS 版本，套件源通常比較齊全。Roslyn 在 8.0 上完全正常：

```bash
# apt
sudo apt-get install -y dotnet-sdk-8.0

# 或 snap
sudo snap install dotnet-sdk --classic --channel=8.0

# 或安裝腳本
./dotnet-install.sh --channel 8.0
```

用 8.0 的話，需要改一行 `.csproj`（已在範例中標註）：

```xml
<!-- VbAnalyzer.csproj 中，把這行 -->
<TargetFramework>net9.0</TargetFramework>
<!-- 改成 -->
<TargetFramework>net8.0</TargetFramework>
```

### 驗證安裝

無論用哪種方式，最終確認：

```bash
dotnet --version
# 顯示 9.0.x 或 8.0.x 即成功
```

---

## Step 2：複製 VbAnalyzer 到任意位置

VbAnalyzer 是獨立的分析工具，跟你的 VB 專案完全無關。放哪裡都行：

```bash
cp -r /path/to/roslyn-cli-guide/example/VbAnalyzer ~/tools/VbAnalyzer
```

---

## Step 3：還原 NuGet 套件

```bash
cd ~/tools/VbAnalyzer
dotnet restore
```

這一步從 `https://api.nuget.org/` 下載三個套件（約 60MB，只需跑一次）：
- `Microsoft.CodeAnalysis.VisualBasic`（Roslyn VB 編譯器，~15MB）
- `Microsoft.NETFramework.ReferenceAssemblies.net48`（.NET Framework 4.8 型別定義，~40MB）
- `System.Text.Json`（JSON 輸出，~3MB）

下載完後快取在 `~/.nuget/packages/`，之後不再需要網路。

### 如果公司封鎖了 nuget.org（資安限制）

#### 方式 A：在外面先 restore，把快取帶進去

在有網路的機器上（家裡、咖啡廳）：

```bash
cd ~/tools/VbAnalyzer
dotnet restore
```

然後把兩個東西複製到公司機器：
1. `~/tools/VbAnalyzer/` 整個目錄（含 restore 產生的 `obj/`）
2. `~/.nuget/packages/` 整個目錄（NuGet 快取，約 60MB）

到公司後直接跑 `dotnet build`，不需要再 restore。

#### 方式 B：離線套件包

在有網路的機器上，把套件下載到一個資料夾：

```bash
cd ~/tools/VbAnalyzer
dotnet restore --packages ./nuget-local
```

把 `nuget-local/` 資料夾複製到公司機器，然後指定本地套件源：

```bash
cd ~/tools/VbAnalyzer
dotnet restore --source ./nuget-local
```

#### 方式 C：公司有內部 NuGet 源（Artifactory、Nexus、Azure Artifacts 等）

問公司的 DevOps 團隊拿到內部 NuGet 源的 URL，設定後 restore：

```bash
# 加入公司內部源
dotnet nuget add source https://your-company-artifactory/nuget/ -n company

# restore 會自動從公司源找
dotnet restore
```

#### 方式 D：手動下載 .nupkg 檔案

在有網路的機器上手動下載三個 `.nupkg` 檔：

1. https://www.nuget.org/packages/Microsoft.CodeAnalysis.VisualBasic/4.12.0 → 點 Download package
2. https://www.nuget.org/packages/Microsoft.NETFramework.ReferenceAssemblies.net48/1.0.3 → 點 Download package
3. https://www.nuget.org/packages/System.Text.Json/9.0.0 → 點 Download package

每個套件頁面可能還有依賴套件，也要一起下載（`Microsoft.CodeAnalysis.Common`、`System.Collections.Immutable` 等）。

把所有 `.nupkg` 檔放到一個資料夾，帶到公司：

```bash
# 指定本地資料夾作為套件源
dotnet restore --source /path/to/nupkg-folder/
```

這個方式比較麻煩（要手動處理依賴），優先用方式 A 或 B。

---

## Step 4：準備第三方 DLL（可選）

如果你的 VB.NET 專案用了第三方元件（AxFPSpread、DevExpress 等），需要從 Windows 機器複製 interop assembly。

詳細步驟見 [windows-copy-guide.md](./windows-copy-guide.md)。

**沒有第三方 DLL 也能跑**——Roslyn 會正常分析大部分程式碼，只是第三方元件相關的 symbol 會標為 `unresolved`。

---

## Step 5：編譯

```bash
cd ~/tools/VbAnalyzer
dotnet build
```

第一次編譯可能需要 1-2 分鐘（下載依賴）。之後秒編。

如果看到 warning 但沒有 error，就是成功了。

---

## Step 6：執行

### 指定 .sln 掃描整個 solution

```bash
dotnet run --project ~/tools/VbAnalyzer/ -- \
  --sln "/path/to/YourSolution.sln" \
  --form "frmOrder" \
  --output "/path/to/output/lsp-index"
```

### 指定單一 .vbproj

```bash
dotnet run --project ~/tools/VbAnalyzer/ -- \
  --project "/path/to/YourProject.vbproj" \
  --form "frmOrder" \
  --output "/path/to/output/lsp-index"
```

### 帶第三方 DLL

```bash
dotnet run --project ~/tools/VbAnalyzer/ -- \
  --sln "/path/to/YourSolution.sln" \
  --form "frmOrder" \
  --output "/path/to/output/lsp-index" \
  --libs "/path/to/libs/"
```

---

## Step 7：檢查輸出

執行成功後，`--output` 目錄下會有：

```
lsp-index/
├── methods.json        ← 方法定義、起訖行、參數
├── controls.json       ← 控制項名稱、型別、宣告位置
├── events.json         ← 事件 handler、掛載方式、位置
├── references.json     ← 呼叫關係、控制項讀寫、RPC
├── files.json          ← 檔案關聯、partial class
└── stats.json          ← 統計：解析成功/失敗數量
```

### 驗證要點

1. `stats.json` 的 `resolved_rate` 大於 80% 就算成功
2. `references.json` 裡 `resolved: false` 的項目檢查是否都是第三方元件
3. 如果 `resolved_rate` 低於 50%，大概率是缺少 reference assembly — 見 [windows-copy-guide.md](./windows-copy-guide.md)

---

## 常見問題

### Q: `dotnet restore` 失敗，說找不到套件

```bash
# 確認 NuGet 源有設定
dotnet nuget list source
# 應該有 https://api.nuget.org/v3/index.json
# 如果沒有：
dotnet nuget add source https://api.nuget.org/v3/index.json -n nuget.org
```

### Q: 編譯成功但 `resolved_rate` 很低

缺 reference assembly。把 stderr 輸出的 `[warn] N compilation errors` 導出來看：

```bash
dotnet run --project ~/tools/VbAnalyzer/ -- \
  --sln "..." --form "..." --output "..." 2> errors.log

# 看缺什麼型別
grep "BC30002" errors.log | head -20
# BC30002 = Type 'xxx' is not defined
```

把缺的 DLL 從 Windows 複製到 `--libs` 目錄。

### Q: 很多 .vb 檔用的是 Big5 / Shift-JIS 編碼

目前程式用 `File.ReadAllText()` 預設讀 UTF-8。如果你的 .vb 檔是其他編碼，修改 `CompilationBuilder.cs` 裡的讀檔方式：

```csharp
// 改成：
var encoding = Encoding.GetEncoding(950); // Big5
var text = File.ReadAllText(f, encoding);
```

### Q: 跑完要多久？

| Solution 規模 | 大約時間 |
|------|------|
| 50 個 .vb 檔 | 2-5 秒 |
| 200 個 .vb 檔 | 5-15 秒 |
| 500+ 個 .vb 檔 | 15-30 秒 |

Roslyn 是 in-memory 編譯，不寫 binary 不跑 codegen，所以很快。
