# NeoFileMagic

NeoFileMagic 是一個以 .NET 9 建置的檔案解析工具庫，現階段聚焦於 OpenDocument 試算表（ODS）讀取，提供安全、可控資源上限與簡潔的 API，方便在應用程式或服務中快速取得 ODS 中的表格資料。

> 注意：本文件與回應均使用「正體中文」。

## 功能特色
- ODS 讀取：以串流或檔案路徑載入 ODS，產生 `OdsDocument`、`OdsSheet`、`OdsRow`、`OdsCell` 等結構化資料。
- 安全預設：禁用 DTD、停用 `XmlResolver`，預設拒絕加密 ODS（可選擇放行）。
- 資源上限：工作表數、列數、欄數、重複列/欄展開上限、文字空白展開上限皆可設定，避免記憶體壓力。
- 單行化輸出：`Ods.OneLine` 提供多種文字整理策略（保留、跳脫、折疊、取第一段）。
- 輕量與明確：以不可變/唯讀結構為主，利於在服務端環境使用。

## 安裝與建置
- 還原套件：`dotnet restore`
- 建置（Debug）：`dotnet build NeoFileMagic -c Debug`
- 建置（Release）：`dotnet build NeoFileMagic -c Release`
- 產生 NuGet 套件：`dotnet pack NeoFileMagic -c Release`

需求：.NET 9（`net9.0`）。

## 快速上手

```csharp
using NeoFileMagic.FileReader.Ods;

var doc = Ods.Load("sample.ods");
var first = doc.Sheets[0];
var cell = first.GetCell(0, 0);
Console.WriteLine(Ods.OneLine(cell));
```

- 命名空間：`NeoFileMagic.FileReader.Ods`
- 型別要點：
  - `Ods.Load(string path)` / `Ods.Load(Stream stream)` 讀取 ODS。
  - `OdsDocument.Sheets` 取得工作表集合；`OdsSheet.GetCell(row, col)` 取得儲存格。
  - `Ods.OneLine(cell, mode)` 將多行/含控制字元文字整理為單行。

## 讀取選項（OdsReaderOptions）
可依情境調整資源上限與安全行為：

- `ThrowOnEncrypted`（預設 true）：偵測到加密 ODS 時丟出 `NotSupportedException`。
- `MaxSheets`（預設 256）：最大工作表數。
- `MaxRowsPerSheet`（預設 1_000_000）：每張工作表最大列數。
- `MaxColumnsPerRow`（預設 16_384）：每列最大欄數。
- `MaxRepeatedRows`（預設 1_000_000）：展開重複列上限。
- `MaxRepeatedColumns`（預設 16_384）：展開重複欄上限。
- `MaxTextSpaceRun`（預設 100）：`text:s` 連續空白展開上限。
- `LimitMode`（預設 `Truncate`）：達上限的處理策略，`Throw` 或 `Truncate`。
- `CollapseEmptyRepeatedRows`（預設 true）：完全空白的重複列不展開。
- `TrimTrailingEmptyCells`（預設 true）：修剪列尾端連續空白欄位。

範例：
```csharp
var doc = Ods.Load("big.ods", new OdsReaderOptions
{
    ThrowOnEncrypted = true,
    MaxSheets = 64,
    MaxRowsPerSheet = 200_000,
    MaxColumnsPerRow = 1024,
    LimitMode = OdsLimitMode.Truncate,
    CollapseEmptyRepeatedRows = true,
});
```

## 文字輸出策略（TextHandling）
- `Keep`：保持原樣。
- `Escape`：以 `\n`、`\t`、`\\` 表示控制字元。
- `CollapseToSpace`：將換行與定位等折疊為單一空白並去除多餘空白。
- `FirstParagraph`：僅取第一段（遇到首個換行即截斷並修剪）。

## 安全性
- XML 解析：預設禁用 DTD、忽略註解/處理指令/多餘空白、`XmlResolver = null`。
- 加密 ODS：預設拒絕（`ThrowOnEncrypted = true`）。如需放行，請明確設定並評估風險。
- 資源上限：務必依應用規模調整，避免惡意/超大檔造成資源耗盡。

## 專案結構
- 方案：`NeoFileMagic.sln`
- 程式庫：`NeoFileMagic/`（`NeoFileMagic.csproj`）
  - ODS 讀取：`NeoFileMagic/FileReader/Ods/` 下的 `Ods*.cs` 與 `TextHandling.cs`
- 測試：`NeoFileMagic.Tests/`（xUnit, `net9.0`）
  - 本地測試檔：`NeoFileMagic.Tests/sample.ods`（隨專案提供）
  - 外部資料集 URL：`NeoFileMagic.Tests/Data/DatasetUrl.txt`

## 測試
- 執行測試：`dotnet test`
- 預設略過的下載測試可透過環境變數開啟：
  - 設定 `NFM_ALLOW_NET=1` 後，手動取消 Skip 或在本機執行對應測試。

## 開發規範（摘要）
- 風格：4 空白縮排、UTF-8、一類型一檔（類別增長後）
- C#：啟用可空性（`<Nullable>enable</Nullable>`），避免 null 抑制。
- 命名：公開 API 使用 PascalCase；區域變數與參數使用 camelCase。
- 格式化：PR 前請執行 `dotnet format`。

## 版本相容性
目前 API 命名空間為 `NeoFileMagic.FileReader.Ods`。如由舊版本升級，請確認 `using` 或完整限定名稱是否需要調整。

---
歡迎 issue 與 PR，一起讓 NeoFileMagic 變得更好！

