# NeoFileMagic

![封面圖](images/cover.png)

提供「安全、輕量、可控資源上限」的多種檔案格式讀取器，
以支援後續應用快速擷取結構化資料（目前先提供 ODS，後續可擴充其他格式）。

## 專案一覽（3 個專案）
- NeoFileMagic（類庫）: 核心檔案讀取框架與公開 API（目前含 ODS 模組）。
  - 重點：安全解析與資源上限控制；現有 ODS 模組提供安全 XML 解析、加密檔預設拒絕、列/欄上限與 `Ods.OneLine` 單行化文字；
    後續可擴充 CSV、XLSX、JSON 等格式模組。
- NeoFileMagic.Tests（測試）: xUnit 測試，驗證解析正確性與安全/上限行為。
  - 內含最小 ODS 測試檔與外部資料集連結（下載測試預設略過）。
- Sample（範例）: 主控台示例，展示如何載入 ODS 與讀取儲存格。

基本使用（C#）：
```csharp
using NeoFileMagic.FileReader.Ods;
var doc = Ods.Load("sample.ods");
var cell = doc.Sheets[0].GetCell(0, 0);
Console.WriteLine(Ods.OneLine(cell));
```

建置/測試：
- 還原/建置：`dotnet restore`、`dotnet build NeoFileMagic -c Debug`
- 測試：`dotnet test`
