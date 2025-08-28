# NeoFileMagic

目標：提供「安全、輕量、可控資源上限」的 ODS（OpenDocument Spreadsheet）讀取器，方便在服務或應用中快速擷取表格資料。

## 專案一覽（3 個專案）
- NeoFileMagic（類庫）: 核心 ODS 讀取器與公開 API。
  - 重點：安全 XML 解析、加密檔預設拒絕、列/欄/重複列欄展開上限、`Ods.OneLine` 單行化文字。
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
