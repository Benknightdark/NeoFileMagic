# NeoFileMagic

![封面圖](https://raw.githubusercontent.com/Benknightdark/NeoFileMagic/refs/heads/master/images/cover.png)

提供「安全、輕量、可控資源上限」的多種檔案格式讀取器，
以支援後續應用快速擷取結構化資料（目前先提供 ODS，後續可擴充其他格式）。

## 專案一覽（3 個專案）
- NeoFileMagic（類庫）: 核心檔案讀取框架與公開 API（目前含 ODS 模組）。
  - 重點：安全解析與資源上限控制；現有 ODS 模組提供安全 XML 解析、列/欄上限與 `Ods.OneLine` 單行化文字；
    後續可擴充 CSV、XLSX、JSON 等格式模組。
- NeoFileMagic.Tests（測試）: xUnit 測試，驗證解析正確性與安全/上限行為。
  - 內含最小 ODS 測試檔與外部資料集連結（下載測試預設略過）。
- Sample（範例）: 範例程式，展示如何載入 ODS 與讀取儲存格。

## 安裝
- NuGet：`dotnet add package NeoFileMagic`

## 基本使用（C#）
```csharp
using NeoFileMagic.FileReader.Ods;

// 從檔案載入 ODS
var doc = Ods.Load("sample.ods");

// 讀取第一個工作表的 (0,0) 儲存格
var cell = doc.Sheets[0].GetCell(0, 0);
Console.WriteLine(Ods.OneLine(cell));
```

## 使用範例

### 1) 遍歷列與欄
```csharp
using NeoFileMagic.FileReader.Ods;

var doc = Ods.Load("sample.ods");
var sheet = doc.Sheets[0];

for (int r = 0; r < sheet.RowCount; r++)
{
    var row = sheet.Rows[r];
    for (int c = 0; c < row.ColumnCount; c++)
    {
        var cell = row.Cells[c];
        // 以單行輸出：換行/Tab 摺疊為空白
        Console.Write(Ods.OneLine(cell));
        Console.Write('\t');
    }
    Console.WriteLine();
}
```

### 2) 安全與資源上限設定
```csharp
using NeoFileMagic.FileReader.Ods;

var options = new OdsReaderOptions
{
    // 預設 true：若檔案加密則丟出 NotSupportedException
    ThrowOnEncrypted = true,

    // 控制上限，避免惡意/異常檔案造成記憶體壓力
    MaxSheets = 64,
    MaxRowsPerSheet = 100_000,
    MaxColumnsPerRow = 256,
    MaxRepeatedRows = 100_000,
    MaxRepeatedColumns = 256,
};

var doc = Ods.Load("sample.ods", options);
```

### 3) 以強型別模型反序列化工作表
嚴格依表頭（或 `[JsonPropertyName]`）對應欄位，欄位順序/缺漏或格式錯誤會拋出具體例外。
```csharp
using System.Text.Json.Serialization;
using NeoFileMagic.FileReader.Ods;

public sealed class Person
{
    [JsonPropertyName("Name")] public string Name { get; set; } = string.Empty;
    [JsonPropertyName("Age")]  public int Age  { get; set; }
}

var doc = Ods.Load("people.ods");
var sheet = doc.Sheets[0];
var list = Ods.DeserializeSheetOrThrow<Person>(sheet);
// list 為強型別結果，若欄位/資料不符會拋出 Ods* 相關例外
```

### 4) 從 URL 讀取（非同步）
```csharp
using NeoFileMagic.FileReader.Ods;

var doc = await Ods.LoadFromUrlAsync("https://example.com/data.ods");
```

建置/測試：
- 還原/建置：`dotnet restore`、`dotnet build NeoFileMagic -c Debug`
- 測試：`dotnet test`
