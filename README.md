# NeoFileMagic

![封面圖](https://raw.githubusercontent.com/Benknightdark/NeoFileMagic/refs/heads/master/images/cover.png)

安全、輕量、資源可控的 .NET 多格式檔案讀取器（目前支援 ODS）。適用於需要嚴格資源限制與資料安全驗證的結構化資料提取場景。

## 專案架構
- **NeoFileMagic**（類別庫）: 核心檔案讀取框架與公開 API（內含 ODS 模組）。
- **NeoFileMagic.Tests**（測試）: xUnit 測試專案，驗證解析正確性與資源限制邊界。
- **Sample**（範例）: 示範如何從本機或遠端載入 ODS 並進行強型別反序列化。

## 安裝
```bash
dotnet add package NeoFileMagic
```

## 基本使用
```csharp
using NeoFileMagic.FileReader.Ods;

// 從檔案載入 ODS
var doc = NeoOds.Load("sample.ods");

// 讀取第一個工作表的 (0,0) 儲存格
var cell = doc.Sheets[0].GetCell(0, 0);
Console.WriteLine(NeoOds.OneLine(cell));
```

## 使用範例

### 1) 遍歷列與欄
```csharp
using NeoFileMagic.FileReader.Ods;

var doc = NeoOds.Load("sample.ods");
var sheet = doc.Sheets[0];

for (int r = 0; r < sheet.RowCount; r++)
{
    var row = sheet.Rows[r];
    for (int c = 0; c < row.ColumnCount; c++)
    {
        var cell = row.Cells[c];
        // 以單行輸出：將換行與 Tab 摺疊為單一空白
        Console.Write(NeoOds.OneLine(cell));
        Console.Write('\t');
    }
    Console.WriteLine();
}
```

### 2) 安全與資源限制設定
```csharp
using NeoFileMagic.FileReader.Ods;

var options = new OdsReaderOptions
{
    // 偵測到檔案加密時拋出 NotSupportedException
    ThrowOnEncrypted = true,

    // 資源限制：避免異常或惡意檔案造成記憶體壓力 (DoS 攻擊防禦)
    MaxSheets = 64,
    MaxRowsPerSheet = 100_000,
    MaxColumnsPerRow = 256,
    MaxRepeatedRows = 100_000,
    MaxRepeatedColumns = 256,
};

var doc = NeoOds.Load("sample.ods", options);
```

### 3) 強型別模型反序列化
嚴格依表頭（或 `[JsonPropertyName]`）對應欄位，欄位順序、缺漏或格式錯誤會拋出具體例外。反序列化內部使用高效能的 `System.Text.Json` 進行。
```csharp
using System.Text.Json.Serialization;
using NeoFileMagic.FileReader.Ods;

public sealed class Person
{
    [JsonPropertyName("姓名")]
    [JsonPropertyOrder(0)]
    public string Name { get; set; } = string.Empty;

    [JsonPropertyName("年齡")]
    [JsonPropertyOrder(1)]
    public int Age { get; set; }
}

var doc = NeoOds.Load("people.ods");
var sheet = doc.Sheets[0];
var list = NeoOds.DeserializeSheetOrThrow<Person>(sheet);
// list 為反序列化結果。若欄位順序或型別不符會拋出 Ods* 相關例外。
```

### 4) 從遠端 URL 載入（非同步）
```csharp
using NeoFileMagic.FileReader.Ods;

var doc = await NeoOds.LoadFromUrlAsync("https://example.com/data.ods");
```

## 開發與驗證

### 1) 建置與測試
```bash
dotnet restore
dotnet build
dotnet test
```

### 2) 程式碼格式化
```bash
# 自動套用專案程式碼風格
dotnet format

# 驗證格式（不修改檔案）
dotnet format --verify-no-changes
```
