# NeoFileMagic Project Context

## Overview
**NeoFileMagic** is a secure, lightweight, and resource-aware file reader library for .NET, currently specializing in **ODS (OpenDocument Spreadsheet)** format. It is designed to facilitate structured data extraction with features like strict header validation, strong-typed deserialization, and configurable resource limits to prevent denial-of-service attacks via malicious files.

## Project Structure
*   **`NeoFileMagic/`**: The core class library containing the ODS parser, object models, and utilities.
    *   `FileReader/Ods/Core/NeoOds.cs`: Main entry point for loading and processing ODS files.
    *   `FileReader/Ods/Options/OdsReaderOptions.cs`: Configuration for security limits (e.g., max rows, encryption handling).
*   **`NeoFileMagic.Tests/`**: xUnit test project covering parsing logic, exception handling, and dataset validation.
*   **`Sample/`**: A console application demonstrating local/remote file loading and object deserialization.

## Key Technologies
*   **Language**: C# (Target Framework: `net10.0`)
*   **Dependencies**: 
    *   `Newtonsoft.Json`: Used for object mapping and deserialization.
    *   `System.IO.Compression`: For handling ODS zip structure.
*   **Testing**: `xUnit`, `coverlet.collector`.

## Security & Resource Limits
防範惡意 ODS 的阻斷服務 (DoS) 攻擊，開發須遵守以下防禦設計：
1. **資源限制優先**：解析邏輯須遵循 `OdsReaderOptions` 限制（如 `MaxRows`、記憶體緩衝與解壓縮限制）。
2. **禁止繞過防禦**：禁止硬編碼繞過核心解析器安全設定，且未經 `OdsReaderOptions` 驗證不得讀取完整內容。
3. **第三方依賴約束**：未經人類明確授權，禁止新增 any NuGet 依賴，以維護庫的輕量化與安全性。

## Development & Usage

### Building
```bash
dotnet restore
dotnet build
```

### Running Tests
```bash
dotnet test
```

### Code Formatting
提交程式碼前，執行以下指令驗證風格：
```bash
# 自動格式化
dotnet format

# 驗證格式規範（不修改檔案）
dotnet format --verify-no-changes
```

### Running the Sample
```bash
dotnet run --project Sample
```

## Coding Conventions & Patterns
*   **Language Features**: Utilizes modern C# features (file-scoped namespaces, global usings, nullable reference types).
*   **Comments**: Code comments and documentation are primarily in **Traditional Chinese**.
*   **Error Handling**: 
    *   Custom exceptions (e.g., `OdsHeaderMismatchException`, `OdsRowConversionException`) are used for precise error reporting during deserialization.
    *   Strict validation modes are available to enforce header order and count.
*   **Async/Await**: Extensive use of asynchronous patterns, especially for network operations (`LoadFromUrlAsync`).
*   **Data Mapping**: Relies on `JsonProperty` attributes to map spreadsheet columns to C# object properties.

## Common Tasks
*   **Loading a File**: Use `NeoOds.Load(path)` or `NeoOds.Load(stream)`.
*   **Remote Loading**: Use `NeoOds.LoadFromUrlAsync(url)` which includes logic for buffering large files to disk.
*   **Deserialization**: Use `NeoOds.DeserializeSheetOrThrow<T>(sheet)` to convert spreadsheet rows into strongly-typed objects with validation.

## Response Guidelines

- 所有的回答和git commit 訊息都要回傳繁體中文。

