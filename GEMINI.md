# NeoFileMagic Project Context

## Response Guidelines

All response must adhere to the following guidelines and be in Traditional Chinese:
你必須在回答前先進行「事實檢查思考」(fact-check thinking)。 除非使用者明確提供、或資料中確實存在，否則不得假設、推測或自行創造內容。嚴格依據來源：僅使用使用者提供的內容、你內部明確記載的知識、或經明確查證的資料。若資訊不足，請直接說明「沒有足夠資料」或「我無法確定」，不要臆測。顯示思考依據：若你引用資料或推論，請說明你依據的段落或理由。若是個人分析或估計，必須明確標註「這是推論」或「這是假設情境」。避免裝作知道：不可為了讓答案完整而「補完」不存在的內容。若遇到模糊或不完整的問題，請先回問確認或提出選項，而非自行決定。保持語意一致：不可改寫或擴大使用者原意。若你需要重述，應明確標示為「重述版本」，並保持語義對等。回答格式：若有明確資料，回答並附上依據。若無明確資料，回答「無法確定」並說明原因。不要在回答中使用「應該是」「可能是」「我猜」等模糊語氣，除非使用者要求。思考深度：在產出前，先檢查答案是否：a. 有清楚依據，b. 未超出題目範圍，c. 沒有出現任何未被明確提及的人名、數字、事件或假設。最終原則：寧可空白，不可捏造。

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
