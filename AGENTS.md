# Repository Guidelines

## All response must be in 「正體中文」

## Project Structure & Module Organization
- Root solution: `NeoFileMagic.sln`.
- Library: `NeoFileMagic/` contains `NeoFileMagic.csproj` (net9.0, implicit usings, nullable) and source (currently `Class1.cs` with the ODS reader types).
- Generated artifacts: `obj/` folders under root and `NeoFileMagic/` are build outputs — do not modify or commit.

## Build, Test, and Development Commands
- `dotnet restore` — restore NuGet packages.
- `dotnet build NeoFileMagic -c Debug` — fast local build; use `-c Release` for packaging.
- `dotnet pack NeoFileMagic -c Release` — create a NuGet package when needed.
- `dotnet test` — runs tests once a `NeoFileMagic.Tests` project exists.

## Coding Style & Naming Conventions
- Indentation: 4 spaces; UTF-8; one type per file when code grows beyond the current single-file layout.
- C#/.NET: prefer `sealed` classes and `readonly struct` where appropriate; keep `<Nullable>enable</Nullable>` and avoid null suppression.
- Naming: PascalCase for public types/members (`OdsDocument`, `GetCell`); camelCase for locals/parameters; meaningful names over abbreviations.
- Usings: implicit usings are enabled; order `System.*` before others.
- Formatting: run `dotnet format` before opening a PR.

## Testing Guidelines
- Framework: xUnit recommended. Create `NeoFileMagic.Tests/NeoFileMagic.Tests.csproj` targeting `net9.0` and reference `NeoFileMagic`.
- Naming: `ClassName_Method_Scenario_Expected()` (e.g., `OdsXml_ParseContent_WithRepeatedRows_ClampsToLimit`).
- Coverage: prioritize parsing paths, option limits, and security behavior; target ≥80% where feasible. Run `dotnet test` (optionally with coverlet for coverage).

## Commit & Pull Request Guidelines
- Commits: short, imperative subject (e.g., "Add ODS date parsing"); small, focused diffs; reference issues (`#123`).
- PRs: include a clear description, linked issues, verification steps, and tests. Provide minimal input files (e.g., tiny `content.xml`) to demonstrate behavior.

## Security & Configuration Tips
- XML safety: DTDs are disabled; keep secure reader settings.
- Encrypted ODS: rejected by default; only allow via `new OdsReaderOptions { ThrowOnEncrypted = false }` with justification.
- Resource caps: respect `MaxSheets/MaxRowsPerSheet/MaxColumnsPerRow/MaxRepeated*` to prevent memory pressure; do not raise defaults casually.

## Quick Usage Example
```csharp
var doc = Ods.Load("sample.ods");
var cell = doc.Sheets[0].GetCell(0, 0);
```

## Git Commit 工作流程

**觸發指令：** `git commit`

當使用者輸入 `git commit` 時，自動執行以下步驟：

1.  執行 `git diff HEAD` 分析變更。
2.  根據變更內容，產生一個結構良好、使用正體中文的 Conventional Commits 訊息。
3.  將變更的檔案加入暫存區。
4.  將 commit 訊息寫入暫存檔。
5.  使用 `git commit -F <temp_file>` 執行 commit。
6.  刪除暫存檔。