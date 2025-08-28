# Repository Guidelines

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

