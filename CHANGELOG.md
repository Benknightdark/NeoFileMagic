# 變更紀錄（CHANGELOG）

本專案遵循 Keep a Changelog 格式，並參考 Conventional Commits 以撰寫提交訊息。

## [Unreleased]

## [2.0.0] - 2025-08-29

### 重大變更
- 反序列化改為僅支援 Newtonsoft.Json 屬性：不再支援 `System.Text.Json` 的 `[JsonPropertyName]`/`[JsonPropertyOrder]`，請改用 `Newtonsoft.Json` 的 `[JsonProperty(PropertyName=..., Order=...)]`；或以表頭文字直接對應。（自 v1.0.3 起變更，含明確 BREAKING CHANGE 註記）
- 調整日期/時間語意：保留原始牆鐘時間與原始值，不再自動轉換為本地或 UTC。
- 統一 ODS 入口類別：移除舊 Ods 入口，改以 `NeoOds` 作為統一進入點。

### 新增
- ODS：依 `office:value-type` 解析 `float/currency/boolean/date/time/string`，資訊不足時回退為字串。
- ODS：保留 `table:formula` 至 `OdsCell.Formula`。

### 變更
- ODS：`IsCellEffectivelyEmpty` 判定邏輯調整，依型別語意判定是否為「有效空白」。
- NeoOds：輸出物件鍵名改採 `[JsonProperty]` 指定之 `PropertyName`（若有），避免 Newtonsoft.Json 在名稱覆寫時的繫結失敗。
- README 與 Sample：補充 ODS 型別解析與公式說明，並更新示例程式與輸出 JSON。

### 其他
- chore：更新範例輸出檔以反映最新行為。

## [0.1.0] - 2025-08-29

### 新增
- ODS：依 `office:value-type` 解析 `float/currency/boolean/date/time/string`，資訊不足時回退為字串。
- ODS：保留 `table:formula` 至 `OdsCell.Formula`。

### 變更
- ODS：調整 `IsCellEffectivelyEmpty` 判定邏輯以符合各型別語意。
- 將輸出物件的鍵名改採 `JsonProperty` 指定名稱（若有），避免 Newtonsoft.Json 在名稱覆寫時無法繫結。
- README：補充 ODS 型別解析與公式相關說明。
- Sample：更新示例程式與輸出 JSON。

[Unreleased]: https://github.com/your-org/NeoFileMagic/compare/v2.0.0...HEAD
[2.0.0]: https://github.com/your-org/NeoFileMagic/compare/v1.0.3...v2.0.0
[0.1.0]: https://github.com/your-org/NeoFileMagic/releases/tag/v0.1.0
