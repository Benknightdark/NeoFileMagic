namespace NeoFileMagic.FileReader.Ods;

/// <summary>
/// 讀取設定與安全限制。
/// </summary>
public sealed class OdsReaderOptions
{
    /// <summary>若偵測為加密檔，是否丟出 NotSupportedException（預設 true）。</summary>
    public bool ThrowOnEncrypted { get; init; } = true;

    /// <summary>最大允許工作表數，避免惡意檔案（預設 256）。</summary>
    public int MaxSheets { get; init; } = 256;

    /// <summary>每個工作表最大列數（預設 1,000,000）。</summary>
    public int MaxRowsPerSheet { get; init; } = 1_000_000;

    /// <summary>每列最大欄數（預設 16,384）。</summary>
    public int MaxColumnsPerRow { get; init; } = 16_384;

    /// <summary>展開重複列時的上限（預設 1,000,000）。</summary>
    public int MaxRepeatedRows { get; init; } = 1_000_000;

    /// <summary>展開重複欄時的上限（預設 16,384）。</summary>
    public int MaxRepeatedColumns { get; init; } = 16_384;

    /// <summary>text:s 連續空白字元展開上限（預設 100）。</summary>
    public int MaxTextSpaceRun { get; init; } = 100;

    /// <summary>達上限時策略（預設 Truncate：截斷）。</summary>
    public OdsLimitMode LimitMode { get; init; } = OdsLimitMode.Truncate;

    /// <summary>對完全空白且被重複的列採折疊（預設 true，不展開空白重複列）。</summary>
    public bool CollapseEmptyRepeatedRows { get; init; } = true;

    /// <summary>修剪列尾部連續空欄（Empty 或空字串），可降低欄數與輸出量（預設 true）。</summary>
    public bool TrimTrailingEmptyCells { get; init; } = true;
}
