namespace NeoFileMagic.FileReader.Ods;

/// <summary>
/// 達上限時的處理策略。
/// </summary>
public enum OdsLimitMode
{
    /// <summary>達上限即拋例外（原行為）。</summary>
    Throw,
    /// <summary>達上限即截斷，不再加入更多資料。</summary>
    Truncate
}
