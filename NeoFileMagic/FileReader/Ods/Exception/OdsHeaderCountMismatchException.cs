using System;

namespace NeoFileMagic.FileReader.Ods.Exception;

/// <summary>
/// 表頭數量與模型屬性數不一致時拋出的例外。
/// </summary>
public sealed class OdsHeaderCountMismatchException : System.Exception
{
    /// <summary>期望的欄位數量（正規化後）。</summary>
    public int ExpectedCount { get; }
    /// <summary>實際的欄位數量（非空表頭）。</summary>
    public int ActualCount { get; }
    /// <summary>期望的表頭清單。</summary>
    public IReadOnlyList<string> ExpectedHeaders { get; }
    /// <summary>實際的表頭清單（非空）。</summary>
    public IReadOnlyList<string> ActualHeaders { get; }


    /// <summary>
    /// 建立例外並帶入期望與實際的數量與表頭。
    /// </summary>
    public OdsHeaderCountMismatchException(int expectedCount, int actualCount,
    IEnumerable<string> expectedHeaders,
    IEnumerable<string> actualHeaders)
    : base($"欄位數量不一致：期望 {expectedCount} 欄，實際 {actualCount} 欄。期望表頭：[{string.Join(", ", expectedHeaders)}]；實際表頭：[{string.Join(", ", actualHeaders)}]")
    {
        ExpectedCount = expectedCount;
        ActualCount = actualCount;
        ExpectedHeaders = expectedHeaders.ToArray();
        ActualHeaders = actualHeaders.ToArray();
    }
}
