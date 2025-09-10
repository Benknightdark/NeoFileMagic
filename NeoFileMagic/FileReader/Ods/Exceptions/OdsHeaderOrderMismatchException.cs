using System;

namespace NeoFileMagic.FileReader.Ods.Exception;

/// <summary>
/// 欄位順序不一致時拋出的例外：實際表頭子序列與期望順序不同。
/// </summary>
public sealed class OdsHeaderOrderMismatchException : System.Exception
{
    /// <summary>期望的欄位順序（正規化後）。</summary>
    public IReadOnlyList<string> ExpectedOrder { get; }
    /// <summary>實際表頭中與期望對應的子序列（正規化後）。</summary>
    public IReadOnlyList<string> ActualSubsequence { get; }
    /// <summary>實際所有表頭（正規化後）。</summary>
    public IReadOnlyList<string> ActualAllHeaders { get; }


    /// <summary>
    /// 建立例外並帶入期望順序、實際子序列與完整表頭。
    /// </summary>
    public OdsHeaderOrderMismatchException(IEnumerable<string> expectedOrder,
    IEnumerable<string> actualSubsequence,
    IEnumerable<string> actualAllHeaders)
    : base(BuildMessage(expectedOrder, actualSubsequence, actualAllHeaders))
    {
        ExpectedOrder = expectedOrder.ToArray();
        ActualSubsequence = actualSubsequence.ToArray();
        ActualAllHeaders = actualAllHeaders.ToArray();
    }


    /// <summary>
    /// 建立可讀的錯誤訊息。
    /// </summary>
    private static string BuildMessage(IEnumerable<string> expected, IEnumerable<string> actualSub, IEnumerable<string> actualAll)
    => $"欄位順序不一致。期望順序：[{string.Join(" → ", expected)}]；實際對應子序列：[{string.Join(" → ", actualSub)}]；完整表頭：[{string.Join(", ", actualAll)}]";
}
