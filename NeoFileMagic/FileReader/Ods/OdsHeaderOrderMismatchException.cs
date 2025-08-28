using System;

namespace NeoFileMagic.FileReader.Ods;

public sealed class OdsHeaderOrderMismatchException : Exception
{
    public IReadOnlyList<string> ExpectedOrder { get; }
    public IReadOnlyList<string> ActualSubsequence { get; }
    public IReadOnlyList<string> ActualAllHeaders { get; }


    public OdsHeaderOrderMismatchException(IEnumerable<string> expectedOrder,
    IEnumerable<string> actualSubsequence,
    IEnumerable<string> actualAllHeaders)
    : base(BuildMessage(expectedOrder, actualSubsequence, actualAllHeaders))
    {
        ExpectedOrder = expectedOrder.ToArray();
        ActualSubsequence = actualSubsequence.ToArray();
        ActualAllHeaders = actualAllHeaders.ToArray();
    }


    private static string BuildMessage(IEnumerable<string> expected, IEnumerable<string> actualSub, IEnumerable<string> actualAll)
    => $"欄位順序不一致。期望順序：[{string.Join(" → ", expected)}]；實際對應子序列：[{string.Join(" → ", actualSub)}]；完整表頭：[{string.Join(", ", actualAll)}]";
}
