using System;

namespace NeoFileMagic.FileReader.Ods.Exception;

/// <summary>
/// 表頭比對不一致時拋出的例外：缺少必要欄位或名稱對不上。
/// </summary>
public sealed class OdsHeaderMismatchException : System.Exception
{
    /// <summary>期望的表頭清單（正規化後）。</summary>
    public IReadOnlyList<string> ExpectedHeaders { get; }
    /// <summary>缺少的表頭（正規化後）。</summary>
    public IReadOnlyList<string> MissingHeaders { get; }
    /// <summary>實際可用表頭（正規化後）。</summary>
    public IReadOnlyList<string> AvailableHeaders { get; }


    /// <summary>
    /// 建立例外並帶入期望、缺少與實際表頭。
    /// </summary>
    public OdsHeaderMismatchException(
    IEnumerable<string> expected,
    IEnumerable<string> missing,
    IEnumerable<string> available)
    : base(BuildMessage(expected, missing, available))
    {
        ExpectedHeaders = expected.ToArray();
        MissingHeaders = missing.ToArray();
        AvailableHeaders = available.ToArray();
    }


    /// <summary>
    /// 建立可讀的錯誤訊息。
    /// </summary>
    private static string BuildMessage(IEnumerable<string> expected, IEnumerable<string> missing, IEnumerable<string> available)
    => $"表頭不一致：缺少 [{string.Join(", ", missing)}]。期望表頭：[{string.Join(", ", expected)}]；實際表頭：[{string.Join(", ", available)}]";
}
