using System;

namespace NeoFileMagic.FileReader.Ods.Exception;

/// <summary>
/// 將多筆列轉換錯誤彙總後拋出的例外，
/// 內含各筆列錯誤的詳細資訊。
/// </summary>
public sealed class OdsAggregateConversionException : System.Exception
{
    /// <summary>彙總的列轉換錯誤清單。</summary>
    public IReadOnlyList<OdsRowConversionException> Errors { get; }


    /// <summary>
    /// 以多筆列轉換錯誤建立例外。
    /// </summary>
    public OdsAggregateConversionException(IEnumerable<OdsRowConversionException> errors)
    : base(BuildMessage(errors))
    {
        Errors = errors.ToArray();
    }


    /// <summary>
    /// 建立可讀的彙總錯誤訊息（僅取前數筆）。
    /// </summary>
    private static string BuildMessage(IEnumerable<OdsRowConversionException> errs)
    {
        var list = errs.ToList();
        var head = $"資料轉換發生 {list.Count} 筆錯誤（僅列出前 5 筆）：";
        var lines = list.Take(5).Select(e => " - " + e.Message);
        return string.Join(Environment.NewLine, new[] { head }.Concat(lines));
    }
}
