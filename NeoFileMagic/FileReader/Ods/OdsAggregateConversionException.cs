using System;

namespace NeoFileMagic.FileReader.Ods;

public sealed class OdsAggregateConversionException : Exception
{
public IReadOnlyList<OdsRowConversionException> Errors { get; }


public OdsAggregateConversionException(IEnumerable<OdsRowConversionException> errors)
: base(BuildMessage(errors))
{
Errors = errors.ToArray();
}


private static string BuildMessage(IEnumerable<OdsRowConversionException> errs)
{
var list = errs.ToList();
var head = $"資料轉換發生 {list.Count} 筆錯誤（僅列出前 5 筆）：";
var lines = list.Take(5).Select(e => " - " + e.Message);
return string.Join(Environment.NewLine, new[] { head }.Concat(lines));
}
}
