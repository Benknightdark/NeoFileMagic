namespace NeoFileMagic.FileReader.Ods;

/// <summary>
/// ODS 儲存格值的型別。
/// </summary>
public enum OdsValueType
{
    /// <summary>空白或未知。</summary>
    Empty,
    /// <summary>文字。</summary>
    String,
    /// <summary>數值（浮點）。</summary>
    Float,
    /// <summary>貨幣數值。</summary>
    Currency,
    /// <summary>日期或日期時間。</summary>
    Date,
    /// <summary>時間段。</summary>
    Time,
    /// <summary>布林值。</summary>
    Boolean
}
