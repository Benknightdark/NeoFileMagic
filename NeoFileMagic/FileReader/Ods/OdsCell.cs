namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Globalization;

/// <summary>
/// ODS 儲存格資料模型，描述儲存格的型別與對應值。
/// </summary>
public readonly struct OdsCell
{
    /// <summary>
    /// 代表空白儲存格的預設實例。
    /// </summary>
    public static readonly OdsCell Empty = new();

    /// <summary>
    /// 儲存格值的型別。
    /// </summary>
    public OdsValueType Type { get; init; }

    /// <summary>
    /// 文字內容（僅在 <see cref="Type"/> 為 <see cref="OdsValueType.String"/> 時使用）。
    /// </summary>
    public string? Text { get; init; }

    /// <summary>
    /// 數值（含貨幣型別共用此欄）。
    /// </summary>
    public double? Number { get; init; }

    /// <summary>
    /// 貨幣代碼（僅在 <see cref="Type"/> 為 <see cref="OdsValueType.Currency"/> 時使用）。
    /// </summary>
    /// <remarks>依 ODF 內容通常為 ISO 4217 代碼，例如「USD」。</remarks>
    public string? Currency { get; init; }

    /// <summary>
    /// 日期或日期時間值。
    /// </summary>
    public DateTimeOffset? DateTime { get; init; }

    /// <summary>
    /// 時間段值。
    /// </summary>
    public TimeSpan? Time { get; init; }

    /// <summary>
    /// 布林值。
    /// </summary>
    public bool? Boolean { get; init; }

    /// <summary>
    /// 儲存格公式（若存在）。格式通常為 <c>of:=...</c>。
    /// </summary>
    public string? Formula { get; init; }

    /// <summary>
    /// 以可讀字串輸出儲存格內容。
    /// </summary>
    /// <returns>
    /// 文字型別輸出 <see cref="Text"/>；數值/貨幣使用 InvariantCulture；
    /// 日期使用 ISO-8601；時間使用 <see cref="TimeSpan.ToString()"/>；
    /// 其他或空值回傳空字串。
    /// </returns>
    public override string ToString() => Type switch
    {
        OdsValueType.String => Text ?? string.Empty,
        OdsValueType.Float or OdsValueType.Currency => Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
        OdsValueType.Boolean => Boolean?.ToString() ?? string.Empty,
        OdsValueType.Date => DateTime?.ToString("o") ?? string.Empty,
        OdsValueType.Time => Time?.ToString() ?? string.Empty,
        _ => string.Empty,
    };
}
