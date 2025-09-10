namespace NeoFileMagic.FileReader.Ods;

using System.Collections.Generic;

/// <summary>
/// 代表單一資料列。欄資料可能經過 RLE（連續區段）壓縮儲存。
/// </summary>
public readonly struct OdsRow
{
    /// <summary>
    /// 此列的欄位集合。
    /// </summary>
    public IReadOnlyList<OdsCell> Cells { get; }

    /// <summary>
    /// 欄數。
    /// </summary>
    public int ColumnCount => Cells.Count;

    /// <summary>
    /// 以指定欄位集合建立資料列。
    /// </summary>
    /// <param name="cells">欄位集合。</param>
    internal OdsRow(IReadOnlyList<OdsCell> cells) => Cells = cells;
}
