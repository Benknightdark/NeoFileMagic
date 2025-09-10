namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Collections.Generic;

/// <summary>
/// 代表單一工作表，提供列集合與索引存取能力。
/// </summary>
public sealed class OdsSheet
{
    /// <summary>
    /// 工作表名稱。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 此工作表的列集合。
    /// </summary>
    public IReadOnlyList<OdsRow> Rows { get; }

    /// <summary>
    /// 列數。
    /// </summary>
    public int RowCount => Rows.Count;

    /// <summary>
    /// 單列最大欄數上限（由讀取選項決定）。
    /// </summary>
    public int MaxColumnCount { get; }

    /// <summary>
    /// 建立 <see cref="OdsSheet"/> 實例。
    /// </summary>
    /// <param name="name">工作表名稱。</param>
    /// <param name="rows">列集合。</param>
    /// <param name="maxColumnCount">每列欄位數上限。</param>
    internal OdsSheet(string name, List<OdsRow> rows, int maxColumnCount)
    {
        Name = name;
        Rows = rows;
        MaxColumnCount = maxColumnCount;
    }

    /// <summary>
    /// 以（列、欄）座標取得儲存格。
    /// </summary>
    /// <param name="rowIndex">列索引（0 起算）。</param>
    /// <param name="columnIndex">欄索引（0 起算）。</param>
    /// <returns>對應的 <see cref="OdsCell"/>；若超出欄界則回傳 <see cref="OdsCell.Empty"/>。</returns>
    /// <exception cref="ArgumentOutOfRangeException">當 <paramref name="rowIndex"/> 不在有效範圍內時。</exception>
    public OdsCell GetCell(int rowIndex, int columnIndex)
    {
        if (rowIndex < 0 || rowIndex >= Rows.Count) throw new ArgumentOutOfRangeException(nameof(rowIndex));
        var row = Rows[rowIndex];
        if (columnIndex < 0 || columnIndex >= row.Cells.Count) return OdsCell.Empty; // 超界回傳空白
        return row.Cells[columnIndex];
    }
}
