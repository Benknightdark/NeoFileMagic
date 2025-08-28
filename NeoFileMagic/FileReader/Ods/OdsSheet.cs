namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Collections.Generic;

public sealed class OdsSheet
{
    public string Name { get; }
    public IReadOnlyList<OdsRow> Rows { get; }
    public int RowCount => Rows.Count;
    public int MaxColumnCount { get; }

    internal OdsSheet(string name, List<OdsRow> rows, int maxColumnCount)
    {
        Name = name;
        Rows = rows;
        MaxColumnCount = maxColumnCount;
    }

    public OdsCell GetCell(int rowIndex, int columnIndex)
    {
        if (rowIndex < 0 || rowIndex >= Rows.Count) throw new ArgumentOutOfRangeException(nameof(rowIndex));
        var row = Rows[rowIndex];
        if (columnIndex < 0 || columnIndex >= row.Cells.Count) return OdsCell.Empty; // 超界回傳空白
        return row.Cells[columnIndex];
    }
}
