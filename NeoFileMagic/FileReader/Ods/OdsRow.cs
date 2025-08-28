namespace NeoFileMagic.FileReader.Ods;

using System.Collections.Generic;

public readonly struct OdsRow
{
    public IReadOnlyList<OdsCell> Cells { get; }
    public int ColumnCount => Cells.Count;

    internal OdsRow(IReadOnlyList<OdsCell> cells) => Cells = cells;
}
