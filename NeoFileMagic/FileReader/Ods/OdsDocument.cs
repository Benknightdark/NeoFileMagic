namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Collections.Generic;
using System.Linq;

public sealed class OdsDocument
{
    public IReadOnlyList<OdsSheet> Sheets { get; }

    internal OdsDocument(List<OdsSheet> sheets) => Sheets = sheets;

    public OdsSheet this[string name] => Sheets.First(s => string.Equals(s.Name, name, StringComparison.Ordinal));
}
