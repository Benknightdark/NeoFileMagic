namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;

internal static class OdsXml
{
    // 命名空間（ODF 1.2 常見）
    private const string NsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private const string NsTable = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private const string NsText = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private const string NsManifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

    internal static bool HasEncryption(Stream manifestXml)
    {
        var settings = SecureXmlReaderSettings();
        using var xr = XmlReader.Create(manifestXml, settings);
        while (xr.Read())
        {
            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "encryption-data" && xr.NamespaceURI == NsManifest)
                return true;
        }
        return false;
    }

    internal static List<OdsSheet> ParseContent(Stream contentXml, OdsReaderOptions options)
    {
        var settings = SecureXmlReaderSettings();
        using var xr = XmlReader.Create(contentXml, settings);

        var sheets = new List<OdsSheet>(capacity: 8);

        // 導航至 office:spreadsheet
        if (!MoveToElement(xr, "spreadsheet", NsOffice))
            throw new InvalidDataException("office:spreadsheet not found in content.xml");

        int sheetCount = 0;
        while (xr.Read())
        {
            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "table" && xr.NamespaceURI == NsTable)
            {
                sheetCount++;
                if (sheetCount > options.MaxSheets)
                    throw new InvalidDataException($"Sheet count exceeds limit {options.MaxSheets}.");

                var sheet = ReadTable(xr, options);
                sheets.Add(sheet);
            }
            else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "spreadsheet" && xr.NamespaceURI == NsOffice)
            {
                break;
            }
        }

        return sheets;
    }

    private static OdsSheet ReadTable(XmlReader xr, OdsReaderOptions options)
    {
        // 目前位於 <table:table>
        var name = xr.GetAttribute("name", NsTable) ?? "Sheet";
        var rows = new List<OdsRow>(capacity: 128);
        var tableDepth = xr.Depth;

        if (xr.IsEmptyElement)
            return new OdsSheet(name, rows, options.MaxColumnsPerRow);

        while (xr.Read())
        {
            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "table-row" && xr.NamespaceURI == NsTable)
            {
                var repeat = GetIntAttr(xr, "number-rows-repeated", NsTable) ?? 1;
                var row = ReadRow(xr, options);

                // 安全上限
                repeat = Math.Min(repeat, options.MaxRepeatedRows);

                // 若此列有效內容為空且允許折疊，則不展開此重複列（避免空白列爆量）
                if (options.CollapseEmptyRepeatedRows && IsRowEffectivelyEmpty(in row))
                {
                    // 若希望至少保留 1 列，可改為：
                    // if (rows.Count < options.MaxRowsPerSheet) rows.Add(row);
                    continue;
                }

                int remaining = options.MaxRowsPerSheet - rows.Count;
                if (remaining <= 0)
                {
                    if (options.LimitMode == OdsLimitMode.Throw)
                        throw new InvalidDataException($"Row count exceeds limit {options.MaxRowsPerSheet} in sheet '{name}'.");
                    // Truncate: 已達上限，忽略後續列
                    continue;
                }

                int toAdd = Math.Min(repeat, remaining);
                for (int i = 0; i < toAdd; i++) rows.Add(row);

                if (toAdd < repeat && options.LimitMode == OdsLimitMode.Throw)
                    throw new InvalidDataException($"Row count exceeds limit {options.MaxRowsPerSheet} in sheet '{name}'.");
            }
            else if (xr.NodeType == XmlNodeType.EndElement && xr.Depth == tableDepth && xr.LocalName == "table" && xr.NamespaceURI == NsTable)
            {
                break;
            }
        }

        return new OdsSheet(name, rows, options.MaxColumnsPerRow);
    }

    private static OdsRow ReadRow(XmlReader xr, OdsReaderOptions options)
    {
        // 目前位於 <table:table-row>
        var rowDepth = xr.Depth;
        var segments = new List<OdsCellSegment>(capacity: 16);

        if (xr.IsEmptyElement)
            return new OdsRow(new OdsCellList(segments, options.MaxColumnsPerRow, options.TrimTrailingEmptyCells));

        while (xr.Read())
        {
            if (xr.NodeType == XmlNodeType.Element)
            {
                if (xr.LocalName == "table-cell" && xr.NamespaceURI == NsTable)
                {
                    var cell = ReadCell(xr, options);
                    var repeat = GetIntAttr(xr, "number-columns-repeated", NsTable) ?? 1;
                    repeat = Math.Min(repeat, options.MaxRepeatedColumns);
                    if (repeat > 0)
                        segments.Add(new OdsCellSegment(cell, repeat));
                }
                else if (xr.LocalName == "covered-table-cell" && xr.NamespaceURI == NsTable)
                {
                    var repeat = GetIntAttr(xr, "number-columns-repeated", NsTable) ?? 1;
                    repeat = Math.Min(repeat, options.MaxRepeatedColumns);
                    if (repeat > 0)
                        segments.Add(new OdsCellSegment(OdsCell.Empty, repeat));

                    if (!xr.IsEmptyElement) xr.Skip();
                }
                else
                {
                    // 其他節點直接略過其子樹
                    if (!xr.IsEmptyElement) xr.Skip();
                }
            }
            else if (xr.NodeType == XmlNodeType.EndElement && xr.Depth == rowDepth && xr.LocalName == "table-row" && xr.NamespaceURI == NsTable)
            {
                break;
            }
        }

        return new OdsRow(new OdsCellList(segments, options.MaxColumnsPerRow, options.TrimTrailingEmptyCells));
    }

    private static OdsCell ReadCell(XmlReader xr, OdsReaderOptions options)
    {
        // 目前位於 <table:table-cell>
        var vt = xr.GetAttribute("value-type", NsOffice);
        var formula = xr.GetAttribute("formula", NsTable); // 格式通常像 of:=SUM([.A1:.A3])
        var cell = new OdsCell();

        switch (vt)
        {
            case "string":
                {
                    string text = ReadCellText(xr, options);
                    cell = new OdsCell { Type = OdsValueType.String, Text = text, Formula = formula };
                    break;
                }
            case "float":
                {
                    var d = GetDoubleAttr(xr, "value", NsOffice);
                    cell = new OdsCell { Type = OdsValueType.Float, Number = d, Formula = formula };
                    if (!xr.IsEmptyElement) xr.Skip();
                    break;
                }
            case "currency":
                {
                    var d = GetDoubleAttr(xr, "value", NsOffice);
                    var cur = xr.GetAttribute("currency", NsOffice);
                    cell = new OdsCell { Type = OdsValueType.Currency, Number = d, Currency = cur, Formula = formula };
                    if (!xr.IsEmptyElement) xr.Skip();
                    break;
                }
            case "boolean":
                {
                    var bStr = xr.GetAttribute("boolean-value", NsOffice);
                    bool? b = bStr is null ? null : string.Equals(bStr, "true", StringComparison.OrdinalIgnoreCase) ? true : string.Equals(bStr, "false", StringComparison.OrdinalIgnoreCase) ? false : null;
                    cell = new OdsCell { Type = OdsValueType.Boolean, Boolean = b, Formula = formula };
                    if (!xr.IsEmptyElement) xr.Skip();
                    break;
                }
            case "date":
                {
                    var dv = xr.GetAttribute("date-value", NsOffice);
                    DateTimeOffset? dto = null;
                    if (!string.IsNullOrEmpty(dv))
                    {
                        try { dto = System.Xml.XmlConvert.ToDateTimeOffset(dv); }
                        catch { /* ignore */ }
                    }
                    cell = new OdsCell { Type = OdsValueType.Date, DateTime = dto, Formula = formula };
                    if (!xr.IsEmptyElement) xr.Skip();
                    break;
                }
            case "time":
                {
                    var tv = xr.GetAttribute("time-value", NsOffice);
                    TimeSpan? ts = null;
                    if (!string.IsNullOrEmpty(tv))
                    {
                        try { ts = System.Xml.XmlConvert.ToTimeSpan(tv); } catch { }
                    }
                    cell = new OdsCell { Type = OdsValueType.Time, Time = ts, Formula = formula };
                    if (!xr.IsEmptyElement) xr.Skip();
                    break;
                }
            default:
                {
                    // 空白或未知型別
                    if (!xr.IsEmptyElement) xr.Skip();
                    cell = OdsCell.Empty;
                    break;
                }
        }

        return cell;
    }

    private static string ReadCellText(XmlReader xr, OdsReaderOptions options)
    {
        if (xr.IsEmptyElement) return string.Empty;

        var sb = new StringBuilder(32);
        var depth = xr.Depth;
        bool firstParagraph = true;

        while (xr.Read())
        {
            if (xr.NodeType == XmlNodeType.Element)
            {
                if (xr.LocalName == "p" && xr.NamespaceURI == NsText)
                {
                    if (!firstParagraph) sb.Append('\n');
                    firstParagraph = false;
                    if (xr.IsEmptyElement) continue;

                    var pDepth = xr.Depth;
                    while (xr.Read())
                    {
                        if (xr.NodeType == XmlNodeType.Text)
                        {
                            sb.Append(xr.Value);
                        }
                        else if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "s" && xr.NamespaceURI == NsText)
                        {
                            var c = GetIntAttr(xr, "c", NsText) ?? 1; // 連續空白
                            c = Math.Clamp(c, 0, options.MaxTextSpaceRun);
                            if (c > 0) sb.Append(' ', c);
                            if (!xr.IsEmptyElement) xr.Skip();
                        }
                        else if (xr.NodeType == XmlNodeType.EndElement && xr.Depth == pDepth && xr.LocalName == "p" && xr.NamespaceURI == NsText)
                        {
                            break;
                        }
                        else if (xr.NodeType == XmlNodeType.Element && !xr.IsEmptyElement)
                        {
                            // 其他內嵌元素（粗體/連結等）—當作純文字讀取
                            using var sub = xr.ReadSubtree();
                            while (sub.Read())
                            {
                                if (sub.NodeType == XmlNodeType.Text) sb.Append(sub.Value);
                            }
                            xr.Skip();
                        }
                    }
                }
                else
                {
                    // 其他子節點直接略過
                    if (!xr.IsEmptyElement) xr.Skip();
                }
            }
            else if (xr.NodeType == XmlNodeType.Text)
            {
                sb.Append(xr.Value);
            }
            else if (xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth && xr.LocalName == "table-cell" && xr.NamespaceURI == NsTable)
            {
                break;
            }
        }

        return sb.ToString();
    }

    private static XmlReaderSettings SecureXmlReaderSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        IgnoreWhitespace = true,
        XmlResolver = null
    };

    private static bool MoveToElement(XmlReader xr, string localName, string ns)
    {
        while (xr.Read())
        {
            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == localName && xr.NamespaceURI == ns)
                return true;
        }
        return false;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int? GetIntAttr(XmlReader xr, string local, string ns)
    {
        var s = xr.GetAttribute(local, ns);
        if (string.IsNullOrEmpty(s)) return null;
        if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)) return v;
        return null;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static double? GetDoubleAttr(XmlReader xr, string local, string ns)
    {
        var s = xr.GetAttribute(local, ns);
        if (string.IsNullOrEmpty(s)) return null;
        if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var v)) return v;
        return null;
    }

    /// <summary>
    /// 判斷一列是否「有效內容為空」。空白字串與 Empty 型別視為空；其他型別只要有值就視為非空。
    /// </summary>
    private static bool IsRowEffectivelyEmpty(in OdsRow row)
    {
        var cells = row.Cells;
        if (cells == null || cells.Count == 0) return true;
        foreach (var c in cells)
        {
            if (IsCellEffectivelyEmpty(c)) continue;
            return false;
        }
        return true;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static bool IsCellEffectivelyEmpty(in OdsCell c)
        => c.Type switch
        {
            OdsValueType.Empty => true,
            OdsValueType.String => string.IsNullOrEmpty(c.Text),
            _ => false
        };

    // ===== RLE（Run-Length Encoding）欄位容器 =====

    private readonly struct OdsCellSegment
    {
        public readonly OdsCell Cell;
        public readonly int Count;
        public OdsCellSegment(OdsCell cell, int count)
        {
            Cell = cell;
            Count = count;
        }
    }

    private sealed class OdsCellList : IReadOnlyList<OdsCell>
    {
        private readonly OdsCellSegment[] _segments;
        private readonly int _count; // 展開後的欄位數（已應用上限與尾端修剪）

        public OdsCellList(List<OdsCellSegment> segments, int maxColumns, bool trimTrailingEmpty)
        {
            if (segments.Count == 0)
            {
                _segments = Array.Empty<OdsCellSegment>();
                _count = 0;
                return;
            }

            // 先裁切到 maxColumns，同時複製到陣列
            var tmp = new List<OdsCellSegment>(segments.Count);
            int total = 0;
            foreach (var seg in segments)
            {
                if (seg.Count <= 0) continue;
                if (total >= maxColumns) break;
                int take = Math.Min(seg.Count, maxColumns - total);
                tmp.Add(new OdsCellSegment(seg.Cell, take));
                total += take;
            }

            if (trimTrailingEmpty && total > 0)
            {
                // 從尾端修剪連續的空欄（Empty 或空字串）
                for (int i = tmp.Count - 1; i >= 0; i--)
                {
                    if (!IsCellEffectivelyEmpty(tmp[i].Cell)) break;
                    total -= tmp[i].Count;
                    tmp.RemoveAt(i);
                }
            }

            _segments = tmp.Count == 0 ? Array.Empty<OdsCellSegment>() : tmp.ToArray();
            _count = total;
        }

        public int Count => _count;

        public OdsCell this[int index]
        {
            get
            {
                if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index));
                int i = index;
                for (int s = 0; s < _segments.Length; s++)
                {
                    var seg = _segments[s];
                    if (i < seg.Count) return seg.Cell;
                    i -= seg.Count;
                }
                // 理論上不會到這裡
                return OdsCell.Empty;
            }
        }

        public IEnumerator<OdsCell> GetEnumerator() => new Enumerator(_segments, _count);
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        private sealed class Enumerator : IEnumerator<OdsCell>
        {
            private readonly OdsCellSegment[] _segs;
            private readonly int _targetCount;
            private int _emitted;
            private int _segIndex;
            private int _segOffset;
            private OdsCell _current;

            public Enumerator(OdsCellSegment[] segs, int count)
            {
                _segs = segs;
                _targetCount = count;
                _emitted = 0;
                _segIndex = 0;
                _segOffset = 0;
                _current = default;
            }

            public OdsCell Current => _current;
            object IEnumerator.Current => _current;

            public bool MoveNext()
            {
                if (_emitted >= _targetCount) return false;

                while (_segIndex < _segs.Length)
                {
                    var seg = _segs[_segIndex];
                    if (_segOffset < seg.Count)
                    {
                        _current = seg.Cell;
                        _segOffset++;
                        _emitted++;
                        return true;
                    }
                    _segIndex++;
                    _segOffset = 0;
                }

                return false;
            }

            public void Reset()
            {
                _emitted = 0;
                _segIndex = 0;
                _segOffset = 0;
                _current = default;
            }

            public void Dispose() { }
        }
    }
}
