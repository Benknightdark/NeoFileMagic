using System;

namespace NeoFileMagic.FileReader.Ods;

public sealed class OdsRowConversionException : Exception
{
public int RowIndex { get; } // 0-based
public int ColumnIndex { get; } // -1 表示未對應
public string JsonPropertyName { get; }
public string? HeaderName { get; }
public Type TargetType { get; }
public OdsValueType CellType { get; }
public string? CellPreview { get; }


public OdsRowConversionException(
int rowIndex, int columnIndex, string jsonPropertyName, string? headerName,
Type targetType, OdsValueType cellType, string? cellPreview, Exception? inner = null)
: base($"第 {rowIndex + 1} 列 [欄:{headerName ?? "(未對應)"}, 屬性:{jsonPropertyName}] 期望型別 {targetType.Name}，實得 {cellType} 值「{cellPreview}」無法轉換。", inner)
{
RowIndex = rowIndex;
ColumnIndex = columnIndex;
JsonPropertyName = jsonPropertyName;
HeaderName = headerName;
TargetType = targetType;
CellType = cellType;
CellPreview = cellPreview;
}
}