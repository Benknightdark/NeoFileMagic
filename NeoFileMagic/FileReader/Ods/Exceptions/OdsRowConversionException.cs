using System;

namespace NeoFileMagic.FileReader.Ods.Exception;

/// <summary>
/// 單列資料轉換為目標模型失敗時拋出的例外，
/// 含列/欄索引、表頭與目標型別等診斷資訊。
/// </summary>
public sealed class OdsRowConversionException : System.Exception
{
    /// <summary>列索引（0 起算）。</summary>
    public int RowIndex { get; }
    /// <summary>欄索引（-1 表示未對應到欄位）。</summary>
    public int ColumnIndex { get; }
    /// <summary>目標模型屬性的 JSON 名稱。</summary>
    public string JsonPropertyName { get; }
    /// <summary>來源表頭名稱（若有）。</summary>
    public string? HeaderName { get; }
    /// <summary>目標屬性型別。</summary>
    public Type TargetType { get; }
    /// <summary>原始儲存格的型別。</summary>
    public OdsValueType CellType { get; }
    /// <summary>原始儲存格的文字預覽。</summary>
    public string? CellPreview { get; }


    /// <summary>
    /// 建立例外並帶入轉換失敗之詳細資訊。
    /// </summary>
    public OdsRowConversionException(
    int rowIndex, int columnIndex, string jsonPropertyName, string? headerName,
    Type targetType, OdsValueType cellType, string? cellPreview, System.Exception? inner = null)
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
