namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// 代表整份 ODS 活頁簿，包含多個工作表。
/// </summary>
public sealed class OdsDocument
{
    /// <summary>
    /// 工作表集合。
    /// </summary>
    public IReadOnlyList<OdsSheet> Sheets { get; }

    /// <summary>
    /// 以已解析的工作表集合建立文件。
    /// </summary>
    /// <param name="sheets">工作表清單。</param>
    internal OdsDocument(List<OdsSheet> sheets) => Sheets = sheets;

    /// <summary>
    /// 以名稱存取指定工作表（以序名完全相等比對）。
    /// </summary>
    /// <param name="name">工作表名稱。</param>
    /// <returns>符合名稱的工作表。</returns>
    /// <exception cref="InvalidOperationException">找不到同名工作表。</exception>
    public OdsSheet this[string name] => Sheets.First(s => string.Equals(s.Name, name, StringComparison.Ordinal));
}
