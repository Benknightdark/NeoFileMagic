namespace NeoFileMagic.FileReader.Ods;

/// <summary>
/// 文字輸出處理策略，用於將多行或含控制字元的文字轉為單行表示。
/// </summary>
public enum TextHandling
{
    /// <summary>
    /// 保持原樣，不處理換行或跳脫字元。
    /// </summary>
    Keep,
    /// <summary>
    /// 以跳脫序列表示控制字元（如 <c>\n</c>、<c>\t</c>、反斜線）。
    /// </summary>
    Escape,
    /// <summary>
    /// 將換行/定位等空白折疊為單一空格，並去除多餘空白。
    /// </summary>
    CollapseToSpace,
    /// <summary>
    /// 只取第一段（遇到第一個換行即截斷並修剪）。
    /// </summary>
    FirstParagraph
}
