using System;

namespace NeoFileMagic.FileReader.Ods;

public sealed class OdsHeaderCountMismatchException : Exception
{
public int ExpectedCount { get; }
public int ActualCount { get; }
public IReadOnlyList<string> ExpectedHeaders { get; }
public IReadOnlyList<string> ActualHeaders { get; }


public OdsHeaderCountMismatchException(int expectedCount, int actualCount,
IEnumerable<string> expectedHeaders,
IEnumerable<string> actualHeaders)
: base($"欄位數量不一致：期望 {expectedCount} 欄，實際 {actualCount} 欄。期望表頭：[{string.Join(", ", expectedHeaders)}]；實際表頭：[{string.Join(", ", actualHeaders)}]")
{
ExpectedCount = expectedCount;
ActualCount = actualCount;
ExpectedHeaders = expectedHeaders.ToArray();
ActualHeaders = actualHeaders.ToArray();
}
}
