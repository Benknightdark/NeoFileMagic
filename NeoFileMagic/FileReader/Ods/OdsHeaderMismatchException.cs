using System;

namespace NeoFileMagic.FileReader.Ods;

public sealed class OdsHeaderMismatchException : Exception
{
public IReadOnlyList<string> ExpectedHeaders { get; }
public IReadOnlyList<string> MissingHeaders { get; }
public IReadOnlyList<string> AvailableHeaders { get; }


public OdsHeaderMismatchException(
IEnumerable<string> expected,
IEnumerable<string> missing,
IEnumerable<string> available)
: base(BuildMessage(expected, missing, available))
{
ExpectedHeaders = expected.ToArray();
MissingHeaders = missing.ToArray();
AvailableHeaders = available.ToArray();
}


private static string BuildMessage(IEnumerable<string> expected, IEnumerable<string> missing, IEnumerable<string> available)
=> $"表頭不一致：缺少 [{string.Join(", ", missing)}]。期望表頭：[{string.Join(", ", expected)}]；實際表頭：[{string.Join(", ", available)}]";
}
