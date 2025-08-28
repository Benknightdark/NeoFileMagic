namespace NeoFileMagic.FileReader.Ods;

using System;
using System.Globalization;

public readonly struct OdsCell
{
    public static readonly OdsCell Empty = new();

    public OdsValueType Type { get; init; }
    public string? Text { get; init; }
    public double? Number { get; init; }
    public string? Currency { get; init; }
    public DateTimeOffset? DateTime { get; init; }
    public TimeSpan? Time { get; init; }
    public bool? Boolean { get; init; }
    public string? Formula { get; init; }

    public override string ToString() => Type switch
    {
        OdsValueType.String => Text ?? string.Empty,
        OdsValueType.Float or OdsValueType.Currency => Number?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
        OdsValueType.Boolean => Boolean?.ToString() ?? string.Empty,
        OdsValueType.Date => DateTime?.ToString("o") ?? string.Empty,
        OdsValueType.Time => Time?.ToString() ?? string.Empty,
        _ => string.Empty,
    };
}
