using System.Globalization;

namespace RetailStore.Data.DataScreens;

public enum FieldKind { Text, Memo, Money, Quantity, Rate, Integer, Bool, Date, DateTime, Choice, Lookup, Info }

/// <summary>A column of the record list: header (Arabic, through Tr), field, width in cm on the Access screen.</summary>
public sealed record ColumnSpec(string Header, string Field, decimal Width);

/// <summary>A choice of a value list: stored value and text.</summary>
public sealed record OptionItem(string Value, string Text);

/// <summary>A look-up list: the table, its key and the name shown (English name when the interface is English).</summary>
public sealed record LookupSpec(string Table, string KeyField, string NameField, string? NameEnField);

public sealed record LookupItem(object Key, string Text, bool Active);

/// <summary>A field of a data screen (forms.Fld + its schema field).</summary>
public sealed class FieldSpec<T> where T : class
{
    public required string Name { get; init; }
    public required FieldKind Kind { get; init; }
    /// <summary>The label on the screen (Arabic; the screen shows it through Tr).</summary>
    public required string Label { get; init; }
    /// <summary>The caption of the field in the back-end.</summary>
    public string Caption { get; init; } = "";
    public string Hint { get; init; } = "";
    public int Span { get; init; } = 1;
    public bool Locked { get; init; }
    public bool Required { get; init; }
    public int MaxLength { get; init; }
    public decimal? Min { get; init; }
    public bool MinExclusive { get; init; }
    public decimal? Max { get; init; }
    public bool MaxExclusive { get; init; }
    public OptionItem[]? Options { get; init; }
    public LookupSpec? Lookup { get; init; }
    public Func<T, object?> Get { get; init; } = _ => null;
    public Action<T, object?> Set { get; init; } = (_, _) => { };

    public bool IsData => Kind != FieldKind.Info;
}

public interface IScreenSpec
{
    string Name { get; }
    string Table { get; }
    string Route { get; }
    string Icon { get; }
    string Caption { get; }
    string Subtitle { get; }
}

/// <summary>A data screen (forms.DataScreen): list + search + detail of one table.</summary>
public sealed class ScreenSpec<T> : IScreenSpec where T : class, new()
{
    public required string Name { get; init; }
    public required string Table { get; init; }
    public required string Key { get; init; }
    public required string Route { get; init; }
    public string Icon { get; init; } = "";
    public required string Caption { get; init; }
    public string Subtitle { get; init; } = "";
    public required Func<T, int> KeyOf { get; init; }
    public required Func<RetailDbContext, object, Task<T?>> FindAsync { get; init; }
    public required Func<RetailDbContext, IQueryable<T>> Query { get; init; }
    public required Func<IQueryable<T>, IOrderedQueryable<T>> OrderBy { get; init; }
    public string[] Search { get; init; } = [];
    public string? ActiveField { get; init; }
    public string? Sequence { get; init; }
    public string? SequenceField { get; init; }
    public string[] Unique { get; init; } = [];
    /// <summary>The name field of a table with English names (master_en): the English interface shows the English one.</summary>
    public string? NameField { get; init; }
    public string? NameEnField { get; init; }
    public required ColumnSpec[] Columns { get; init; }
    public required FieldSpec<T>[] Fields { get; init; }

    public FieldSpec<T>? Field(string name) => Fields.FirstOrDefault(f => f.Name == name);
}

/// <summary>The conversions of the generated setters: a value typed on a screen -> the type of the field.</summary>
public static class ValueConvert
{
    private static bool Empty(object? v) => v is null || (v is string s && s.Trim().Length == 0);

    public static string Text(object? v) => v?.ToString() ?? "";
    public static string? TextN(object? v) => Empty(v) ? null : v!.ToString();
    public static int Int(object? v) => IntN(v) ?? 0;
    public static int? IntN(object? v) => Empty(v) ? null : Convert.ToInt32(v, CultureInfo.InvariantCulture);
    public static short Short(object? v) => ShortN(v) ?? 0;
    public static short? ShortN(object? v) => Empty(v) ? null : Convert.ToInt16(v, CultureInfo.InvariantCulture);
    public static byte Byte(object? v) => ByteN(v) ?? 0;
    public static byte? ByteN(object? v) => Empty(v) ? null : Convert.ToByte(v, CultureInfo.InvariantCulture);
    public static decimal Decimal(object? v) => DecimalN(v) ?? 0m;
    public static decimal? DecimalN(object? v) => Empty(v) ? null : Convert.ToDecimal(v, CultureInfo.InvariantCulture);
    public static bool Bool(object? v) => v is not null && Convert.ToBoolean(v, CultureInfo.InvariantCulture);
    public static DateTime Date(object? v) => DateN(v) ?? DateTime.Today;
    public static DateTime? DateN(object? v) => Empty(v) ? null : Convert.ToDateTime(v, CultureInfo.InvariantCulture);
}
