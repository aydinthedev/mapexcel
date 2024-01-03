using System.Globalization;

namespace MapExcel;

public class WorksheetOptions
{
    public static readonly WorksheetOptions Default = new();

    public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;

    public TimeZoneInfo TimeZone { get; set; } = TimeZoneInfo.Local;

    public DateTimeKind DateTimeKind { get; set; } = DateTimeKind.Unspecified;

    public NumberStyles NumberStyles { get; set; } = NumberStyles.Any;

    public string CollectionSeparator { get; set; } = ", ";

    public StringFrom ConvertStringFrom { get; set; } = StringFrom.NonText;

    public CharFrom ConvertCharFrom { get; set; } = CharFrom.Number | CharFrom.Text;

    public BooleanFrom ConvertBooleanFrom { get; set; } = BooleanFrom.Number | BooleanFrom.Text;

    public DateTimeFrom ConvertDateTimeFrom { get; set; } = DateTimeFrom.Number | DateTimeFrom.Text;

    public NumberFrom ConvertNumberFrom { get; set; } = NumberFrom.Text | NumberFrom.TimeSpan | NumberFrom.DateTime;
}

[Flags]
public enum StringFrom
{
    None = 0,
    NonText = 1
}

[Flags]
public enum CharFrom
{
    None = 0,
    Number = 1,
    Text = 2
}

[Flags]
public enum BooleanFrom
{
    None = 0,
    Number = 1,
    Text = 2
}

[Flags]
public enum DateTimeFrom
{
    None = 0,
    Number = 1,
    Text = 2
}

[Flags]
public enum NumberFrom
{
    None = 0,
    Text = 1,
    TimeSpan = 2,
    DateTime = 4
}