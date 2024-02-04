using MapExcel.Constants;

namespace MapExcel.Metadata.Builders;

public sealed class WorksheetBuilder
{
    private readonly ExcelType _excelType;

    internal WorksheetBuilder(ExcelType excelType)
    {
        ArgumentNullException.ThrowIfNull(excelType);
        _excelType = excelType;
    }

    public WorksheetBuilder Number(int number)
    {
        if (number < WorksheetConstants.MinNumber)
            throw new ArgumentOutOfRangeException(nameof(number));

        _excelType.WorksheetNumber = number;
        return this;
    }

    public WorksheetBuilder Name(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Name cannot be empty or null.", nameof(name));

        _excelType.WorksheetName = name.Trim();
        return this;
    }

    public WorksheetBuilder Match(WorksheetMatch match)
    {
        ArgumentNullException.ThrowIfNull(match);
        _excelType.WorksheetMatch = match;
        return this;
    }

    public WorksheetBuilder AddCaption(Action<CaptionBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);

        var caption = new Caption();
        var builder = new CaptionBuilder(caption);
        configure(builder);

        _excelType.Captions.Add(caption);
        return this;
    }
}