using ClosedXML.Excel;

namespace MapExcel.Metadata.Builders;

public sealed class CaptionBuilder
{
    private readonly Caption _caption;

    internal CaptionBuilder(Caption caption)
    {
        ArgumentNullException.ThrowIfNull(caption);
        _caption = caption;
    }

    public CaptionBuilder Range(
        int firstRowNumber, int firstColumnNumber, int lastRowNumber, int lastColumnNumber)
    {
        _caption.AddressRange =
            new CellAddressRange(firstRowNumber, firstColumnNumber, lastRowNumber, lastColumnNumber);
        return this;
    }

    public CaptionBuilder Range(string firstCellAddress, string lastCellAddress)
    {
        _caption.AddressRange = new CellAddressRange(firstCellAddress, lastCellAddress);
        return this;
    }

    public CaptionBuilder Range(string rangeAddress)
    {
        _caption.AddressRange = new CellAddressRange(rangeAddress);
        return this;
    }

    public CaptionBuilder Range(CellAddressRange addressRange)
    {
        _caption.AddressRange = addressRange;
        return this;
    }

    public CaptionBuilder Name(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Name cannot be empty or null.", nameof(name));

        _caption.Name = name.Trim();
        return this;
    }

    public CaptionBuilder Style(Action<IXLStyle> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        _caption.Style = configure;
        return this;
    }
}