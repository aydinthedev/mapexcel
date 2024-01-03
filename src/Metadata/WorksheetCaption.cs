using ClosedXML.Excel;

namespace MapExcel.Metadata;

internal sealed class WorksheetCaption
{
    // Initialize the AddressRange property to ensure it is not set to default values.
    // This prevents it from having zeros for addresses.
    internal WorksheetCaption()
    {
        AddressRange = new CellAddressRange(1, 1, 1, 1);
    }

    internal CellAddressRange AddressRange { get; set; }

    internal string? Name { get; set; }

    internal Action<IXLStyle>? Style { get; set; }
}