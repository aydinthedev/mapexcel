namespace MapExcel.Metadata;

public class CaptionMetadata
{
    // Prevent outside initialization
    internal CaptionMetadata()
    {
    }

    public string? Name { get; init; }

    public CellAddressRange ExpectedAt { get; init; }

    public CellAddressRange? FoundAt { get; internal set; }

    public bool Validate(bool checkPosition = false) =>
        FoundAt != null && (!checkPosition || ExpectedAt.Equals(FoundAt.Value));
}