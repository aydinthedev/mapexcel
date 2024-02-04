namespace MapExcel;

public readonly struct CellAddressRange : IEquatable<CellAddressRange>
{
    public CellAddressRange(int firstRowNumber, int firstColumnNumber, int lastRowNumber, int lastColumnNumber)
    {
        FirstCellAddress = new CellAddress(firstRowNumber, firstColumnNumber);
        LastCellAddress = new CellAddress(lastRowNumber, lastColumnNumber);
    }

    public CellAddressRange(string firstCellAddress, string lastCellAddress)
    {
        FirstCellAddress = new CellAddress(firstCellAddress);
        LastCellAddress = new CellAddress(lastCellAddress);
    }

    public CellAddressRange(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
            throw new ArgumentException("Range address cannot be empty or null.", nameof(rangeAddress));

        var split = rangeAddress.Split(":");
        if (split.Length != 2)
            throw new ArgumentException($"Value '{rangeAddress}' is not a valid range address.", nameof(rangeAddress));

        FirstCellAddress = new CellAddress(split[0]);
        LastCellAddress = new CellAddress(split[1]);
    }

    public CellAddress FirstCellAddress { get; }

    public CellAddress LastCellAddress { get; }

    public string Address => string.Concat(FirstCellAddress.Address, ":", LastCellAddress);

    public bool Equals(CellAddressRange other) =>
        FirstCellAddress == other.FirstCellAddress && LastCellAddress == other.LastCellAddress;

    public override bool Equals(object? obj) => obj is CellAddressRange other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(FirstCellAddress, LastCellAddress);

    public static bool operator ==(CellAddressRange left, CellAddressRange right) => left.Equals(right);

    public static bool operator !=(CellAddressRange left, CellAddressRange right) => !left.Equals(right);

    public override string ToString() => Address;
}