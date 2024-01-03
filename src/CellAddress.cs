using System.Text;
using MapExcel.Constants;

namespace MapExcel;

public readonly struct CellAddress : IEquatable<CellAddress>
{
    public CellAddress()
    {
        RowNumber = 1;
        ColumnNumber = 1;
    }

    public CellAddress(int rowNumber, int columnNumber)
    {
        if (rowNumber is < WorksheetConstants.MinRowNumber or > WorksheetConstants.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(rowNumber));

        if (columnNumber is < WorksheetConstants.MinColumnNumber or > WorksheetConstants.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnNumber));

        RowNumber = rowNumber;
        ColumnNumber = columnNumber;
    }

    public CellAddress(string columnName, int rowNumber)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be empty or null.", nameof(columnName));

        var columnNumber = ToColumnNumber(columnName);

        if (columnNumber is < WorksheetConstants.MinColumnNumber or > WorksheetConstants.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnName));

        if (rowNumber is < WorksheetConstants.MinRowNumber or > WorksheetConstants.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(rowNumber));

        RowNumber = rowNumber;
        ColumnNumber = columnNumber;
    }

    public CellAddress(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            throw new ArgumentException("Address cannot be empty or null.", nameof(address));

        var lastLetterNumber = 0;
        while (lastLetterNumber < address.Length && char.IsLetter(address[lastLetterNumber]))
            lastLetterNumber++;

        var columnNumber = ToColumnNumber(address[..lastLetterNumber]);

        if (columnNumber is < WorksheetConstants.MinColumnNumber or > WorksheetConstants.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(address));

        if (!int.TryParse(address[lastLetterNumber..], out var rowNumber))
            throw new ArgumentException($"Value '{address}' is not a valid cell address.", nameof(address));

        if (rowNumber is < WorksheetConstants.MinRowNumber or > WorksheetConstants.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(address));

        RowNumber = rowNumber;
        ColumnNumber = columnNumber;
    }

    public int RowNumber { get; }

    public int ColumnNumber { get; }

    public string Address => ToColumnName(ColumnNumber) + RowNumber;

    /// <summary>
    ///     Returns the column Number for the given column name.
    /// </summary>
    internal static int ToColumnNumber(string columnName) =>
        columnName.Aggregate(0, (current, c) => current * 26 + c - 'A' + 1);

    /// <summary>
    ///     Returns the column name for the given column Number.
    /// </summary>
    internal static string ToColumnName(int columnNumber)
    {
        const int asciiOffset = 65;
        var dividend = columnNumber;
        var columnNameBuilder = new StringBuilder();

        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            columnNameBuilder.Insert(0, (char)(asciiOffset + modulo));
            dividend = (dividend - modulo) / 26;
        }

        return columnNameBuilder.ToString();
    }

    public bool Equals(CellAddress other) => RowNumber == other.RowNumber && ColumnNumber == other.ColumnNumber;

    public override bool Equals(object? obj) => obj is CellAddress address && Equals(address);

    public override int GetHashCode() => HashCode.Combine(RowNumber, ColumnNumber);

    public static bool operator ==(CellAddress left, CellAddress right) => left.Equals(right);

    public static bool operator !=(CellAddress left, CellAddress right) => !left.Equals(right);

    public override string ToString() => Address;
}