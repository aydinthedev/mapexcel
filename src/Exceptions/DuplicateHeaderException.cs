using ClosedXML.Excel;

namespace MapExcel.Exceptions;

public class DuplicateHeaderException : Exception
{
    public DuplicateHeaderException(string message)
        : base(message)
    {
    }

    public DuplicateHeaderException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public DuplicateHeaderException(
        IXLWorksheet worksheet,
        CellAddress address,
        CellAddress duplicateAddress,
        string header)
        : base(
            $"Duplicate header '{header}' found in '{worksheet.Name}' at '{address}' and '{duplicateAddress}'.")
    {
    }

    public DuplicateHeaderException(
        IXLWorksheet worksheet,
        CellAddress address,
        CellAddress duplicateAddress,
        string header,
        Exception innerException)
        : base(
            $"Duplicate header '{header}' found in '{worksheet.Name}' at '{address}' and '{duplicateAddress}'.",
            innerException)
    {
    }
}