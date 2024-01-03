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

    public DuplicateHeaderException(IXLWorksheet worksheet, CellAddress address1, CellAddress address2, string header)
        : base(
            $"Duplicate header '{header}' found in '{worksheet.Name}' at '{address1}' and '{address2}'.")
    {
    }

    public DuplicateHeaderException(
        IXLWorksheet worksheet, CellAddress address1, CellAddress address2, string header, Exception innerException)
        : base(
            $"Duplicate header '{header}' found in '{worksheet.Name}' at '{address1}' and '{address2}'.",
            innerException)
    {
    }
}