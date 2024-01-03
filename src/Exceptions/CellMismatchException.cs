namespace MapExcel.Exceptions;

public class CellMismatchException : Exception
{
    public CellMismatchException(string message)
        : base(message)
    {
    }

    public CellMismatchException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public CellMismatchException(CellAddress address, string header)
        : base($"Cell at '{address}' does not match with header '{header}'. ")
    {
    }

    public CellMismatchException(CellAddress address, string header, Exception innerException)
        : base($"Cell at '{address}' does not match with header '{header}'. ", innerException)
    {
    }
}