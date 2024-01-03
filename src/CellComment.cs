using ClosedXML.Excel;

namespace MapExcel;

public class CellComment
{
    private readonly IXLCell _cell;

    public CellComment(IXLCell cell)
    {
        _cell = cell;
    }

    public IXLComment Comment => _cell.HasComment ? _cell.GetComment() : _cell.CreateComment();
}