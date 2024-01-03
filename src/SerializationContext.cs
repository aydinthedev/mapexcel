using System.Reflection;
using ClosedXML.Excel;

namespace MapExcel;

public class SerializationContext
{
    public SerializationContext(WorksheetOptions worksheetOptions, PropertyInfo propertyInfo, IXLStyle cellStyle)
    {
        WorksheetOptions = worksheetOptions;
        PropertyInfo = propertyInfo;
        CellStyle = cellStyle;
    }

    public WorksheetOptions WorksheetOptions { get; }

    public PropertyInfo PropertyInfo { get; }

    public IXLStyle CellStyle { get; }
}

public class DeserializationContext
{
    public DeserializationContext(WorksheetOptions worksheetOptions, PropertyInfo propertyInfo, IXLStyle cellStyle)
    {
        WorksheetOptions = worksheetOptions;
        PropertyInfo = propertyInfo;
        CellStyle = cellStyle;
        IsValid = true;
    }

    public WorksheetOptions WorksheetOptions { get; }

    public PropertyInfo PropertyInfo { get; }

    public IXLStyle CellStyle { get; }

    /// <summary>
    /// Set this property if deserialization failed.
    /// This will used to invalidate the whole row.
    /// </summary>
    public bool IsValid { get; set; }
}