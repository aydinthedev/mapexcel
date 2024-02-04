using MapExcel.Constants;

namespace MapExcel.Metadata;

internal sealed class ExcelType
{
    internal ExcelType(Type type)
    {
        ArgumentNullException.ThrowIfNull(type);

        Type = type;
        Properties = new List<ExcelProperty>();
        WorksheetNumber = WorksheetConstants.MinNumber;
        WorksheetName = WorksheetConstants.DefaultName;
        WorksheetMatch = WorksheetMatch.Default;
        Captions = new List<Caption>();
        HeaderRows = 0;
    }

    internal Type Type { get; }

    internal List<ExcelProperty> Properties { get; }

    internal int WorksheetNumber { get; set; }

    internal string WorksheetName { get; set; }

    internal WorksheetMatch WorksheetMatch { get; set; }

    internal List<Caption> Captions { get; }

    internal int HeaderRows { get; set; }

    internal bool HeaderAutoFilter { get; set; }
}