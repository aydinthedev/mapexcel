namespace MapExcel;

public enum ColumnMatch
{
    /// <summary>
    ///     Match by column name.
    /// </summary>
    ByColumn,

    /// <summary>
    ///     Match by column header.
    /// </summary>
    ByHeader,

    /// <summary>
    ///     Match by column name and header.
    /// </summary>
    ByColumnAndHeader,

    /// <summary>
    ///     Use a default strategy if not explicitly specified. Default is <see cref="ByColumn" />.
    /// </summary>
    Default = ByColumn
}