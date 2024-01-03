namespace MapExcel;

public enum WorksheetMatch
{
    /// <summary>
    ///     Ensure the worksheet exists at the specified number
    /// </summary>
    ByNumber,

    /// <summary>
    ///     Ensure the worksheet exists with the specified name.
    /// </summary>
    ByName,

    /// <summary>
    ///     Ensure the worksheet exists at the specified number and name.
    /// </summary>
    ByNumberAndName,

    /// <summary>
    ///     Prefer the worksheet at the specified number; fallback to the specified name if not found.
    /// </summary>
    PreferNumber,

    /// <summary>
    ///     Prefer the worksheet with the specified name; fallback to the specified number if not found.
    /// </summary>
    PreferName,

    /// <summary>
    ///     Use a default strategy if not explicitly specified. Default is <see cref="ByNumber" />.
    /// </summary>
    Default = ByNumber
}