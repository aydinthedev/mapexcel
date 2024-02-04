namespace MapExcel;

public enum HeaderMatch
{
    /// <summary>
    ///     Find header by matching name. Header position is not considered.
    /// </summary>
    NameOnly,

    /// <summary>
    ///     Find header by matching name and position. Header must be found at the expected position.
    /// </summary>
    NameAndColumn,

    /// <summary>
    ///     Use a default strategy if not explicitly specified. Default is <see cref="NameOnly" />.
    /// </summary>
    Default = NameOnly
}