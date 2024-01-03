namespace MapExcel.Extensions;

internal static class DateTimeExtensions
{
    internal static DateTime AdjustToKind(
        this DateTime source,
        TimeZoneInfo sourceTimeZone,
        DateTimeKind newKind)
    {
        return source.Kind switch
        {
            DateTimeKind.Local => newKind switch
            {
                DateTimeKind.Local => source,
                DateTimeKind.Utc => source.ToUniversalTime(),
                _ => source
            },

            DateTimeKind.Utc => newKind switch
            {
                DateTimeKind.Local => source.ToLocalTime(),
                DateTimeKind.Utc => source,
                _ => source
            },

            DateTimeKind.Unspecified => newKind switch
            {
                DateTimeKind.Local =>
                    TimeZoneInfo.ConvertTime(source, sourceTimeZone, TimeZoneInfo.Local),
                DateTimeKind.Utc =>
                    TimeZoneInfo.ConvertTimeToUtc(source, sourceTimeZone),
                _ => source
            },

            _ => throw new ArgumentOutOfRangeException(nameof(source))
        };
    }

    internal static DateTime AdjustToTimeZone(
        this DateTime source,
        TimeZoneInfo destinationTimeZone)
    {
        switch (source.Kind)
        {
            case DateTimeKind.Local:
                var fromLocal = TimeZoneInfo
                    .ConvertTime(source, TimeZoneInfo.Local, destinationTimeZone);

                return DateTime.SpecifyKind(fromLocal, DateTimeKind.Unspecified);

            case DateTimeKind.Utc:
                var fromUtc = TimeZoneInfo
                    .ConvertTimeFromUtc(source, destinationTimeZone);

                return DateTime.SpecifyKind(fromUtc, DateTimeKind.Unspecified);

            case DateTimeKind.Unspecified:
                return source;

            default:
                throw new ArgumentOutOfRangeException(nameof(source));
        }
    }
}