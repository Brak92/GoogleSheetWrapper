using GoogleSheetWrapper.Interfaces;

namespace GoogleSheetWrapper;
public static partial class GoogleSheetService
{
    /// <summary>
    /// Returns the all the values of a sheet with the same name as <see cref="ISheet.SheetName"/>
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <returns>Returns a collection with all the values of a sheet</returns>
    public static IEnumerable<T> GetSheet<T>() where T : ISheet, new() =>
        GetSheet<T>(1, "", 0, 0);


    /// <summary>
    /// Returns the all the values of a sheet with the same name as <see cref="ISheet.SheetName"/>
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="skipFirstFewRows">Amount of rows to skip, useful to skip header rows</param>
    /// <returns>Returns a collection with all the values of a sheet</returns>
    public static IEnumerable<T> GetSheet<T>(int skipFirstFewRows) where T : ISheet, new() =>
        GetSheet<T>(skipFirstFewRows, "", 0, 0);

    /// <summary>
    /// Returns the all the values of a sheet with the same name as <see cref="ISheet.SheetName"/>
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="skipFirstFewRows">Amount of rows to skip, useful to skip header rows</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <returns>Returns a collection with all the values of a sheet</returns>
    public static IEnumerable<T> GetSheet<T>(int skipFirstFewRows, string additionalSheetName) where T : ISheet, new() =>
        GetSheet<T>(skipFirstFewRows, additionalSheetName, 0, 0);

    /// <summary>
    /// Returns the all the values of a sheet with the same name as <see cref="ISheet.SheetName"/>
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="skipFirstFewRows">Amount of rows to skip, useful to skip header rows</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <param name="skipStartAmountColumns">Amount of columns from the left to skip</param>
    /// <returns>Returns a collection with all the values of a sheet</returns>
    public static IEnumerable<T> GetSheet<T>(int skipFirstFewRows, string additionalSheetName, short skipStartAmountColumns) where T : ISheet, new() =>
        GetSheet<T>(skipFirstFewRows, additionalSheetName, skipStartAmountColumns, 0);

    /// <summary>
    /// Returns the all the values of a sheet with the same name as <see cref="ISheet.SheetName"/>
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="skipFirstFewRows">Amount of rows to skip, useful to skip header rows</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <param name="skipStartAmountColumns">Amount of columns from the left to skip</param>
    /// <param name="skipEndAmountColumns">Amount of columns from the right to skip, starting from the rightmost column set with the <see cref="Attributes.SheetAttribute.ColumnIndex"/></param>
    /// <returns>Returns a collection with all the values of a sheet</returns>
    public static IEnumerable<T> GetSheet<T>(int skipFirstFewRows = 1, string additionalSheetName = "", short skipStartAmountColumns = 0, short skipEndAmountColumns = 0) where T : ISheet, new()
    {
        if (skipFirstFewRows < 0)
            skipFirstFewRows = 0;

        if (skipStartAmountColumns < 0)
            skipStartAmountColumns = 0;

        if (skipEndAmountColumns < 0)
            skipEndAmountColumns = 0;

        additionalSheetName ??= "";

        return SheetHelper<T>.GetSheet(skipFirstFewRows, additionalSheetName, skipStartAmountColumns, skipEndAmountColumns);
    }
}
