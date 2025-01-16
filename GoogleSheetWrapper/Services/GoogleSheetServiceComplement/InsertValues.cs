using GoogleSheetWrapper.Interfaces;

namespace GoogleSheetWrapper;
public static partial class GoogleSheetService
{

    /// <summary>
    /// Adds a range of values to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instanceCollection">The instance of the model that will be added to the sheet</param>
    /// <returns>True if any row was added, false if it failed</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static bool InsertValues<T>(IEnumerable<T> instanceCollection) where T : ISheet, new() =>
        InsertValues(instanceCollection, "", 0, 0, 0);

    /// <summary>
    /// Adds a range of values to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instanceCollection">The instance of the model that will be added to the sheet</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <returns>True if any row was added, false if it failed</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static bool InsertValues<T>(IEnumerable<T> instanceCollection, string additionalSheetName) where T : ISheet, new() =>
        InsertValues(instanceCollection, additionalSheetName, 0, 0, 0);

    /// <summary>
    /// Adds a range of values to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instanceCollection">The instance of the model that will be added to the sheet</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <param name="skipStartAmountColumns">Amount of columns from the left to skip</param>
    /// <returns>True if any row was added, false if it failed</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static bool InsertValues<T>(IEnumerable<T> instanceCollection, string additionalSheetName, int skipStartAmountColumns) where T : ISheet, new() =>
        SheetHelper<T>.InsertValues(instanceCollection, additionalSheetName, skipStartAmountColumns, 0, 0);

    /// <summary>
    /// Adds a range of values to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instanceCollection">The instance of the model that will be added to the sheet</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <param name="skipStartAmountColumns">Amount of columns from the left to skip</param>
    /// <param name="skipEndAmountColumns">Amount of columns from the right to skip, starting from the rightmost column set with the <see cref="Attributes.SheetAttribute.ColumnIndex"/></param>
    /// <returns>True if any row was added, false if it failed</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static bool InsertValues<T>(IEnumerable<T> instanceCollection, string additionalSheetName, int skipStartAmountColumns, int skipEndAmountColumns) where T : ISheet, new() =>
        SheetHelper<T>.InsertValues(instanceCollection, additionalSheetName, skipStartAmountColumns, skipEndAmountColumns, 0);

    /// <summary>
    /// Adds a range of values to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instanceCollection">The instance of the model that will be added to the sheet</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <param name="skipStartAmountColumns">Amount of columns from the left to skip</param>
    /// <param name="skipEndAmountColumns">Amount of columns from the right to skip, starting from the rightmost column set with the <see cref="Attributes.SheetAttribute.ColumnIndex"/></param>
    /// <param name="skipFirstFewRows">When this value is set higher than 0, it will replace the values after the amount of rows declared in this parameter, when setting to 1, it will skip the header. Setting to <= 0 will just append the values after the last line inserted in the sheet</param>
    /// <returns>True if any row was added, false if it failed</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static bool InsertValues<T>(IEnumerable<T> instanceCollection, string additionalSheetName = "", int skipStartAmountColumns = 0, int skipEndAmountColumns = 0, int skipFirstFewRows = 0) where T : ISheet, new()
    {
        ArgumentNullException.ThrowIfNull(instanceCollection);

        if (skipStartAmountColumns < 0)
            skipStartAmountColumns = 0;

        if (skipEndAmountColumns < 0)
            skipEndAmountColumns = 0;

        if(skipFirstFewRows < 0)
            skipFirstFewRows = 0;

        additionalSheetName ??= "";

        return SheetHelper<T>.InsertValues(instanceCollection, additionalSheetName, skipStartAmountColumns, skipEndAmountColumns, skipFirstFewRows);
    }

}
