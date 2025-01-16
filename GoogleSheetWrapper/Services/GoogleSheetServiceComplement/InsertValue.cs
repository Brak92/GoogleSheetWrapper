using GoogleSheetWrapper.Interfaces;

namespace GoogleSheetWrapper;
public static partial class GoogleSheetService
{
    /// <summary>
    /// Adds a value to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instance">The instance of the model that will be added to the sheet</param>
    /// <returns>True if any row was added, false if it failed</returns>
    public static bool InsertValue<T>(T instance) where T : ISheet, new() =>
        InsertValue(instance, false, "", 0, 0);

    /// <summary>
    /// Adds a value to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instance">The instance of the model that will be added to the sheet</param>
    /// <param name="isHeaders">Settings this to true will add only the headers based on the name of the properties of the <see cref="ISheet"/> model, disregarding the current values</param>
    /// <returns>True if any row was added, false if it failed</returns>
    public static bool InsertValue<T>(T instance, bool isHeaders) where T : ISheet, new() =>
        InsertValue(instance, isHeaders, "", 0, 0);

    /// <summary>
    /// Adds a value to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instance">The instance of the model that will be added to the sheet</param>
    /// <param name="isHeaders">Settings this to true will add only the headers based on the name of the properties of the <see cref="ISheet"/> model, disregarding the current values</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <returns>True if any row was added, false if it failed</returns>
    public static bool InsertValue<T>(T instance, bool isHeaders, string additionalSheetName) where T : ISheet, new() =>
        InsertValue(instance, isHeaders, additionalSheetName, 0, 0);

    /// <summary>
    /// Adds a value to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instance">The instance of the model that will be added to the sheet</param>
    /// <param name="isHeaders">Settings this to true will add only the headers based on the name of the properties of the <see cref="ISheet"/> model, disregarding the current values</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <param name="skipStartAmountColumns">Amount of columns from the left to skip</param>
    /// <returns>True if any row was added, false if it failed</returns>
    public static bool InsertValue<T>(T instance, bool isHeaders, string additionalSheetName, int skipStartAmountColumns) where T : ISheet, new() =>
        InsertValue(instance, isHeaders, additionalSheetName, skipStartAmountColumns, 0);

    /// <summary>
    /// Adds a value to the bottom of the sheet
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="instance">The instance of the model that will be added to the sheet</param>
    /// <param name="isHeaders">Settings this to true will add only the headers based on the name of the properties of the <see cref="ISheet"/> model, disregarding the current values</param>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <param name="skipStartAmountColumns">Amount of columns from the left to skip</param>
    /// <param name="skipEndAmountColumns">Amount of columns from the right to skip, starting from the rightmost column set with the <see cref="Attributes.SheetAttribute.ColumnIndex"/></param>
    /// <returns>True if any row was added, false if it failed</returns>
    public static bool InsertValue<T>(T instance, bool isHeaders = false, string additionalSheetName = "", int skipStartAmountColumns = 0, int skipEndAmountColumns = 0) where T : ISheet, new() =>
        SheetHelper<T>.InsertValue(instance, isHeaders, additionalSheetName, skipStartAmountColumns, skipEndAmountColumns);
}
