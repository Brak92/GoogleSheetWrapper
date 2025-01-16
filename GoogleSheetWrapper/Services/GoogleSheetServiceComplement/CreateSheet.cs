using GoogleSheetWrapper.Interfaces;

namespace GoogleSheetWrapper;
public static partial class GoogleSheetService
{
    /// <summary>
    /// Create a new Sheet based on the name of <see cref="ISheet.SheetName"/>
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <returns>True if the sheet was created succesfully, False if not</returns>
    public static bool CreateSheet<T>() where T : ISheet, new() =>
        CreateSheet<T>("");

    /// <summary>
    /// Create a new Sheet based on the name of <see cref="ISheet.SheetName"/>
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="additionalSheetName">An extra name for the sheet that goes after the value of <see cref="ISheet.SheetName"/> separated by a blank space.</param>
    /// <returns>True if the sheet was created succesfully, False if not</returns>
    public static bool CreateSheet<T>(string additionalSheetName = "") where T : ISheet, new()
    {
        additionalSheetName ??= "";

        return SheetHelper<T>.CreateSheet(additionalSheetName);
    }
}
