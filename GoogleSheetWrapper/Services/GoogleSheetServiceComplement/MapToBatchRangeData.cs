using GoogleSheetWrapper.Interfaces;

namespace GoogleSheetWrapper;
public static partial class GoogleSheetService
{

    /// <summary>
    /// Maps the collection to the format of data that GoogleSheets expects
    /// </summary>
    /// <typeparam name="T">A model sheet that inherits from <see cref="ISheet"/></typeparam>
    /// <param name="items">The collection to be mapped</param>
    /// <returns>The mapped data from the collection informed</returns>
    /// <exception cref="ArgumentNullException"></exception>
    public static IList<IList<object>> MapToBatchRangeData<T>(IEnumerable<T> items) where T : ISheet, new()
    {
        ArgumentNullException.ThrowIfNull(items);

       return SheetHelper<T>.MapToBatchRangeData(items);
    }
}
