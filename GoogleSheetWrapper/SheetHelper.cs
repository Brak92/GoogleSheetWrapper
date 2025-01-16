using Google.Apis.Sheets.v4.Data;
using GoogleSheetWrapper.Attributes;
using GoogleSheetWrapper.Interfaces;
using System.Reflection;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace GoogleSheetWrapper;
internal static class SheetHelper<T> where T : ISheet, new()
{
    public static IEnumerable<T> GetSheet(int skipFirstFewRows = 1, string additionalSheetName = "", short skipStartAmountColumns = 0, short skipEndAmountColumns = 0)
    {
        T instance = new();
        string range = GetRange(skipStartAmountColumns, skipEndAmountColumns);

        if (!string.IsNullOrWhiteSpace(additionalSheetName))
            if (additionalSheetName[..1] != " ")
                additionalSheetName = additionalSheetName.Insert(0, " ");

        string lookupRange = $"{instance.SheetName}{additionalSheetName}!{range}";

        ValueRange response = GoogleSheetService.Sheet
                                                .Spreadsheets
                                                .Values
                                                .Get(GoogleSheetService.SheetId, lookupRange)
                                                .Execute();

        IList<IList<object>> values = response.Values;

        List<T> result = MapFromRangeData(values, skipFirstFewRows);

        return result;
    }

    public static bool CreateSheet(string additionalSheetName = "")
    {
        T instance = new();
        string sheetName = $"{instance.SheetName} {additionalSheetName}".TrimEnd();

        if (SheetHelper.GetSheetNames(sheetName).Any())
            return true;

        AddSheetRequest addSheetRequest = new()
        {
            Properties = new()
            {
                Title = sheetName
            }
        };

        BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new()
        {
            Requests = []
        };

        batchUpdateSpreadsheetRequest.Requests.Add(new()
        {
            AddSheet = addSheetRequest
        });

        GoogleSheetService.Sheet
                          .Spreadsheets
                          .BatchUpdate(batchUpdateSpreadsheetRequest, GoogleSheetService.SheetId)
                          .Execute();

        return InsertValue(instance, true, additionalSheetName);
    }

    public static bool InsertValue(T instance, bool isHeaders = false, string additionalSheetName = "", int skipStartAmountColumns = 0, int skipEndAmountColumns = 0)
    {
        ValueRange value;
        string sheetName = $"{instance.SheetName} {additionalSheetName}".TrimEnd();
        string range = GetRange(skipStartAmountColumns, skipEndAmountColumns);
        string sheetRange = $"{sheetName}!{range}";

        if (isHeaders)
        {
            IOrderedEnumerable<PropertyInfo> typeProperties = typeof(T).GetProperties()
                                                                       .Where(prop => prop.GetCustomAttributes<SheetAttribute>().Any())
                                                                       .OrderBy(prop => prop.GetCustomAttribute<SheetAttribute>()!.ColumnIndex);

            foreach (PropertyInfo property in typeProperties)
            {
                property.SetValue(instance, property.Name);
            }

        }

        value = new()
        {
            Values = MapToRangeData(instance)
        };

        var request = GoogleSheetService.Sheet
                                        .Spreadsheets
                                        .Values.Append(value, GoogleSheetService.SheetId, sheetRange);
        request.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;
        request.Execute();

        return true;
    }

    public static bool InsertValues(IEnumerable<T> instanceCollection, string additionalSheetName = "", int skipStartAmountColumns = 0, int skipEndAmountColumns = 0, int skipFirstFewRows = 0)
    {
        AppendRequest? appendRequest;

        if (!instanceCollection.Any())
            return false;

        T instance = instanceCollection.First();
        string sheetName = $"{instance.SheetName} {additionalSheetName}";

        if (string.IsNullOrWhiteSpace(additionalSheetName))
            sheetName = sheetName.TrimEnd();

        string range = GetRange(skipStartAmountColumns, skipEndAmountColumns);
        string appendRange = $"{sheetName}!{range}";
        ValueRange appendValue = new()
        {
            Values = MapToBatchRangeData(instanceCollection)
        };

        appendRequest = GoogleSheetService.Sheet
                                          .Spreadsheets
                                          .Values
                                          .Append(appendValue, GoogleSheetService.SheetId, appendRange);

        appendRequest.ValueInputOption = AppendRequest.ValueInputOptionEnum.USERENTERED;

        if (skipFirstFewRows > 0)
            if (appendRequest is not null)
            {
                {
                    int firstRow = 1 + skipFirstFewRows;
                    string[] splitRange = range.Split(':');
                    string deleteRange = $"{sheetName}!{string.Join($"{firstRow}:", splitRange)}";
                    ClearValuesRequest body = new();
                    ClearRequest deleteRequest = GoogleSheetService.Sheet
                                                                   .Spreadsheets
                                                                   .Values
                                                                   .Clear(body, GoogleSheetService.SheetId, deleteRange);

                    deleteRequest.Execute();
                }

                appendRequest.Execute();
            }

        return true;
    }

    public static IList<IList<object>> MapToBatchRangeData(IEnumerable<T> items)
    {
        List<IList<object>> rangeData = [];
        foreach (T item in items)
        {
            List<object> lObject = [];
            IOrderedEnumerable<PropertyInfo> typeProperties = typeof(T).GetProperties()
                                                                       .Where(prop => prop.GetCustomAttributes<SheetAttribute>().Any())
                                                                       .OrderBy(prop => prop.GetCustomAttribute<SheetAttribute>()!.ColumnIndex);

            foreach (PropertyInfo property in typeProperties)
            {
                lObject.Add(property.GetValue(item)!);
            }

            rangeData.Add(lObject);
        }
        return rangeData;
    }

    private static List<T> MapFromRangeData(IList<IList<object>> values, int skipFirstFewRows)
    {
        List<T> items = [];

        foreach (var value in values.Where(value => !value.All(val => string.IsNullOrWhiteSpace(val?.ToString()))).Skip(skipFirstFewRows))
        {
            T item = new();

            IOrderedEnumerable<PropertyInfo> typeProperties = typeof(T).GetProperties().Where(prop => prop.GetCustomAttributes<SheetAttribute>().Any())
                                                                                       .OrderBy(prop => prop.GetCustomAttribute<SheetAttribute>()!.ColumnIndex);
            foreach (var property in typeProperties.Select((prop, i) => new { Property = prop, Index = i }))
            {
                Type type = property.Property.PropertyType;

                property.Property.SetValue(item, Convert.ChangeType(value[property.Index] ?? default, type));
            }

            items.Add(item);
        }

        return items;
    }

    private static List<IList<object>> MapToRangeData(T item)
    {
        List<IList<object>> rangeData = [];
        List<object> lObject = [];
        IOrderedEnumerable<PropertyInfo> typeProperties = typeof(T).GetProperties()
                                                                   .Where(prop => prop.GetCustomAttributes<SheetAttribute>().Any())
                                                                   .OrderBy(prop => prop.GetCustomAttribute<SheetAttribute>()!.ColumnIndex);

        foreach (var property in typeProperties)
        {
            lObject.Add(property.GetValue(item)!);
        }

        rangeData.Add(lObject);

        return rangeData;
    }

    private static string GetRange(int skipStartAmountColumns, int skipEndAmountColumns)
    {
        int columnAmount = typeof(T).GetProperties().Count(prop => prop.GetCustomAttributes<SheetAttribute>().Any()) - 1;
        string startColumnLetter = GetColumnName(skipStartAmountColumns);
        string endColumnLetter = GetColumnName(columnAmount - skipEndAmountColumns);

        return $"{startColumnLetter}:{endColumnLetter}";
    }

    private static string GetColumnName(int index)
    {
        const int A = 'A';
        int dividend = index + 1;
        string columnName = "";

        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(A + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }

        return columnName;
    }

}

internal static class SheetHelper
{
    public static IEnumerable<string> GetSheetNames(string title = "")
    {
        Google.Apis.Sheets.v4.SpreadsheetsResource.GetRequest request = GoogleSheetService.Sheet
                                                                                          .Spreadsheets
                                                                                          .Get(GoogleSheetService.SheetId);

        Spreadsheet spreadsheet = request.Execute();

        foreach (Sheet? sheet in spreadsheet.Sheets)
        {
            if (!sheet.Properties.Title.Contains(title))
                continue;

            yield return sheet.Properties.Title.Replace(title, "").Trim();
        }
    }
}
