using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System.Reflection;

namespace GoogleSheetWrapper;
public static partial class GoogleSheetService
{
    private static string _credentialFile = default!;
    private static string _applicationName = default!;
    private static string _sheetId = default!;
    private static SheetsService _sheet = default!;
    public static SheetsService Sheet => _sheet ??= AuthorizeGoogleApp();

    public static string SheetId => _sheetId;

    /// <summary>
    /// Register the Json File that contains the credential information so it can be used by <see cref="AuthorizeGoogleApp"/>.
    /// </summary>
    /// <param name="credentialJsonFileName">The file name that is registered as a Resource File that will be read with <see cref="Assembly.GetManifestResourceStream(string)"/>.</param>
    public static void RegisterCredentialJsonFile(string credentialJsonFileName) => 
        _credentialFile = credentialJsonFileName;

    /// <summary>
    /// Registeres the name of the application to be read within the <see cref="SheetsService.SheetsService(Google.Apis.Services.BaseClientService.Initializer)"/> in the property <see cref="Google.Apis.Services.BaseClientService.ApplicationName"/>
    /// and the Sheet Id to be used within the <see cref="SheetHelper{T}"/> and <see cref="SheetHelper"/>
    /// </summary>
    /// <param name="applicationName">Name of the application set within google settings.</param>
    /// <param name="sheetId">The sheet id for the sheet you're going to use. Usually can be seen within the URL when accessing the spreadsheet. Looks like a big amount of random letters and numbers.</param>
    public static void RegisterSheet(string applicationName, string sheetId) => 
        (_applicationName, _sheetId) = (applicationName, sheetId);

    /// <summary>
    /// Uses the credentials set in the credential files registered with <see cref="RegisterCredentialJsonFile(string)"/>
    /// </summary>
    /// <returns>An instance of the <see cref="SheetsService"/> using the application name set in <see cref="RegisterSheet(string, string)"/> </returns>
    private static SheetsService AuthorizeGoogleApp()
    {
        Validate();
        
        GoogleCredential credential;
        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(_credentialFile))
        {
            credential = GoogleCredential.FromStream(stream!).CreateScoped(SheetsService.Scope.Spreadsheets);
        }

        return new SheetsService(new()
        {
            HttpClientInitializer = credential,
            ApplicationName = _applicationName
        });
    }

    /// <summary>
    /// Tests if the information has been set before authorizing
    /// </summary>
    /// <exception cref="ArgumentException"></exception>
    private static void Validate()
    {
        List<string> errorMessage = [];

        if (string.IsNullOrWhiteSpace(_credentialFile))
            errorMessage.Add("Credential File has not been set.");
        else if (_credentialFile.Split('.').LastOrDefault() != "json")
            errorMessage.Add("Credential File is not a JSON file.");

        if (string.IsNullOrWhiteSpace(_applicationName))
            errorMessage.Add("Application Name has not been set.");

        if (string.IsNullOrWhiteSpace(_sheetId))
            errorMessage.Add("Sheet Id has not been set.");

        if(errorMessage.Count > 0)
            throw new ArgumentException(string.Join(Environment.NewLine, errorMessage));
    }
}