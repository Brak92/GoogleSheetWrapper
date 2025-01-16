namespace GoogleSheetWrapper.Interfaces;
public interface ISheet
{
    /// <summary>
    /// The name of the sheet, accepts most of the ascii characters that a string understands
    /// </summary>
    public string SheetName { get; }
}