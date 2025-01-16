namespace GoogleSheetWrapper.Attributes;

[AttributeUsage(AttributeTargets.Property)]
public class SheetAttribute(int columnIndex) : Attribute
{
    /// <summary>
    /// Column A = 0 ColumnIndex, self explanatory after that
    /// </summary>
    public int ColumnIndex { get; private set; } = columnIndex;

}