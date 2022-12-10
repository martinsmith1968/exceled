namespace ExcelEditor.Lib.Excel.Elements
{
    public interface IRange
    {
        string Reference { get; }
        int Row { get; }
        int Column { get; }
    }
}
