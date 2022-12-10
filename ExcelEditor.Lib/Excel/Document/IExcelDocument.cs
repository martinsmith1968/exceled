using ExcelEditor.Lib.Excel.Elements;
using OfficeOpenXml;

namespace ExcelEditor.Lib.Excel.Document
{
    public interface IExcelDocument
    {
        string FileName { get; }

        ExcelPackage ExcelPackage { get; }

        ExcelWorksheet ActiveWorksheet { get; set; }

        IRange Selection { get; set; }

        IRange GetActiveRange(string addressText);

        IRange GetSetActiveRange(string addressText);
    }
}
