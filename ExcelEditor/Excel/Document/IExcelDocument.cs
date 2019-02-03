using ExcelEditor.Excel.Elements;
using OfficeOpenXml;

namespace ExcelEditor.Excel.Document
{
    public interface IExcelDocument
    {
        string FileName { get; }

        ExcelPackage ExcelPackage { get; }

        ExcelWorksheet ActiveWorksheet { get; set; }

        Range Selection { get; set; }

        Range GetActiveRange(string addressText);
    }
}
