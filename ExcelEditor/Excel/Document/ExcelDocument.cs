using ExcelEditor.Excel.Elements;
using OfficeOpenXml;

namespace ExcelEditor.Excel.Document
{
    public class ExcelDocument  : IExcelDocument
    {
        public string FileName { get; set; }

        public ExcelPackage ExcelPackage { get; set;  }

        public ExcelWorksheet ActiveWorksheet { get; set; }
        public Range Selection { get; set; }

        public Range GetActiveRange(string addressText)
        {
            var range = Range.Parse(addressText) ?? Selection;

            return range;
        }

        public ExcelDocument()
        {

        }
    }
}
