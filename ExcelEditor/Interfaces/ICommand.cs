using ExcelEditor.Excel;
using ExcelEditor.Excel.Document;
using OfficeOpenXml;

namespace ExcelEditor.Interfaces
{
    public interface ICommand
    {
        string Name { get; }

        void Execute(IExcelDocument excelManager, string[] args);
    }
}
