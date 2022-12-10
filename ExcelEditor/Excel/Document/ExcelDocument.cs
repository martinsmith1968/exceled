using System.IO;
using ExcelEditor.Excel.Elements;
using ExcelEditor.Lib.Excel.Document;
using ExcelEditor.Lib.Excel.Elements;
using OfficeOpenXml;

namespace ExcelEditor.Excel.Document
{
    public class ExcelDocument  : IExcelDocument
    {
        public string FileName { get; set; }

        public ExcelPackage ExcelPackage { get; private set; }

        public ExcelWorksheet ActiveWorksheet { get; set; }

        public IRange Selection { get; set; }

        public IRange GetActiveRange(string addressText)
        {
            var range = Range.Parse(addressText) ?? Selection;

            return range;
        }

        public IRange GetSetActiveRange(string addressText)
        {
            Selection = GetActiveRange(addressText);

            return Selection;
        }

        public ExcelDocument(string fileName, string templateFileName)
        {
            FileName = fileName;
            ExcelPackage = CreateExcelPackage(fileName, templateFileName);
        }

        private static ExcelPackage CreateExcelPackage(string fileName, string templateFileName)
        {
            var fileInfo = new FileInfo(fileName);
            var templateFileInfo = string.IsNullOrEmpty(templateFileName)
                ? null
                : new FileInfo(templateFileName);

            return templateFileInfo == null
                ? new ExcelPackage(fileInfo)
                : new ExcelPackage(fileInfo, templateFileInfo);
        }
    }
}
