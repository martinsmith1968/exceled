using System.IO;
using ExcelEditor.Excel.Document;

namespace ExcelEditor.Commands.File
{
    public class SaveCommand : BaseCommand<SaveCommand.SaveArguments>
    {
        public class SaveArguments
        {
        }

        public override void Execute(IExcelDocument document, SaveArguments arguments)
        {
            var fileInfo = new FileInfo(document.FileName);

            document.ExcelPackage.SaveAs(fileInfo);
        }
    }
}
