using System.IO;
using ExcelEditor.Excel.Document;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.File
{
    public class SaveAsCommand : BaseCommand<SaveAsCommand.SaveAsArguments>
    {
        public class SaveAsArguments
        {
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string FileName { get; set; }
        }

        public override void Execute(IExcelDocument document, SaveAsArguments arguments)
        {
            var fileInfo = new FileInfo(arguments.FileName);

            document.ExcelPackage.SaveAs(fileInfo);
        }
    }
}
