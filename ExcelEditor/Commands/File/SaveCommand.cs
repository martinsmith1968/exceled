using System.IO;
using ExcelEditor.Excel.Document;
using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.File
{
    public class SaveCommand : BaseCommand<SaveCommand.SaveArguments>
    {
        public class SaveArguments
        {
            [CommandLineArgument(IsRequired = false, DefaultValue = true)]
            public bool Overwrite { get; set; }
        }

        public override void Execute(IExcelDocument document, SaveArguments arguments)
        {
            var fileInfo = new FileInfo(document.FileName);
            if (fileInfo.Exists && !arguments.Overwrite)
                throw new IOException($"Unable to overwrite file: {fileInfo.Name}");

            document.ExcelPackage.SaveAs(fileInfo);
        }
    }
}
