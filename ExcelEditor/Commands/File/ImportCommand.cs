using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.File
{
    public class ImportCommand : BaseCommand<ImportCommand.ImportArguments>
    {
        public enum FileType
        {
            CSV
        }

        public class ImportArguments
        {
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string FileName { get; set; }

            [CommandLineArgument(IsRequired = false, DefaultValue = FileType.CSV)]
            public FileType Type { get; set; }

            [Alias("r")]
            [CommandLineArgument(IsRequired = false, DefaultValue = null)]
            public string Range { get; set; }
        }

        public override void Execute(IExcelDocument document, ImportArguments arguments)
        {
            // TODO
        }
    }
}
