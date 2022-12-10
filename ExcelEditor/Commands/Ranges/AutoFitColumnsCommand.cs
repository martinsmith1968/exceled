using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.Ranges
{
    public class AutoFitColumnsCommand : BaseCommand<AutoFitColumnsCommand.AutoFitColumnsArguments>
    {
        public class AutoFitColumnsArguments
        {
            [Alias("r")]
            [CommandLineArgument(IsRequired = false, DefaultValue = null)]
            public string Range { get; set; }
        }

        public override void Execute(IExcelDocument document, AutoFitColumnsArguments arguments)
        {
            var range = document.GetActiveRange(arguments.Range);

            Logger.Information("{Reference}: AutoFitColumns", range.Reference);
            document.ActiveWorksheet.Cells[range.Reference].AutoFitColumns();
        }
    }
}
