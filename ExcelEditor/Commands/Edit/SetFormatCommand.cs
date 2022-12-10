using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.Edit
{
    public class SetFormatCommand : BaseCommand<SetFormatCommand.SetFormulaArguments>
    {
        public class SetFormulaArguments
        {
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string Format { get; set; }

            [Alias("r")]
            [CommandLineArgument(IsRequired = false, DefaultValue = null)]
            public string Range { get; set; }
        }

        public override void Execute(IExcelDocument document, SetFormulaArguments arguments)
        {
            var range = document.GetActiveRange(arguments.Range);

            Logger.Information("{Reference}: Format: {Format}", range.Reference, arguments.Format);
            document.ActiveWorksheet.Cells[range.Reference].Style.Numberformat.Format = arguments.Format;
        }
    }
}
