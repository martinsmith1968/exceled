using DNX.Helpers.Strings;
using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using OfficeOpenXml;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.Edit
{
    public class SetFormulaCommand : BaseCommand<SetFormulaCommand.SetFormulaArguments>
    {
        public class SetFormulaArguments
        {
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string Formula { get; set; }

            [Alias("r")]
            [CommandLineArgument(IsRequired = false, DefaultValue = null)]
            public string Range { get; set; }

            [Alias("c")]
            [CommandLineArgument(IsRequired = false, DefaultValue = true)]
            public bool Calculate { get; set; }

            [Alias("h")]
            [CommandLineArgument(IsRequired = false, DefaultValue = false)]
            public bool Hidden { get; set; }
        }

        public override void Execute(IExcelDocument document, SetFormulaArguments arguments)
        {
            var range = document.GetActiveRange(arguments.Range);

            Logger.Information("{Reference}: {Formula}", range.Reference, arguments.Formula);
            document.ActiveWorksheet.Cells[range.Reference].Formula = arguments.Formula.RemoveStartsWith(ExecutionContext.CommandCommentPrefix);

            document.ActiveWorksheet.Cells[range.Reference].Style.Hidden = arguments.Hidden;

            if (arguments.Calculate)
            {
                document.ActiveWorksheet.Cells[range.Reference].Calculate();
            }
        }
    }
}
