using ExcelEditor.Excel.Document;
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
        }

        public override void Execute(IExcelDocument document, SetFormulaArguments arguments)
        {
            var range = document.GetActiveRange(arguments.Range);

            Logger.WriteText($"{range.Reference}: {arguments.Formula}");
            document.ActiveWorksheet.Cells[range.Reference].Formula = arguments.Formula;

            if (arguments.Calculate)
                document.ActiveWorksheet.Cells[range.Reference].Calculate();
        }
    }
}
