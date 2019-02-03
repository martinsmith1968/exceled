using System;
using ExcelEditor.Excel.Document;
using ExcelEditor.Excel.Elements;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.Selection
{
    public class SelectCommand : BaseCommand<SelectCommand.SelectArguments>
    {
        public class SelectArguments
        {
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string Range { get; set; }
        }

        public override void Execute(IExcelDocument document, SelectArguments arguments)
        {
            var range = Range.Parse(arguments.Range)
                ?? throw new ArgumentOutOfRangeException(nameof(arguments.Range), $"Unable to parse Range: {arguments.Range}");

            document.Selection = range;
        }
    }
}
