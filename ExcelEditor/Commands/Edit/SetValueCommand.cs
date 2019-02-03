using System;
using ExcelEditor.Excel.Document;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.Edit
{
    public class SetValueCommand : BaseCommand<SetValueCommand.SetValueArguments>
    {
        public class SetValueArguments
        {
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string Text { get; set; }

            [Alias("r")]
            [CommandLineArgument(IsRequired = false, DefaultValue = null)]
            public string Range { get; set; }
        }

        public override void Execute(IExcelDocument document, SetValueArguments arguments)
        {
            var range = document.GetActiveRange(arguments.Range);

            var value = int.TryParse(arguments.Text, out var intValue) ? (object)intValue
                : long.TryParse(arguments.Text, out var longValue) ? (object)longValue
                : double.TryParse(arguments.Text, out var doubleValue) ? (object)doubleValue
                : DateTime.TryParse(arguments.Text, out var dateTimeValue) ? (object)dateTimeValue
                : arguments.Text;

            Logger.WriteText($"{range.Reference}: {value}");
            document.ActiveWorksheet.Cells[range.Reference].Value = value;
        }
    }
}
