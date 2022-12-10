using ExcelEditor.Interfaces;
using ExcelEditor.Lib.Arguments;
using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using Ookii.CommandLine;

namespace ExcelEditor.Commands.Edit
{
    public class SetStyleCommand : BaseCommand<SetStyleCommand.SetStyleArguments>
    {
        public class SetStyleArguments : IValidatableArguments
        {
            [CommandLineArgument(IsRequired = false)]
            public string StyleName { get; set; }

            [Alias("h")]
            [CommandLineArgument(IsRequired = false)]
            public bool? Hidden { get; set; }

            [Alias("l")]
            [CommandLineArgument(IsRequired = false)]
            public bool? Locked { get; set; }

            [Alias("w")]
            [CommandLineArgument(IsRequired = false)]
            public bool? WrapText { get; set; }

            [Alias("r")]
            [CommandLineArgument(IsRequired = false, DefaultValue = null)]
            public string Range { get; set; }

            public void Validate()
            {

            }
        }

        public override void Execute(IExcelDocument document, SetStyleArguments arguments)
        {
            var range = document.GetActiveRange(arguments.Range);

            if (!string.IsNullOrEmpty(arguments.StyleName))
            {
                Logger.Information("{Reference}: StyleName: {StyleName}", range.Reference, arguments.StyleName);
                document.ActiveWorksheet.Cells[range.Reference].StyleName = arguments.StyleName;
            }

            if (arguments.Hidden.HasValue)
            {
                Logger.Information("{Reference}: Hidden: {Hidden}", range.Reference, arguments.Hidden.Value);
                document.ActiveWorksheet.Cells[range.Reference].Style.Hidden = arguments.Hidden.Value;
            }

            if (arguments.Locked.HasValue)
            {
                Logger.Information("{Reference}: Locked: {Locked}", range.Reference, arguments.Locked.Value);
                document.ActiveWorksheet.Cells[range.Reference].Style.Locked = arguments.Locked.Value;
            }

            if (arguments.WrapText.HasValue)
            {
                Logger.Information("{Reference}: WrapText: {WrapText}", range.Reference, arguments.WrapText.Value);
                document.ActiveWorksheet.Cells[range.Reference].Style.WrapText = arguments.WrapText.Value;
            }
        }
    }
}
