using System;
using System.ComponentModel;
using System.Linq;
using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using Ookii.CommandLine;
using Range = ExcelEditor.Excel.Elements.Range;

namespace ExcelEditor.Commands.Worksheet
{
    [Description("Activate a Worksheet")]
    public class ActivateCommand : BaseCommand<ActivateCommand.ActivateArguments>
    {
        public class ActivateArguments
        {
            [CommandLineArgument(Position = 0, IsRequired = true)]
            public string WorksheetName { get; set; }

            [Alias("c")]
            [CommandLineArgument(IsRequired = false, DefaultValue = true)]
            public bool CreateIfNotExists { get; set; }
        }

        public override void Execute(IExcelDocument document, ActivateArguments arguments)
        {
            var selectedWorksheet = document.ExcelPackage.Workbook.Worksheets.FirstOrDefault(w => w.Name.Equals(arguments.WorksheetName));
            if (selectedWorksheet == null)
            {
                if (arguments.CreateIfNotExists)
                {
                    selectedWorksheet = document.ExcelPackage.Workbook.Worksheets.Add(arguments.WorksheetName);
                }
            }

            if (selectedWorksheet == null)
                throw new ArgumentOutOfRangeException($"Invalid Worksheet Name : {arguments.WorksheetName}");

            document.ActiveWorksheet = selectedWorksheet;
            document.Selection = Range.Parse(Range.DefaultAddress);
        }
    }
}
