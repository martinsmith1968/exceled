using ExcelEditor.Excel;
using ExcelEditor.Excel.Document;
using ExcelEditor.Extensions;
using Ookii.CommandLine;

namespace ExcelEditor.Commands
{
    public abstract class BaseCommand<T> : BaseCommand
    {
        private CommandLineParser Parser;

        public override void Execute(IExcelDocument document, string[] args)
        {
            var arguments = args.ParseArguments<T>(out Parser);

            Execute(document, arguments);
        }

        public abstract void Execute(IExcelDocument document, T arguments);
    }
}
