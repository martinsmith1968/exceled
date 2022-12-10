using System;
using ExcelEditor.Lib.Excel.Document;
using ExcelEditor.Lib.Extensions;
using Ookii.CommandLine;

namespace ExcelEditor.Lib.Commands
{
    public interface IBaseParseableCommand
    {
        CommandLineParser Parser { get; }
    }

    public abstract class BaseCommand<T> : BaseCommand, IBaseParseableCommand
    {
        private CommandLineParser _parser;
        public CommandLineParser Parser => _parser;

        public override void Execute(IExcelDocument document, string[] args)
        {
            var arguments = args.ParseArguments<T>(out _parser);
            if (!_parser.IsValid(arguments))
                throw new ArgumentException($"{nameof(Parser)} or {nameof(arguments)} invalid");

            Logger.Debug("Executing: {@Arguments}", arguments);

            Execute(document, arguments);
        }

        public abstract void Execute(IExcelDocument document, T arguments);
    }
}
