using System;
using System.Collections.Generic;
using System.Linq;
using Autofac;
using ExcelEditor.Commands;
using ExcelEditor.Excel.Document;
using ExcelEditor.Interfaces;
using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Excel.Document;
using ExcelEditor.Lib.Extensions;
using Ookii.CommandLine;
using Serilog;
using Serilog.Context;

namespace ExcelEditor
{
    public class Application : IApplication
    {
        private readonly ILogger _logger = Log.ForContext<Application>();
        private readonly ICommandReader _commandReader;
        private readonly ICommandFactory _commandFactory;
        private readonly ApplicationArguments _arguments;

        public Application(
            ICommandReader commandReader,
            ICommandFactory commandFactory,
            ApplicationArguments arguments
            )
        {
            _commandReader  = commandReader;
            _commandFactory = commandFactory;
            _arguments      = arguments;
        }

        public void Execute(IContainer container)
        {
            var document = new ExcelDocument(_arguments.OutputFileName, null);

            var commands = _commandReader.ParseCommands(_arguments.Commands);

            ProcessCommands(document, commands);
        }

        private void ProcessCommands(IExcelDocument document, IEnumerable<string> commands)
        {
            foreach (var commandText in commands)
            {
                var commandParts = ParserExtensions.SplitCommandLineArguments(commandText);

                var commandName = commandParts.FirstOrDefault() ?? string.Empty;
                var commandArgs = commandParts.Skip(1).ToArray();

                _logger.Debug("Parsing: {CommandName} {commandArguments}", commandName, commandArgs);

                var command = _commandFactory.GetCommandByName(commandName);
                if (command == null)
                    throw new ArgumentOutOfRangeException(nameof(commandName), $"Invalid or unsupported {nameof(ICommand)} : {commandName}");


                var success = ExecuteCommand(document, command, commandArgs);

                // TODO: Decide how to proceed after command failure
            }
        }

        private static bool ExecuteCommand(IExcelDocument document, ICommand command, string[] commandArgs)
        {
            using (LogContext.PushProperty("CommandName", command.Name))
            {
                try
                {
                    command.Execute(document, commandArgs);

                    return true;
                }
                catch (CommandLineArgumentException e)
                {
                    if (command is IBaseParseableCommand parseableCommand)
                    {
                        parseableCommand.Parser.ShowHelpToConsole();
                    }
                    else
                    {
                        Console.Error.WriteLine($"{e.GetType().Name} Error: {e.Message}");
                    }
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine($"{e.GetType().Name} Error: {e.Message}{Environment.NewLine}{e}");
                }

                return false;
            }
        }
    }
}
