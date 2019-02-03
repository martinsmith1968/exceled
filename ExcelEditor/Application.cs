using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelEditor.Excel.Document;
using ExcelEditor.Interfaces;
using OfficeOpenXml;

namespace ExcelEditor
{
    public class Application : IApplication
    {
        private readonly ICommandFactory _commandFactory;

        public Application(ICommandFactory commandFactory)
        {
            _commandFactory = commandFactory;
        }


        public void Execute(ApplicationArguments arguments)
        {
            var excelPackage = CreateExcelPackage(arguments.FileName, null);

            var document = new ExcelDocument()
            {
                FileName     = arguments.FileName,
                ExcelPackage = excelPackage
            };

            ProcessCommands(document, arguments.Commands);
        }

        private void ProcessCommands(IExcelDocument document, IEnumerable<string> commands)
        {
            foreach (var commandText in commands)
            {
                var commandParts = commandText.Split(" ".ToCharArray());
                var commandName = commandParts.FirstOrDefault() ?? string.Empty;
                var commandArgs = commandParts.Skip(1).ToArray();

                if (commandName.StartsWith("@"))
                {
                    var fileName = commandName.TrimStart("@".ToCharArray());

                    var fileCommands = File.ReadAllLines(fileName)
                        .Where(x => !string.IsNullOrEmpty(x))
                        .Where(x => !x.StartsWith("#"))
                        .ToArray();

                    ProcessCommands(document, fileCommands);

                    continue;
                }

                var command = _commandFactory.GetCommandByName(commandName);
                if (command == null)
                    throw new ArgumentOutOfRangeException(nameof(commandName), $"Unknown command: {commandName}");

                command.Execute(document, commandArgs);
            }
        }

        private ExcelPackage CreateExcelPackage(string fileName, string templateFileName)
        {
            // TODO
            return new ExcelPackage();
        }
    }
}
