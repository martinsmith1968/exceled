using System.IO;
using System.Linq;
using DNX.Helpers.Strings;
using ExcelEditor.Interfaces;
using ExcelEditor.Lib.Execution;

namespace ExcelEditor.Factories
{
    public class CommandReader : ICommandReader
    {
        private readonly IExecutionContext ExecutionContext;

        public CommandReader(IExecutionContext executionContext)
        {
            ExecutionContext = executionContext;
        }

        public string[] ParseCommands(string[] commands)
        {
            var validCommands = commands
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Where(c => !c.StartsWith(ExecutionContext.CommandCommentPrefix))
                .SelectMany(ExpandIncludeCommands)
                .ToArray();

            return validCommands;
        }

        private string[] ExpandIncludeCommands(string command)
        {
            // TODO - Need relative folders
            if (command.StartsWith(ExecutionContext.ScriptIncludePrefix))
            {
                var fileName = command.RemoveStartsWith(ExecutionContext.ScriptIncludePrefix);

                return ReadCommandScript(fileName);
            }

            return new[] { command };
        }

        private string[] ReadCommandScript(string fileName)
        {
            var commands = File.ReadAllLines(fileName);

            return ParseCommands(commands);
        }
    }
}
