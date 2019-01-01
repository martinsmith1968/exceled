using System;
using System.Collections.Generic;
using System.Linq;
using ExcelEditor.Interfaces;

namespace ExcelEditor.Factories
{
    public class CommandFactory : ICommandFactory
    {
        public IEnumerable<ICommand> Commands { get; }

        public CommandFactory(IEnumerable<ICommand> commands)
        {
            Commands = commands;
        }

        public ICommand GetCommandByName(string commandName)
        {
            var command = Commands
                    .SingleOrDefault(c => string.Equals(c.Name, commandName, StringComparison.OrdinalIgnoreCase))
                ;
            if (command == null)
                throw new ArgumentOutOfRangeException(nameof(commandName), $"Invalid or unsupported {nameof(ICommand)} : {commandName}");

            return command;
        }
    }
}
