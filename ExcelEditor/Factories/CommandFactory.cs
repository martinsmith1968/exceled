using System;
using System.Collections.Generic;
using System.Linq;
using ExcelEditor.Interfaces;
using ExcelEditor.Lib.Commands;
using Serilog;

// ReSharper disable ParameterTypeCanBeEnumerable.Local

namespace ExcelEditor.Factories
{
    public class CommandFactory : ICommandFactory
    {
        private readonly ILogger _logger = Log.ForContext<CommandFactory>();

        public IEnumerable<ICommand> Commands { get; }

        public CommandFactory(
            IList<ICommand> commands
            )
        {
            Commands = commands;
            _logger.Debug("Discovered {CommandCount} Commands", Commands.Count());

            foreach (var command in Commands)
            {
                _logger.Debug("Discovered Command: {CommandName} - {CommandFullName} {AssemblyLocation}",
                    command.Name,
                    command.GetType().FullName,
                    command.GetType().Assembly.Location
                    );
            }
        }

        public ICommand GetCommandByName(string commandName)
        {
             _logger.Debug("Locating Command: {CommandName}", commandName);

            var command = Commands
                    .SingleOrDefault(c => string.Equals(c.Name, commandName, StringComparison.OrdinalIgnoreCase))
                ;

            return command;
        }
    }
}
