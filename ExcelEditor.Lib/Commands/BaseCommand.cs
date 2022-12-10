using System.Collections.Generic;
using Autofac;
using DNX.Helpers.Strings;
using ExcelEditor.Lib.Excel.Document;
using ExcelEditor.Lib.Execution;
using Serilog;

namespace ExcelEditor.Lib.Commands
{
    public abstract class BaseCommand : ICommand
    {
        private static readonly IDictionary<string, string> Variables = new Dictionary<string, string>();

        public virtual ILogger Logger { get; set; }

        public ILifetimeScope Container { get; set; }   // TODO: Internal set

        public IExecutionContext ExecutionContext { get; set; } // TODO: Internal set

        public virtual string Name => GetType().Name
            .RemoveEndsWith(
                typeof(ICommand).Name.RemoveStartsWith("I")
            );

        protected string GetVariable(string name, string commandName)
        {
            var variableName = string.IsNullOrEmpty(commandName)
                ? Name
                : $"{commandName}:{name}";

            return Variables[variableName];
        }

        protected void SetVariable(string name, string value, string commandName)
        {
            var variableName = string.IsNullOrEmpty(commandName)
                ? Name
                : $"{commandName}:{name}";

            Variables[variableName] = value;
        }

        public abstract void Execute(IExcelDocument document, string[] args);
    }
}
