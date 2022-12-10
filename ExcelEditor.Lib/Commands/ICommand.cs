using Autofac;
using ExcelEditor.Lib.Excel.Document;
using ExcelEditor.Lib.Execution;
using Serilog;

namespace ExcelEditor.Lib.Commands
{
    public interface ICommand
    {
        ILogger Logger { get; }

        string Name { get; }

        ILifetimeScope Container { get; }

        IExecutionContext ExecutionContext { get; }

        void Execute(IExcelDocument excelManager, string[] args);
    }
}
