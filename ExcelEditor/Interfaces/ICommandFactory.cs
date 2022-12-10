using ExcelEditor.Lib.Commands;

namespace ExcelEditor.Interfaces
{
    public interface ICommandFactory
    {
        ICommand GetCommandByName(string commandName);
    }
}
