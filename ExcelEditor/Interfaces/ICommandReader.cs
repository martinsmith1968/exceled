namespace ExcelEditor.Interfaces
{
    public interface ICommandReader
    {
        string[] ParseCommands(string[] commands);
    }
}
