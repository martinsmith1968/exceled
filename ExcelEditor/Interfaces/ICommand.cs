namespace ExcelEditor.Interfaces
{
    public interface ICommand
    {
        string Name { get; }

        void Execute();
    }
}
