namespace ExcelEditor.Logging
{
    public interface ILogger
    {
        void WriteText(string text);

        void WriteText(string format, params object[] args);
    }
}
