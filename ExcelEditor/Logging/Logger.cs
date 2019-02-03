using System;
using ExcelEditor.Interfaces;

namespace ExcelEditor.Logging
{
    public class Logger : ILogger
    {
        public string OwnerName { get; private set; }

        public Logger(ICommand owner)
        {
            OwnerName = owner?.Name ?? "Unknown";
        }

        public void WriteText(string text)
        {
            Console.WriteLine($"{OwnerName}: {text}");
        }

        public void WriteText(string format, params object[] args)
        {
            WriteText(string.Format(format, args));
        }
    }
}
