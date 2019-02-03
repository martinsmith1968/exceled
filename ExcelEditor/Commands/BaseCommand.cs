using DNX.Helpers.Strings;
using ExcelEditor.Excel.Document;
using ExcelEditor.Interfaces;
using ExcelEditor.Logging;

namespace ExcelEditor.Commands
{
    public abstract class BaseCommand : ICommand
    {
        private ILogger _logger;

        protected ILogger Shit => _logger = _logger ?? new Logger(this);

        protected ILogger Logger
        {
            get
            {
                _logger = _logger ?? new Logger(this);

                return _logger;
            }
        }



        public virtual string Name => GetType().Name.RemoveEndsWith(typeof(ICommand).Name.RemoveStartsWith("I"));

        public abstract void Execute(IExcelDocument document, string[] args);
    }
}
