using DNX.Helpers.Strings;
using ExcelEditor.Interfaces;

namespace ExcelEditor.Commands
{
    public abstract class BaseCommand : ICommand
    {
        public virtual string Name => GetType().Name.RemoveEndsWith(typeof(ICommand).Name.RemoveStartsWith("I"));

        public abstract void Execute();
    }
}
