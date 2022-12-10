using Autofac;

namespace ExcelEditor.Interfaces
{
    public interface IApplication
    {
        void Execute(IContainer container);
    }
}
