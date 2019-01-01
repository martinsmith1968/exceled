using System;
using Autofac;
using ExcelEditor.Extensions;
using ExcelEditor.Interfaces;
using ExcelEditor.Ioc;

namespace ExcelEditor
{
    internal class Program
    {
        private static IContainer Container { get; set; }

        private static int Main(string[] args)
        {
            try
            {
                Container = AutofacIocRegistrations.BuildContainer();

                var arguments = args.ParseArguments<ApplicationArguments>(out var parser);

                var application = Container.Resolve<IApplication>();
                application.Execute(arguments);

                return 0;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"{e.GetType().Name} Error: {e.Message}");

                return 4;
            }
            finally
            {
                if (System.Diagnostics.Debugger.IsAttached)
                {
                    while (Console.KeyAvailable)
                        Console.ReadKey(true);

                    Console.Write("Press any key to exit...");
                    Console.ReadKey(true);
                    Console.WriteLine();
                }
            }

        }
    }
}
