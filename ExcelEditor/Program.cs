using System;
using Autofac;
using ExcelEditor.Extensions;
using ExcelEditor.Interfaces;
using ExcelEditor.Ioc;
using Ookii.CommandLine;

namespace ExcelEditor
{
    internal class Program
    {
        private static IContainer Container { get; set; }

        private static CommandLineParser Parser;

        private static int Main(string[] args)
        {
            try
            {
                Container = AutofacIocRegistrations.BuildContainer();

                var arguments = args.ParseArguments<ApplicationArguments>(out Parser);

                var application = Container.Resolve<IApplication>();
                application.Execute(arguments);

                return 0;
            }
            catch (CommandLineArgumentException e)
            {
                Console.Error.WriteLine($"{e.GetType().Name} Error: {e.Message}");

                Parser?.WriteUsageToConsole(new WriteUsageOptions()
                {
                    IncludeAliasInDescription = true
                });

                return 3;
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
