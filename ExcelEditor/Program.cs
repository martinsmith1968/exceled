using System;
using Autofac;
using ExcelEditor.Interfaces;
using ExcelEditor.Ioc;
using ExcelEditor.Lib.Extensions;
using ExcelEditor.Logging.Enrichers;
using Microsoft.Extensions.Configuration;
using Ookii.CommandLine;
using Serilog;

// ReSharper disable InconsistentNaming

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
                var configuration = new ConfigurationBuilder()
                    .AddJsonFile("appSettings.json", true)
                    .Build();

                #if DEBUG
                Serilog.Debugging.SelfLog.Enable(Console.Error);
                #endif

                Log.Logger = new LoggerConfiguration()
                    .Enrich.FromLogContext()
                    .Enrich.With<AppInfoEnricher>()
                    .ReadFrom.Configuration(configuration)
                    .CreateLogger();

                Log.Debug(new string('-', 100));
                Log.Information("{Application} Starting");

                var arguments = args.ParseArguments<ApplicationArguments>(out Parser);

                Container = AutofacIocRegistrations.BuildContainer(arguments.AssemblyFolders, arguments);

                var application = Container.Resolve<IApplication>();
                application.Execute(Container);

                return 0;
            }
            catch (CommandLineArgumentException e)
            {
                Console.Error.WriteLine($"{e.GetType().Name} Error: {e.Message}");

                Parser.ShowHelpToConsole();

                return 3;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"{e.GetType().Name} Error: {e.Message}{Environment.NewLine}{e}");

                return 4;
            }
            #if DEBUG
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
            #endif
        }
    }
}
