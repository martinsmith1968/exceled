using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autofac;
using AutofacSerilogIntegration;
using DNX.Helpers.Comparisons;
using ExcelEditor.Factories;
using ExcelEditor.Interfaces;
using ExcelEditor.Lib.Commands;
using ExcelEditor.Lib.Execution;

namespace ExcelEditor.Ioc
{
    public static class AutofacIocRegistrations
    {
        private static readonly Type[] DiscoverableTypes =
        {
            typeof(ICommand)
        };

        public static IContainer BuildContainer(string[] assemblyFolders, params object[] instances)
        {
            var builder = new ContainerBuilder()
                    .RegisterInternals()
                    .RegisterImplementableTypes(assemblyFolders)
                ;

            builder.RegisterLogger();

            instances?.ToList()
                .ForEach(i => builder.RegisterInstance(i)
                    .AsSelf()
                );

            var container = builder.Build();

            return container;
        }

        private static ContainerBuilder RegisterInternals(this ContainerBuilder builder)
        {
            builder.RegisterType<Application>().As<IApplication>().SingleInstance();
            builder.RegisterType<ExecutionContext>().As<IExecutionContext>().SingleInstance();
            builder.RegisterType<CommandFactory>().As<ICommandFactory>().SingleInstance();
            builder.RegisterType<CommandReader>().As<ICommandReader>().SingleInstance();

            return builder;
        }

        private static bool AssemblyComparisonFunc(Assembly x, Assembly y)
        {
            return string.Equals(
                x?.Location ?? Guid.NewGuid().ToString(),
                y?.Location ?? Guid.NewGuid().ToString(),
                StringComparison.CurrentCultureIgnoreCase
            );
        }

        private static ContainerBuilder RegisterImplementableTypes(this ContainerBuilder builder, IList<string> assemblyFolders)
        {
            var assemblyFolderLocations = BuildAssemblyFolders(assemblyFolders);

            var assemblies = BuildAssemblyList(assemblyFolderLocations)
                .Distinct(EqualityComparerFunc<Assembly>.Create(AssemblyComparisonFunc))
                .ToArray();

            builder.RegisterAssemblyTypes(assemblies)
                .Where(t => DiscoverableTypes.Any(dt => dt.IsAssignableFrom(t)))
                .SingleInstance()
                .PropertiesAutowired()
                .AsImplementedInterfaces();

            return builder;
        }

        private static IEnumerable<string> BuildAssemblyFolders(IList<string> assemblyFolders)
        {
            var locations = new List<string>()
            {
                Path.GetDirectoryName(Assembly.GetEntryAssembly()?.Location),
                Directory.GetCurrentDirectory()
            };

            if (assemblyFolders?.Any() ?? false)
            {
                locations.AddRange(
                    assemblyFolders.Where(af => new DirectoryInfo(af).Exists)
                );
            }

            return locations;
        }

        private static IEnumerable<Assembly> BuildAssemblyList(IEnumerable<string> assemblyLocations)
        {
            var externalAssemblyFileNames = assemblyLocations
                .Where(fl => !string.IsNullOrWhiteSpace(fl))
                .SelectMany(fl =>
                    Directory.GetFiles(fl, "*.dll")
                );

            var baseAssemblies = new[]
                {
                    Assembly.GetEntryAssembly(),
                    Assembly.GetExecutingAssembly(),
                    Assembly.GetCallingAssembly(),
                }
                .Distinct()
                .ToArray();

            var baseReferencedAssemblies = baseAssemblies
                .SelectMany(a => a.GetReferencedAssemblies())
                .Distinct()
                .Select(Assembly.Load)
                .ToArray();

            var externalAssemblies = externalAssemblyFileNames
                .Select(Assembly.LoadFile)
                .ToArray();

            var assemblies = baseReferencedAssemblies
                .Union(baseAssemblies)
                .Union(externalAssemblies)
                .Distinct()
                .ToArray();

            return assemblies;
        }
    }
}
