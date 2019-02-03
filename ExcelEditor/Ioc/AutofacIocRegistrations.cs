using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autofac;
using ExcelEditor.Factories;
using ExcelEditor.Interfaces;

namespace ExcelEditor.Ioc
{
    public static class AutofacIocRegistrations
    {
        public static IContainer BuildContainer(params object[] instances)
        {
            var builder = new ContainerBuilder()
                    .RegisterInternals()
                    .RegisterImplementableTypes()
                ;

            instances?.ToList()
                .ForEach(i => builder.RegisterInstance(i)
                    .AsSelf()
                );

            var container = builder.Build();

            return container;
        }

        private static IList<Assembly> BuildAssemblyList()
        {
            var fileLocation = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            var externalAssemblyFileNames = Directory.GetFiles(fileLocation, "*.dll");

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

        private static ContainerBuilder RegisterInternals(this ContainerBuilder builder)
        {
            builder.RegisterType<Application>().As<IApplication>().SingleInstance();
            builder.RegisterType<CommandFactory>().As<ICommandFactory>().SingleInstance();

            return builder;
        }

        private static ContainerBuilder RegisterImplementableTypes(this ContainerBuilder builder)
        {
            var assemblies = BuildAssemblyList()
                .ToArray();

            builder.RegisterAssemblyTypes(assemblies)
                .Where(t => typeof(ICommand).IsAssignableFrom(t))
                .SingleInstance()
                .AsImplementedInterfaces();

            return builder;
        }
    }
}
