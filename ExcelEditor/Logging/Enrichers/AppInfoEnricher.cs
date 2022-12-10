using System.Collections.Generic;
using DNX.Helpers.Assemblies;
using Serilog.Core;
using Serilog.Events;

namespace ExcelEditor.Logging.Enrichers
{
    public class AppInfoEnricher : ILogEventEnricher
    {
        private readonly IAssemblyDetails _assembly;

        public AppInfoEnricher()
        {
            _assembly = AssemblyDetails.ForEntryPoint();
        }

        public void Enrich(LogEvent logEvent, ILogEventPropertyFactory propertyFactory)
        {
            var dict = new Dictionary<string, object>()
            {
                { nameof(_assembly.Name), _assembly.Name },
                { nameof(_assembly.Version), _assembly.Version.Simplify() },
                { nameof(_assembly.FileName), _assembly.FileName },
                { nameof(_assembly.Location), _assembly.Location },
            };

            foreach (var (key, value) in dict)
            {
                logEvent.AddPropertyIfAbsent(propertyFactory.CreateProperty(key, value));
            }
        }
    }
}
