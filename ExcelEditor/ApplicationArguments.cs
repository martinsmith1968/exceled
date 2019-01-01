using System.Collections.Generic;
using ExcelEditor.Interfaces;
using Ookii.CommandLine;

namespace ExcelEditor
{
    public class ApplicationArguments : IValidatableArguments, IHelpArguments
    {
        [Alias("e")]
        [CommandLineArgument(IsRequired = false)]
        public IEnumerable<string> Commands { get; set; }

        [Alias("?")]
        [CommandLineArgument(IsRequired = false, DefaultValue = false)]
        public bool Help { get; }

        public void Validate()
        {
        }
    }
}
