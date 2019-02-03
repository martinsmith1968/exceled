using System.Collections.Generic;
using ExcelEditor.Interfaces;
using Ookii.CommandLine;

namespace ExcelEditor
{
    public class ApplicationArguments : IValidatableArguments, IHelpArguments
    {
        [CommandLineArgument(Position = 0, IsRequired = true)]
        public string FileName { get; set; }

        [Alias("e")]
        [CommandLineArgument(IsRequired = false)]
        public string[] Commands { get; set; }

        [Alias("?")]
        [CommandLineArgument(IsRequired = false, DefaultValue = false)]
        public bool Help { get; set; }

        public void Validate()
        {
        }
    }
}
