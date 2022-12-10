using System.IO;
using System.Linq;
using ExcelEditor.Interfaces;
using ExcelEditor.Lib.Arguments;
using Ookii.CommandLine;

namespace ExcelEditor
{
    public class ApplicationArguments : IValidatableArguments, IHelpArguments
    {
        [CommandLineArgument(Position = 0, IsRequired = true)]
        public string OutputFileName { get; set; }

        [Alias("e")]
        [CommandLineArgument(IsRequired = false)]
        public string[] Commands { get; set; }

        [Alias("c")]
        [CommandLineArgument(IsRequired = false)]
        public string CommandScriptFileName { get; set; }

        [Alias("af")]
        [CommandLineArgument(IsRequired = false)]
        public string[] AssemblyFolders { get; set; }

        [Alias("cp")]
        [CommandLineArgument(IsRequired = false, DefaultValue = Constants.DefaultCommandCommentPrefix)]
        public char CommentPrefix { get; set; }

        [Alias("sip")]
        [CommandLineArgument(IsRequired = false, DefaultValue = Constants.DefaultScriptIncludePrefix)]
        public char ScriptIncludePrefix { get; set; }

        [Alias("fep")]
        [CommandLineArgument(IsRequired = false, DefaultValue = Constants.DefaultFunctionEvaluationPrefix)]
        public char FunctionEvaluationPrefix { get; set; }

        [Alias("fes")]
        [CommandLineArgument(IsRequired = false, DefaultValue = Constants.DefaultFunctionEvaluationSuffix)]
        public char FunctionEvaluationSuffix { get; set; }

        [Alias("eveps")]
        [CommandLineArgument(IsRequired = false, DefaultValue = Constants.DefaultEnvironmentVariableExpansionPrefixSuffix)]
        public char EnvironmentVariableExpansionPrefixSuffix { get; set; }

        [Alias("exfp")]
        [CommandLineArgument(IsRequired = false, DefaultValue = Constants.Excel.DefaultFormulaPrefix)]
        public char ExcelFormulaPrefix { get; set; }

        [Alias("?")]
        [CommandLineArgument(IsRequired = false, DefaultValue = false)]
        public bool Help { get; set; }

        public string[] CommandNames => Commands
                                         ?? File.ReadAllLines(CommandScriptFileName);

        public void Validate()
        {
            if (!Commands.Any() && string.IsNullOrWhiteSpace(CommandScriptFileName))
                throw new CommandLineArgumentException(
                    $"Must specify {nameof(Commands)} or {nameof(CommandScriptFileName)}",
                    CommandLineArgumentErrorCategory.MissingRequiredArgument
                    );

            if (!string.IsNullOrWhiteSpace(CommandScriptFileName) && !File.Exists(CommandScriptFileName))
                throw new FileNotFoundException(
                    $"File not found : {CommandScriptFileName}",
                    CommandScriptFileName
                    );

            if (!CommandNames.Any())
            {
                throw new CommandLineArgumentException(
                    $"",
                    CommandLineArgumentErrorCategory.MissingNamedArgumentValue);
            }
        }
    }
}
