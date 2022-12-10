using System.Linq;
using System.Text.RegularExpressions;
using ExcelEditor.Lib.Arguments;
using Ookii.CommandLine;

namespace ExcelEditor.Lib.Extensions
{
    public static class ParserExtensions
    {
        public static string[] SplitCommandLineArguments(string text)
        {
            var re = new Regex("(?<=\")[^\"]*(?=\")|[^\" ]+");

            var strings = re.Matches(text)
                .Cast<Match>()
                .Select(m => m.Value)
                .ToArray();

            return strings;
        }

        public static T ParseArguments<T>(this string[] args, out CommandLineParser parser)
        {
            parser = new CommandLineParser(typeof(T));

            var arguments = (T)parser.Parse(args);

            var helpRequested = (arguments is IHelpArguments helpArguments) && helpArguments.Help;

            if (arguments is IValidatableArguments validatableArguments)
            {
                if (!helpRequested)
                {
                    validatableArguments.Validate();
                }
            }

            return arguments;
        }

        public static bool IsValid(this CommandLineParser parser)
        {
            return parser != null;
        }

        public static bool IsValid<T>(this CommandLineParser parser, T arguments)
        {
            return parser.IsValid()
                   && arguments != null;
        }

        public static void ShowHelpToConsole(this CommandLineParser parser, WriteUsageOptions options = null)
        {
            options ??= new WriteUsageOptions()
            {
                IncludeAliasInDescription = true,
                IncludeApplicationDescription = true,
                IncludeDefaultValueInDescription = true,
            };

            parser.WriteUsageToConsole(options);
        }
    }
}
