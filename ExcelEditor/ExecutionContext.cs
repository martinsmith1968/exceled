using ExcelEditor.Lib.Execution;

namespace ExcelEditor
{
    public class ExecutionContext : IExecutionContext
    {
        public string CommandCommentPrefix { get; set; }
        public string ScriptIncludePrefix { get; set; }
        public string EnvironmentVariableExpansionPrefixSuffix { get; set; }
        public string FunctionEvaluationPrefix { get; set; }
        public string FunctionEvaluationSuffix { get; set; }
        public string ExcelFormulaPrefix { get; set; }

        public ExecutionContext(ApplicationArguments applicationArguments)
        {
            CommandCommentPrefix                     = applicationArguments.CommentPrefix.ToString();
            ScriptIncludePrefix                      = applicationArguments.ScriptIncludePrefix.ToString();
            EnvironmentVariableExpansionPrefixSuffix = applicationArguments.EnvironmentVariableExpansionPrefixSuffix.ToString();
            FunctionEvaluationPrefix                 = applicationArguments.FunctionEvaluationPrefix.ToString();
            FunctionEvaluationSuffix                 = applicationArguments.FunctionEvaluationSuffix.ToString();
            ExcelFormulaPrefix                       = applicationArguments.ExcelFormulaPrefix.ToString();
        }
    }
}
