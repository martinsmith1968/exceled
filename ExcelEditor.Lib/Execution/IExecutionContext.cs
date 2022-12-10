namespace ExcelEditor.Lib.Execution
{
    public interface IExecutionContext
    {
        string CommandCommentPrefix { get; }
        string ScriptIncludePrefix { get; }
        string EnvironmentVariableExpansionPrefixSuffix { get; }
        string FunctionEvaluationPrefix { get; }
        string FunctionEvaluationSuffix { get; }

        string ExcelFormulaPrefix { get; }
    }
}
