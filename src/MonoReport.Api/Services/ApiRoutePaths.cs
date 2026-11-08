namespace MonoReport.Api.Services;

public static class ApiRoutePaths
{
    public const string DocumentAnalyze = "/api/documents/analyze";
    public const string DocumentSample = "/api/documents/sample";
    public const string DocumentRender = "/api/documents/render";
    
    public const string Formatters = "/api/formatters";
    public const string FormatterExecution = "/api/formatters/{name}/execute";
}
