using System;

namespace MonoReport.Api.Services;

public sealed class RequestInputException(int status, string code, string? field = null) : Exception
{
    public int Status { get; } = status;
    public string Code { get; } = code;
    public string? Field { get; } = field;
}