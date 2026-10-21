using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using MonoReport.Models;
using MonoReport.Models.Xlsx;
using MonoReport.Services.Xlsx;

namespace MonoReport.Tests.Services;

public sealed partial class Feature012BindingTests
{
    private readonly XlsxFixture _fixture = new();

    [Fact]
    public void Scopes_Keep_first_tuple_leaf_header_parent_and_root_values_separate()
    {
        using var services = _fixture.Services(); var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        using var data = JsonDocument.Parse("{\"Name\":\"root\",\"Items\":[{\"$level\":10,\"$path\":\"x\",\"Name\":\"first\"},{\"$level\":10,\"$path\":\"x\",\"Name\":\"second\"}],\"Bids\":[{\"Company\":\"parent\"}],\"Criteria\":[{\"Name\":\"child\"}]}");
        var row = planner.GroupRows(data.RootElement.GetProperty("Items"), "$.Items", true).Single();
        Assert.Equal("first", planner.RowScope(data.RootElement, row, "$.Items", null).Values[0].GetProperty("Name").GetString());
        var leaf = planner.RowScope(data.RootElement, row, "$.Items", 1);
        Assert.Equal("second", leaf.Values[0].GetProperty("Name").GetString()); Assert.Equal("$.Items[1]", leaf.Paths[0]);
        var header = planner.HeaderScope(data.RootElement,
            [new(1, "$.Bids", [data.RootElement.GetProperty("Bids")[0]]), new(2, "$.Criteria", [data.RootElement.GetProperty("Criteria")[0]])], [0, 0]);
        Assert.Equal("child", header.Values[0].GetProperty("Name").GetString());
        Assert.Equal("parent", header.Values[1].GetProperty("Company").GetString());
        Assert.Equal("root", header.Root.GetProperty("Name").GetString());
        Assert.Equal(new[] { "$.Criteria[0]", "$.Bids[0]", "$" }, header.Paths);
        Assert.Equal("root", leaf.Values[1].GetProperty("Name").GetString()); // no implicit header ancestor
    }

    [Theory]
    [InlineData(2, 3, 6)] [InlineData(2, 0, 1)] [InlineData(0, 3, 1)] [InlineData(1, 1, 1)]
    public void Arity_Requires_complete_product_or_one_descriptor(int outer, int inner, int arity)
    {
        using var services = _fixture.Services(); var planner = services.GetRequiredService<XlsxTemplatePlanner>();
        Assert.Equal(arity, planner.RecordArity([outer, inner]));
        foreach (var count in new[] { arity - 1, arity, arity + 1 }.Where(count => count > 0))
        {
            using var data = JsonDocument.Parse(JsonSerializer.Serialize(Enumerable.Range(0, count).Select(_ => new Dictionary<string, object> { ["$path"] = "row", ["$level"] = 10 })));
            var rows = planner.GroupRows(data.RootElement, "$.Items", true);
            if (count == arity) planner.ValidateArity(rows, arity, "$.Items");
            else Assert.Contains("exactly " + arity, Assert.Throws<TemplateException>(() => planner.ValidateArity(rows, arity, "$.Items")).Message);
        }
        using var empty = JsonDocument.Parse("[]");
        planner.ValidateArity(planner.GroupRows(empty.RootElement, "$.Items", true), arity, "$.Items");
    }

    [Fact]
    public void Arity_Uses_checked_product_arithmetic()
    {
        using var services = _fixture.Services();
        Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxTemplatePlanner>().RecordArity([int.MaxValue, int.MaxValue, 3]));
        Assert.Equal(1, services.GetRequiredService<XlsxTemplatePlanner>().RecordArity([int.MaxValue, int.MaxValue, 3, 0]));
    }

    [Fact]
    public void Groups_Preserve_input_order_first_indices_and_opaque_keys()
    {
        using var services = _fixture.Services();
        using var data = JsonDocument.Parse("[{\"$level\":20,\"$path\":\"z/child\"},{\"$level\":20,\"$path\":\"z/child\"},{\"$level\":10,\"$path\":\"a\"}]");
        var rows = services.GetRequiredService<XlsxTemplatePlanner>().GroupRows(data.RootElement, "$.Items", true);
        Assert.Equal(new[] { "z/child", "a" }, rows.Select(row => row.Key));
        Assert.Equal(new[] { 0, 2 }, rows.Select(row => row.FirstIndex));
        Assert.Equal(new[] { 20, 10 }, rows.Select(row => row.Level));
        Assert.Equal(new[] { 2, 1 }, rows.Select(row => row.Records.Length));
        using var empty = JsonDocument.Parse("[]");
        Assert.Empty(services.GetRequiredService<XlsxTemplatePlanner>().GroupRows(empty.RootElement, "$.Items", false));
    }

    [Theory]
    [InlineData("[{\"$level\":1,\"$path\":\"a\"},{\"$level\":1,\"$path\":\"b\"},{\"$level\":1,\"$path\":\"a\"}]")]
    [InlineData("[{\"$level\":1,\"$path\":\"a\"},{\"$level\":2,\"$path\":\"a\"}]")]
    [InlineData("[{\"$level\":1.5,\"$path\":\"a\"}]")]
    [InlineData("[{\"$level\":\"1\",\"$path\":\"a\"}]")]
    [InlineData("[{\"$level\":1}]")]
    public void Groups_Reject_malformed_or_noncontiguous_records(string json)
    {
        using var services = _fixture.Services(); using var data = JsonDocument.Parse(json);
        var error = Assert.Throws<TemplateException>(() => services.GetRequiredService<XlsxTemplatePlanner>().GroupRows(data.RootElement, "$.Items", true));
        Assert.Equal(DiagnosticCodes.GenerationInvalidData, error.Code); Assert.Contains("$.Items[", error.Location);
    }
}
