using System.Collections.Immutable;
using MonoReport.Models.Xlsx;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxTemplatePlanner
{
    private static ImmutableArray<XlsxLogicalRow> BindHierarchy(ImmutableArray<XlsxLogicalRow> rows)
    {
        if (rows.IsEmpty) return rows;
        var parents = new int[rows.Length];
        var children = new List<int>?[rows.Length];
        var stack = new List<int>();
        for (var index = 0; index < rows.Length; index++)
        {
            while (stack.Count > 0 && rows[stack[^1]].Level >= rows[index].Level)
                stack.RemoveAt(stack.Count - 1);
            var parent = stack.Count == 0 ? -1 : stack[^1];
            parents[index] = parent;
            if (parent >= 0) (children[parent] ??= []).Add(rows[index].FirstIndex);
            stack.Add(index);
        }

        // Child identities are source record indices, so immutable parent links can
        // share the already bound row without cycles or copied ancestor chains.
        var bound = ImmutableArray.CreateBuilder<XlsxLogicalRow>(rows.Length);
        for (var index = 0; index < rows.Length; index++)
            bound.Add(rows[index] with
            {
                Parent = parents[index] < 0 ? null : bound[parents[index]],
                ChildFirstIndices = children[index]?.ToImmutableArray() ?? []
            });
        return bound.ToImmutable();
    }
}
