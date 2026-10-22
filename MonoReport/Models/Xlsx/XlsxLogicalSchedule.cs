using System.Collections.Immutable;

namespace MonoReport.Models.Xlsx;

internal enum XlsxLogicalBlockRole { Header, Body, Footer }

internal sealed record XlsxLogicalBlockInstance(XlsxVerticalBlock Source, XlsxLogicalBlockRole Role,
    VerticalInstancePath Owner);
internal sealed record XlsxVerticalRegionSchedule(int RegionId, ImmutableArray<XlsxLogicalBlockInstance> Blocks);
internal sealed record XlsxLogicalSheetSchedule(uint SheetId, ImmutableArray<XlsxVerticalRegionSchedule> Regions);
