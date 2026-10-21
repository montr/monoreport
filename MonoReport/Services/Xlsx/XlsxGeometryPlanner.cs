using MonoReport.Models;

namespace MonoReport.Services.Xlsx;

internal sealed partial class XlsxGeometryPlanner
{
    internal (int Rows, int Columns) Dimensions(long sourceRows, long sourceColumns, long rowCopies, long columnCopies)
    {
        try
        {
            var rows = checked(sourceRows * rowCopies);
            var columns = checked(sourceColumns * columnCopies);
            if (sourceRows < 0 || sourceColumns < 0 || rowCopies < 0 || columnCopies < 0 || rows > 1048576 || columns > 16384)
                throw new OverflowException();
            return (checked((int)rows), checked((int)columns));
        }
        catch (OverflowException)
        {
            throw new TemplateException("Reduce the generated worksheet dimensions.", code: DiagnosticCodes.XlsxDimensionLimit);
        }
    }
}
