using System.Globalization;

namespace MonoReport.Xlsx.Planning;

internal readonly record struct XlsxAddress(int Row, int Column)
{
    public override string ToString()
    {
        var column = Column;
        var letters = string.Empty;
        while (column > 0)
        {
            column--;
            letters = (char) ('A' + column % 26) + letters;
            column /= 26;
        }

        return letters + Row.ToString(CultureInfo.InvariantCulture);
    }
}

internal readonly record struct XlsxRectangle(XlsxAddress First, XlsxAddress Last)
{
    internal int Height => checked(Last.Row - First.Row + 1);
    internal int Width => checked(Last.Column - First.Column + 1);

    internal bool Contains(XlsxAddress cell)
    {
        return cell.Row >= First.Row && cell.Row <= Last.Row &&
               cell.Column >= First.Column && cell.Column <= Last.Column;
    }

    public override string ToString()
    {
        return First + ":" + Last;
    }
}