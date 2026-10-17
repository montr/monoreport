using System.Globalization;
namespace MonoReport.Services
{
    /// <summary>Russian RUB amounts, rounded to kopecks away from zero.</summary>
    public sealed class RubSpellOutFormatter
    {
        private const string FeminineOne = "одна";
        private const string FeminineTwo = "две";
        private const string NegativePrefix = "минус ";
        private const string Zero = "ноль";
        private const string KopecksFormat = "00";
        private static readonly string[] KopeckForms = ["копейка", "копейки", "копеек"];

        private static readonly string[] Ones = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"];
        private static readonly string[] Teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"];
        private static readonly string[] Tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"];
        private static readonly string[] Hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"];
        private static readonly string[][] Scales = [["рубль", "рубля", "рублей"], ["тысяча", "тысячи", "тысяч"], ["миллион", "миллиона", "миллионов"], ["миллиард", "миллиарда", "миллиардов"], ["триллион", "триллиона", "триллионов"], ["квадриллион", "квадриллиона", "квадриллионов"]];

        internal string Format(decimal amount)
        {
            amount = decimal.Round(amount, 2, MidpointRounding.AwayFromZero);
            var absolute = Math.Abs(amount);
            if (absolute >= 1_000_000_000_000_000_000m)
            {
                throw new FormatException("spellOut supports absolute amounts below 10^18 RUB.");
            }

            var whole = (long)decimal.Truncate(absolute);
            var kopecks = (int)((absolute - whole) * 100);
            var groups = new List<string>();
            var remaining = whole;
            for (var scale = 0; remaining > 0; scale++, remaining /= 1000)
            {
                var group = (int)(remaining % 1000);
                if (group == 0)
                {
                    continue;
                }

                var words = new List<string>();
                if (group / 100 > 0)
                {
                    words.Add(Hundreds[group / 100]);
                }

                var rest = group % 100;
                if (rest is >= 10 and <= 19)
                {
                    words.Add(Teens[rest - 10]);
                }
                else
                {
                    if (rest / 10 > 0)
                    {
                        words.Add(Tens[rest / 10]);
                    }

                    if (rest % 10 > 0)
                    {
                        words.Add(scale == 1 && rest % 10 is 1 or 2 ? (rest % 10 == 1 ? FeminineOne : FeminineTwo) : Ones[rest % 10]);
                    }
                }

                if (scale > 0)
                {
                    words.Add(Decline(group, Scales[scale]));
                }

                groups.Insert(0, string.Join(' ', words));
            }

            var result = (amount < 0 ? NegativePrefix : string.Empty) + (whole == 0 ? Zero : string.Join(' ', groups)) + " " + Decline(whole, Scales[0]) + " " + kopecks.ToString(KopecksFormat, CultureInfo.InvariantCulture) + " " + Decline(kopecks, KopeckForms);
            return char.ToUpperInvariant(result[0]) + result[1..];
        }

        private static string Decline(long value, string[] forms) => forms[value % 100 is >= 11 and <= 14 ? 2 : value % 10 == 1 ? 0 : value % 10 is >= 2 and <= 4 ? 1 : 2];
    }
}
