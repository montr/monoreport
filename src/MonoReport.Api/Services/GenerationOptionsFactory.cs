using System.Globalization;

namespace MonoReport.Api.Services;

public class GenerationOptionsFactory
{
    public virtual GenerationOptions Create(CultureInfo culture)
    {
        return new GenerationOptions {Culture = culture};
    }
}