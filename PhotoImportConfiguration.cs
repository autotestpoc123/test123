using Microsoft.Extensions.Configuration;

namespace COD.FirmwideDirectory.PhotoImportTool;

internal static class PhotoImportConfiguration
{
    public const string ExternalFileVariable = "FWD_PHOTO_CONFIG_FILE";

    public static IConfigurationRoot Load(string baseDirectory)
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(baseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false);

        var externalPath = Environment.GetEnvironmentVariable(ExternalFileVariable);
        if (!string.IsNullOrWhiteSpace(externalPath))
        {
            if (!Path.IsPathFullyQualified(externalPath))
                throw new ArgumentException($"{ExternalFileVariable} must be an absolute path.");

            // Explicitly configured files are required; never silently fall back on a typo.
            builder.AddJsonFile(externalPath, optional: false, reloadOnChange: false);
        }

        // Business settings come only from JSON; the environment selects the external file only.
        return builder.Build();
    }
}
