using Microsoft.Extensions.Logging;
using Xunit;

namespace COD.FirmwideDirectory.PhotoImportTool.Verify;

// Unlike NullLogger, exercise the same message formatter used by console logging.
// Do not retain user data or emit test logs to disk.
internal sealed class FormattingLogger : ILogger
{
    public static FormattingLogger Instance { get; } = new();

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
    public bool IsEnabled(LogLevel logLevel) => true;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
        Exception? exception, Func<TState, Exception?, string> formatter)
    {
        _ = formatter(state, exception);
    }
}

public class LoggingFormatTests
{
    [Fact]
    public void Logger_rejects_missing_arguments()
    {
        ILogger logger = FormattingLogger.Instance;
#pragma warning disable CA2017 // Reproduce the remote failure: more placeholders than values.
        Assert.Throws<FormatException>(() =>
            logger.LogInformation("XML planned msid={Msid} reason={Reason}", "59NVN"));
#pragma warning restore CA2017
    }

    [Fact]
    public void Trigger_details_are_a_value_not_a_message_template()
    {
        ILogger logger = FormattingLogger.Instance;
        logger.LogInformation("XML triggers {Details}",
            "photoChanged=True usersChanged=True xmlChanged=True manifestMissing=False");
        // Braces in a value must never be parsed as template placeholders.
        logger.LogInformation("XML triggers {Details}", "literal {unclosed");
    }

    [Fact]
    public void Logger_formats_named_properties_and_nullable_versions()
    {
        ILogger logger = FormattingLogger.Instance;
        DateTimeOffset? previous = null;
        logger.LogInformation("version={Version:o} previous={Previous:o} exists={Exists}",
            DateTimeOffset.UtcNow, previous, false);
    }

    [Fact]
    public void Logger_rejects_malformed_placeholder()
    {
        ILogger logger = FormattingLogger.Instance;
#pragma warning disable CA2023 // Deliberately malformed template verifies the regression guard.
        Assert.Throws<FormatException>(() =>
            logger.LogInformation("XML decision applyWrites={ApplyWrites", true));
#pragma warning restore CA2023
    }
}
