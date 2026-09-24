using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Console;

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// Thin entry point: configure settings, logging, and cancellation; acquire the single-instance lock;
/// run PhotoImportJob; map the result to an exit code.
/// All business logic resides in PhotoImportJob to support future hosting in Hangfire or other hosts.
/// Exit codes: 0 = success/skipped; 1 = runtime error; 2 = startup failure (configuration/lock).
/// </summary>
public static class Program
{
    public static int Main(string[] args)
    {
        using var loggerFactory = LoggerFactory.Create(b => b.AddSimpleConsole(o =>
        {
            o.SingleLine = true;
            o.TimestampFormat = "yyyy-MM-dd HH:mm:ss ";
        }).SetMinimumLevel(LogLevel.Information)
          .AddFilter<ConsoleLoggerProvider>(level => level >= LogLevel.Warning));
        var logger = loggerFactory.CreateLogger("PhotoImport");
        LocalFileLoggerProvider? fileLogger = null;

        // Request graceful cancellation on Ctrl+C / scheduled task interruption (section 6, pitfall 5).
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

        PhotoImportOptions options;
        try
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false)
                .AddEnvironmentVariables("FWD_PHOTO_")   // Allow environment variable overrides.
                .AddCommandLine(args)                     // Example: --PhotoImport:DryRun=false
                .Build();

            options = config.GetSection(PhotoImportOptions.SectionName).Get<PhotoImportOptions>()
                      ?? throw new InvalidOperationException($"Missing configuration section {PhotoImportOptions.SectionName}");
            fileLogger = new LocalFileLoggerProvider(options.ResolveLogDirectory(), options.LogRetentionDays, options.LogMaxFileBytes);
            loggerFactory.AddProvider(fileLogger);
            logger.LogInformation("Run started runId={RunId} logFile={LogFile} dryRun={DryRun}",
                fileLogger.RunId, fileLogger.CurrentPath, options.DryRun);
            options.Validate();
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Failed to load configuration; unable to start");
            return 2;
        }

        // Single-instance lock: contention (IOException -> null) means a previous run is still active; skip normally (section 2, LK).
        // Failure to create/access the lock file (e.g. insufficient permissions) is a startup failure, not contention; exit 2 instead of skipping.
        SingleInstanceLock? mutex;
        try
        {
            mutex = SingleInstanceLock.TryAcquire(options.LockFilePath);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Unable to create/access lock file {Lock} (check permissions); unable to start", options.LockFilePath);
            return 2;
        }

        if (mutex is null)
        {
            logger.LogWarning("A previous run is still active (lock held); skipping this run. lock={Lock}", options.LockFilePath);
            fileLogger.Dispose();
            return fileLogger.HasFailed ? 1 : 0;
        }

        using (mutex)
        {
            try
            {
                var job = new PhotoImportJob(options, logger);
                var summary = job.Run(cts.Token);
                logger.LogInformation("Completed: {Summary}", summary);
                logger.LogInformation("Run finished runId={RunId} exitCode={ExitCode}",
                    fileLogger.RunId, summary.Errors > 0 || fileLogger.HasFailed ? 1 : 0);
                fileLogger.Dispose();
                return summary.Errors > 0 || fileLogger.HasFailed ? 1 : 0;
            }
            catch (OperationCanceledException)
            {
                logger.LogWarning("Canceled (shutdown/timeout); this run is incomplete. Idempotent processing allows a subsequent run to continue");
                return 1;
            }
            catch (Exception ex)
            {
                // Includes C4a: CheckZip throws for an invalid ZIP; exit 1 here and release the lock through using.
                logger.LogError(ex, "Run failed");
                return 1;
            }
        }
    }
}
