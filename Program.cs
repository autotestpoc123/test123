using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace MorganStanley.COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 薄壳 Main:装配配置/日志/取消 → 取单实例锁 → 跑 PhotoImportJob → 映射退出码。
/// 所有业务逻辑在 PhotoImportJob(便于将来改挂 Hangfire 等宿主)。
/// 退出码:0 成功/跳过;1 业务错误;2 无法启动(配置/锁)。
/// </summary>
public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        using var loggerFactory = LoggerFactory.Create(b => b.AddSimpleConsole(o =>
        {
            o.SingleLine = true;
            o.TimestampFormat = "yyyy-MM-dd HH:mm:ss ";
        }).SetMinimumLevel(LogLevel.Information));
        var logger = loggerFactory.CreateLogger("PhotoImport");

        // Ctrl+C / 计划任务停止 → 优雅取消(§6 坑5)
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

        PhotoImportOptions options;
        try
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false)
                .AddEnvironmentVariables("FWD_PHOTO_")   // 允许环境变量覆盖
                .AddCommandLine(args)                     // 例:--PhotoImport:DryRun=false
                .Build();

            options = config.GetSection(PhotoImportOptions.SectionName).Get<PhotoImportOptions>()
                      ?? throw new InvalidOperationException($"缺少配置节 {PhotoImportOptions.SectionName}");
            options.Validate();
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "配置加载失败,无法启动");
            return 2;
        }

        // 单实例锁:拿不到说明上一轮仍在跑 → 正常跳过(§2 LK)
        using var mutex = SingleInstanceLock.TryAcquire(options.LockFilePath);
        if (mutex is null)
        {
            logger.LogWarning("上一轮仍在运行(锁被占用),本次跳过。lock={Lock}", options.LockFilePath);
            return 0;
        }

        try
        {
            var job = new PhotoImportJob(options, logger);
            var summary = await job.RunAsync(cts.Token);
            logger.LogInformation("完成:{Summary}", summary);
            return summary.Errors > 0 ? 1 : 0;
        }
        catch (OperationCanceledException)
        {
            logger.LogWarning("被取消(停机/超时),本轮未完成;因幂等下次可续跑");
            return 1;
        }
        catch (Exception ex)
        {
            // 含 C4a:IsReadyToLoad 对失效 zip 抛异常 → 走这里 → 退出 1,锁在 finally/using 释放
            logger.LogError(ex, "运行失败");
            return 1;
        }
    }
}
