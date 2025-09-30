using Serilog;
using Serilog.Core;
using Serilog.Events;
using System.IO;
using System.Configuration;

namespace Outlook2DAM.Services;

public static class LoggerService
{
    private static bool _isInitialized;
    private static readonly object _lock = new();

    public static void Initialize(bool rewriteLog = false)
    {
        if (_isInitialized)
            return;

        lock (_lock)
        {
            if (_isInitialized)
                return;

            // Ler configurações
            var config = ConfigurationManager.AppSettings;
            var logLevelStr = config["LogLevel"] ?? "Information";
            var logRetentionDays = int.TryParse(config["LogRetentionDays"], out var retention) ? retention : 31;
            var logPath = config["LogPath"] ?? "logs";

            // Parse do LogLevel
            if (!Enum.TryParse<LogEventLevel>(logLevelStr, ignoreCase: true, out var logLevel))
            {
                logLevel = LogEventLevel.Information;
            }

            // Criar path completo
            if (!Path.IsPathRooted(logPath))
            {
                logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logPath);
            }

            Directory.CreateDirectory(logPath);
            var logFile = Path.Combine(logPath, "outlook2dam-.log");

            // Configurar Serilog com enrichers e retention
            var loggerConfig = new LoggerConfiguration()
                .MinimumLevel.Is(logLevel)
                // Enrichers para contexto adicional
                .Enrich.WithMachineName()
                .Enrich.WithEnvironmentUserName()
                .Enrich.WithProcessId()
                .Enrich.WithThreadId()
                .Enrich.WithProperty("Application", "Outlook2DAM")
                .Enrich.WithProperty("Version", typeof(LoggerService).Assembly.GetName().Version?.ToString() ?? "1.0.0")
                .WriteTo.Console(
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] [{MachineName}] {Message:lj}{NewLine}{Exception}")
                .WriteTo.File(logFile,
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: logRetentionDays,
                    shared: rewriteLog,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] [Machine:{MachineName}] [User:{EnvironmentUserName}] [Process:{ProcessId}] [Thread:{ThreadId}] {Message:lj}{NewLine}{Exception}");

            Log.Logger = loggerConfig.CreateLogger();

            _isInitialized = true;

            Log.Information("Logger inicializado. Nível: {LogLevel}, Arquivo: {LogFile}, Retention: {RetentionDays} dias",
                logLevel, logFile, logRetentionDays);
        }
    }

    public static ILogger GetLogger<T>() where T : class
    {
        if (!_isInitialized)
            Initialize();

        return Log.ForContext<T>();
    }

    public static ILogger GetLogger(Type type)
    {
        if (!_isInitialized)
            Initialize();

        return Log.ForContext(type);
    }

    public static void CloseAndFlush()
    {
        Log.CloseAndFlush();
    }
}