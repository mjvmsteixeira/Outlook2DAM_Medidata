using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;
using Serilog;
using System.Configuration;
using Outlook2DAM.Services;
using System.Collections.Generic;

namespace Outlook2DAM;

internal static class Program
{
    private static readonly ILogger _logger;
    private static bool _isCliMode;

    static Program()
    {
        _logger = LoggerService.GetLogger(typeof(Program));
    }

    [STAThread]
    static async Task Main(string[] args)
    {
        try
        {
            _isCliMode = args.Contains("--cli");

            // Inicializar logger
            var rewriteLog = bool.TryParse(ConfigurationManager.AppSettings["RewriteLog"], out var rewrite) && rewrite;
            LoggerService.Initialize(rewriteLog);

            _logger.Information("A iniciar a aplicação em modo {Mode}...", _isCliMode ? "CLI" : "GUI");

            // Carregar e validar configurações críticas
            var settings = ConfigurationManager.AppSettings;

            var appConfig = new AppConfiguration
            {
                TenantId = settings["TenantId"] ?? string.Empty,
                ClientId = settings["ClientId"] ?? string.Empty,
                ClientSecret = settings["ClientSecret"] ?? string.Empty,
                UserEmail = settings["UserEmail"] ?? string.Empty,
                TempFolder = settings["TempFolder"] ?? string.Empty,
                ServiceIntervalSeconds = int.TryParse(settings["ServiceIntervalSeconds"], out var interval) ? interval : 60,
                ConnectionTestTimeoutSeconds = int.TryParse(settings["ConnectionTestTimeoutSeconds"], out var timeout) ? timeout : 30,
                EmailsPerCycle = int.TryParse(settings["EmailsPerCycle"], out var emails) ? emails : 1,
                MaxRetries = int.TryParse(settings["MaxRetries"], out var retries) ? retries : 3,
                LogRetentionDays = int.TryParse(settings["LogRetentionDays"], out var retention) ? retention : 31
            };

            // Validar toda a configuração
            _logger.Information("A validar configurações da aplicação...");
            var validationResult = InputValidator.ValidateConfiguration(appConfig);

            if (!validationResult.IsValid)
            {
                var errorMsg = $"Configuração inválida: {validationResult.ErrorMessage}";
                _logger.Error(errorMsg);
                throw new InvalidOperationException(errorMsg);
            }

            _logger.Information("Configurações validadas com sucesso!");

            // Logar configurações com valores sensíveis mascarados
            _logger.Debug("TenantId: {TenantId}", appConfig.TenantId);
            _logger.Debug("ClientId: {ClientId}", appConfig.ClientId);
            _logger.Debug("ClientSecret: {ClientSecret}", SensitiveDataFilter.MaskValue(appConfig.ClientSecret));
            _logger.Debug("UserEmail: {UserEmail}", appConfig.UserEmail);
            _logger.Debug("TempFolder: {TempFolder}", appConfig.TempFolder);
            _logger.Debug("ServiceIntervalSeconds: {ServiceIntervalSeconds}", appConfig.ServiceIntervalSeconds);

            var tenantId = appConfig.TenantId;
            var clientId = appConfig.ClientId;
            var clientSecret = appConfig.ClientSecret;

            // Configurar autenticação Microsoft Graph
            _logger.Debug("A iniciar o TokenProvider...");
            var tokenProvider = new TokenProvider();
            var authProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
            _logger.Information("Cliente Microsoft Graph configurado com sucesso!");

            // Configurar serviços
            var services = new ServiceCollection();
            services.AddSingleton<OutlookService>();
            services.AddSingleton<HealthCheckService>();
            services.AddSingleton<GraphServiceClient>(sp =>
            {
                var tokenProvider = new TokenProvider();
                var authProvider = new BaseBearerTokenAuthenticationProvider(tokenProvider);
                return new GraphServiceClient(authProvider);
            });

            if (!_isCliMode)
            {
                ApplicationConfiguration.Initialize();
                services.AddSingleton<MainForm>();
            }

            var serviceProvider = services.BuildServiceProvider();

            // Executar health check inicial
            _logger.Information("A executar verificação de saúde inicial...");
            var healthCheckService = serviceProvider.GetRequiredService<HealthCheckService>();
            var healthResult = await healthCheckService.CheckHealthAsync();

            if (!healthResult.IsHealthy)
            {
                _logger.Warning("Health check detectou problemas:\n{HealthCheck}", healthResult);

                if (!_isCliMode)
                {
                    var failedChecks = string.Join("\n", healthResult.Checks
                        .Where(c => !c.Value.IsHealthy)
                        .Select(c => $"• {c.Key}: {c.Value.Message}"));

                    var result = MessageBox.Show(
                        $"Health check detectou os seguintes problemas:\n\n{failedChecks}\n\nDeseja continuar mesmo assim?",
                        "Aviso - Health Check",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        _logger.Information("Aplicação encerrada pelo utilizador devido a falhas no health check");
                        return;
                    }
                }
                else
                {
                    _logger.Warning("Continuando mesmo com falhas no health check...");
                }
            }
            else
            {
                _logger.Information("Health check passou: Sistema pronto para operar");
            }

            if (_isCliMode)
            {
                RunCliMode(serviceProvider);
            }
            else
            {
                var mainForm = serviceProvider.GetRequiredService<MainForm>();
                Application.Run(mainForm);
            }
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "Erro fatal na aplicação");
            if (!_isCliMode)
            {
                MessageBox.Show($"Erro fatal na aplicação: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        finally
        {
            LoggerService.CloseAndFlush();
        }
    }

    private static void RunCliMode(IServiceProvider serviceProvider)
    {
        _logger.Information("Iniciando serviço em modo CLI...");
        var outlookService = serviceProvider.GetRequiredService<OutlookService>();
        
        // Criar um CancellationTokenSource para controlar o encerramento
        using var cts = new CancellationTokenSource();
        
        // Configurar o handler para o evento de encerramento
        Console.CancelKeyPress += (s, e) =>
        {
            e.Cancel = true;
            cts.Cancel();
        };

        AppDomain.CurrentDomain.ProcessExit += (s, e) =>
        {
            cts.Cancel();
        };

        try
        {
            // Iniciar o serviço em background
            outlookService.StartService(cts.Token);

            // Manter o processo rodando até receber sinal de cancelamento
            while (!cts.Token.IsCancellationRequested)
            {
                Thread.Sleep(1000);
            }
        }
        finally
        {
            _logger.Information("Encerrando serviço em modo CLI...");
            outlookService.StopService();
        }
    }
}
