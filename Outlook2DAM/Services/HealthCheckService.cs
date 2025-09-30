using Serilog;
using System.Configuration;
using System.Data.OleDb;
using Microsoft.Graph;

namespace Outlook2DAM.Services;

/// <summary>
/// Serviço para verificar a saúde da aplicação e dependências
/// </summary>
public class HealthCheckService
{
    private readonly ILogger _logger;
    private readonly GraphServiceClient _graphClient;

    public HealthCheckService(GraphServiceClient graphClient)
    {
        _logger = LoggerService.GetLogger<HealthCheckService>();
        _graphClient = graphClient;
    }

    /// <summary>
    /// Executa todas as verificações de saúde
    /// </summary>
    public async Task<HealthCheckResult> CheckHealthAsync()
    {
        _logger.Information("A iniciar verificação de saúde da aplicação...");

        var result = new HealthCheckResult
        {
            Timestamp = DateTime.UtcNow
        };

        // Verificar configurações
        var configCheck = CheckConfiguration();
        result.Checks.Add("Configuration", configCheck);

        // Verificar acesso ao TempFolder
        var folderCheck = CheckTempFolder();
        result.Checks.Add("TempFolder", folderCheck);

        // Verificar conexão com Microsoft Graph
        var graphCheck = await CheckMicrosoftGraphAsync();
        result.Checks.Add("MicrosoftGraph", graphCheck);

        // Verificar conexão com base de dados
        var dbCheck = CheckDatabaseConnection();
        result.Checks.Add("Database", dbCheck);

        result.IsHealthy = result.Checks.All(c => c.Value.IsHealthy);

        if (result.IsHealthy)
        {
            _logger.Information("Verificação de saúde concluída: Sistema saudável");
        }
        else
        {
            var failedChecks = result.Checks.Where(c => !c.Value.IsHealthy).Select(c => c.Key);
            _logger.Warning("Verificação de saúde concluída: Falhas detectadas em {FailedChecks}",
                string.Join(", ", failedChecks));
        }

        return result;
    }

    /// <summary>
    /// Verifica se todas as configurações obrigatórias estão presentes
    /// </summary>
    private CheckStatus CheckConfiguration()
    {
        try
        {
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

            var validationResult = InputValidator.ValidateConfiguration(appConfig);

            if (!validationResult.IsValid)
            {
                return CheckStatus.Unhealthy($"Configuração inválida: {validationResult.ErrorMessage}");
            }

            return CheckStatus.Healthy("Configuração validada com sucesso");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao verificar configuração");
            return CheckStatus.Unhealthy($"Erro ao verificar configuração: {ex.Message}");
        }
    }

    /// <summary>
    /// Verifica acesso ao TempFolder
    /// </summary>
    private CheckStatus CheckTempFolder()
    {
        try
        {
            var tempFolder = ConfigurationManager.AppSettings["TempFolder"];

            if (string.IsNullOrEmpty(tempFolder))
            {
                return CheckStatus.Unhealthy("TempFolder não configurado");
            }

            // Tentar criar o diretório se não existir
            if (!Directory.Exists(tempFolder))
            {
                Directory.CreateDirectory(tempFolder);
                _logger.Information("TempFolder criado: {TempFolder}", tempFolder);
            }

            // Verificar permissões de escrita
            var testFile = Path.Combine(tempFolder, $"healthcheck_{Guid.NewGuid()}.tmp");
            File.WriteAllText(testFile, "health check test");
            File.Delete(testFile);

            return CheckStatus.Healthy($"TempFolder acessível: {tempFolder}");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao verificar TempFolder");
            return CheckStatus.Unhealthy($"TempFolder inacessível: {ex.Message}");
        }
    }

    /// <summary>
    /// Verifica conexão com Microsoft Graph API
    /// </summary>
    private async Task<CheckStatus> CheckMicrosoftGraphAsync()
    {
        try
        {
            var userEmail = ConfigurationManager.AppSettings["UserEmail"]?.Split(';').First().Trim();

            if (string.IsNullOrEmpty(userEmail))
            {
                return CheckStatus.Unhealthy("UserEmail não configurado");
            }

            // Para shared mailboxes, não podemos usar Users[email].GetAsync()
            // Em vez disso, tentamos acessar as mensagens diretamente
            try
            {
                var messages = await _graphClient.Users[userEmail].Messages
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Top = 1;
                    });

                // Se conseguimos acessar as mensagens, a mailbox está acessível
                return CheckStatus.Healthy($"Microsoft Graph API acessível (mailbox: {userEmail})");
            }
            catch (Exception ex) when (ex.Message.Contains("does not exist"))
            {
                // Tentar como user normal (não shared mailbox)
                var user = await _graphClient.Users[userEmail].GetAsync();

                if (user == null)
                {
                    return CheckStatus.Unhealthy($"Mailbox não encontrada: {userEmail}");
                }

                return CheckStatus.Healthy($"Microsoft Graph API acessível (usuário: {user.DisplayName})");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao verificar Microsoft Graph API");
            return CheckStatus.Unhealthy($"Microsoft Graph API inacessível: {ex.Message}");
        }
    }

    /// <summary>
    /// Verifica conexão com base de dados
    /// </summary>
    private CheckStatus CheckDatabaseConnection()
    {
        try
        {
            var connectionString = ConfigurationManager.ConnectionStrings["Outlook2DAM"]?.ConnectionString;

            if (string.IsNullOrEmpty(connectionString))
            {
                return CheckStatus.Unhealthy("Connection string não configurada");
            }

            using var connection = new OleDbConnection(connectionString);
            connection.Open();

            // Detectar provider e executar query apropriada
            var isOracle = connectionString.ToLowerInvariant().Contains("oracle");
            using var command = connection.CreateCommand();
            command.CommandText = isOracle ? "SELECT 1 FROM DUAL" : "SELECT 1";
            command.CommandTimeout = 10;
            command.ExecuteScalar();

            return CheckStatus.Healthy("Base de dados acessível");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao verificar conexão com base de dados");
            return CheckStatus.Unhealthy($"Base de dados inacessível: {ex.Message}");
        }
    }

    /// <summary>
    /// Verifica conectividade de rede
    /// </summary>
    public static async Task<bool> CheckNetworkConnectivityAsync(string host = "graph.microsoft.com", int timeoutSeconds = 5)
    {
        try
        {
            using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(timeoutSeconds) };
            var response = await httpClient.GetAsync($"https://{host}");
            return response.IsSuccessStatusCode || response.StatusCode == System.Net.HttpStatusCode.Unauthorized;
        }
        catch
        {
            return false;
        }
    }
}

/// <summary>
/// Resultado completo da verificação de saúde
/// </summary>
public class HealthCheckResult
{
    public DateTime Timestamp { get; set; }
    public bool IsHealthy { get; set; }
    public Dictionary<string, CheckStatus> Checks { get; set; } = new();

    public override string ToString()
    {
        var status = IsHealthy ? "HEALTHY" : "UNHEALTHY";
        var details = string.Join("\n  ", Checks.Select(c => $"{c.Key}: {c.Value}"));
        return $"Health Check ({Timestamp:yyyy-MM-dd HH:mm:ss}): {status}\n  {details}";
    }
}

/// <summary>
/// Status de uma verificação individual
/// </summary>
public class CheckStatus
{
    public bool IsHealthy { get; private set; }
    public string Message { get; private set; }

    private CheckStatus(bool isHealthy, string message)
    {
        IsHealthy = isHealthy;
        Message = message;
    }

    public static CheckStatus Healthy(string message) => new(true, message);
    public static CheckStatus Unhealthy(string message) => new(false, message);

    public override string ToString() => $"{(IsHealthy ? "✓" : "✗")} {Message}";
}