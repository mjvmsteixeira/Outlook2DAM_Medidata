using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using Oracle.ManagedDataAccess.Client;
using Serilog;
using System;
using System.Configuration;

namespace Outlook2DAM.Services;

public class ConnectionTester : IDisposable
{
    private readonly ILogger _logger;
    private readonly string _connectionString;
    private readonly int _timeoutSeconds;
    private readonly string _userEmail;
    private readonly GraphServiceClient _graphClient;
    private readonly TokenProvider _tokenProvider;

    public ConnectionTester()
    {
        _logger = new LoggerConfiguration()
            .WriteTo.Console()
            .WriteTo.File("logs/outlook2dam-.log", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        var config = ConfigurationManager.AppSettings;
        _timeoutSeconds = int.Parse(config["ConnectionTestTimeoutSeconds"] ?? "30");
        _connectionString = ConfigurationManager.ConnectionStrings["Outlook2DAM"].ConnectionString;
        _userEmail = config["UserEmail"] ?? string.Empty;

        var tenantId = config["TenantId"];
        var clientId = config["ClientId"];
        var clientSecret = config["ClientSecret"];

        var app = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithTenantId(tenantId)
            .WithClientSecret(clientSecret)
            .Build();

        var scopes = new[] { "https://graph.microsoft.com/.default" };

        _tokenProvider = new TokenProvider();
        var authProvider = new BaseBearerTokenAuthenticationProvider(_tokenProvider);
        _graphClient = new GraphServiceClient(authProvider);
    }

    public async Task<bool> TestAllConnections()
    {
        try
        {
            _logger.Information("Testing all connections...");

            var tasks = new[]
            {
                TestDatabaseConnection(),
                TestGraphConnection()
            };

            await Task.WhenAll(tasks);

            var allSuccessful = tasks.All(t => t.Result);
            _logger.Information("All connection tests {Result}", allSuccessful ? "passed" : "failed");
            return allSuccessful;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Error testing connections");
            return false;
        }
    }

    private async Task<bool> TestDatabaseConnection()
    {
        try
        {
            _logger.Debug("Testing database connection...");

            // Detectar provider
            var lowerConnectionString = _connectionString.ToLowerInvariant();
            var isOracle = lowerConnectionString.Contains("provider=oraoledb") || lowerConnectionString.Contains("oracle");
            var isSqlServer = lowerConnectionString.Contains("provider=sqloledb") || lowerConnectionString.Contains("provider=sqlncli") ||
                             lowerConnectionString.Contains("provider=msoledbsql") || lowerConnectionString.Contains("sql server");

            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(_timeoutSeconds));

            if (isOracle)
            {
                // Teste específico para Oracle
                using var connection = new OracleConnection(_connectionString);
                await connection.OpenAsync(cts.Token);
                _logger.Information("Database connection test passed (Oracle)");
            }
            else
            {
                // Teste genérico com OLEDB para SQL Server e outros
                using var connection = new System.Data.OleDb.OleDbConnection(_connectionString);
                await connection.OpenAsync(cts.Token);

                var provider = isSqlServer ? "SQL Server" : "Unknown Provider";
                _logger.Information("Database connection test passed ({Provider})", provider);
            }

            return true;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Database connection test failed");
            return false;
        }
    }

    private async Task<bool> TestGraphConnection()
    {
        try
        {
            _logger.Debug("A testar ligação com Microsoft Graph...");

            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(_timeoutSeconds));

            var messages = await _graphClient.Users[_userEmail]
                .Messages
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Top = 1;
                }, cts.Token);

            if (messages?.Value?.Any() ?? false)
            {
                _logger.Information("A testar a ligação com Microsoft Graph testada com sucesso");
                return true;
            }

            _logger.Error("Não foi possível obter informações do utilizador");
            return false;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao testar ligação com Microsoft Graph");
            return false;
        }
    }

    public void Dispose()
    {
        _graphClient?.Dispose();
        GC.SuppressFinalize(this);
    }
}
