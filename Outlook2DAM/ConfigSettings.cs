using System;
using System.ComponentModel;
using System.Configuration;
using System.Collections.Specialized;

namespace Outlook2DAM;

public class ConfigSettings
{
    private string? _clientSecret;

    public ConfigSettings()
    {
        LoadSettings();
    }

    [Category("Microsoft Graph")]
    [Description("ID do tenant no Azure AD")]
    public string TenantId { get; set; } = string.Empty;

    [Category("Microsoft Graph")]
    [Description("ID do cliente no Azure AD")]
    public string ClientId { get; set; } = string.Empty;

    [Category("Microsoft Graph")]
    [Description("Segredo do cliente no Azure AD")]
    [DisplayName("ClientSecret")]
    [PasswordPropertyText(true)]
    public string ClientSecret 
    { 
        get 
        {
            var secret = _clientSecret ?? string.Empty;
            if (string.IsNullOrEmpty(secret)) return string.Empty;
            return secret.Length > 5 ? $"{secret[..5]}{"*".PadRight(secret.Length - 5, '*')}" : secret;
        }
        set
        {
            _clientSecret = value;
            // Salvar no arquivo de configuração
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["ClientSecret"].Value = value;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
    }

    [Category("Microsoft Graph")]
    [Description("Emails para recolha (separados por ponto e vírgula)")]
    public string UserEmail { get; set; } = string.Empty;

    [Category("Microsoft Graph")]
    [Description("Pasta de entrada (Inbox ou email:pasta;email:pasta)")]
    public string InboxFolder { get; set; } = "Inbox";

    [Category("Serviço")]
    [Description("Intervalo em segundos entre verificações")]
    public int ServiceIntervalSeconds { get; set; }

    [Category("Serviço")]
    [Description("Timeout em segundos para testes de ligação")]
    public int ConnectionTestTimeoutSeconds { get; set; }

    [Category("Log")]
    [Description("Nível de log (Debug, Information, Warning, Error)")]
    public string LogLevel { get; set; } = "Debug";

    [Category("Log")]
    [Description("Número de dias para manter os logs")]
    public int LogRetentionDays { get; set; } = 31;

    [Category("Log")]
    [Description("Se true, continua escrevendo no mesmo arquivo de log. Se false, comprime o log anterior antes de iniciar um novo")]
    public bool RewriteLog { get; set; } = false;

    private void LoadSettings()
    {
        var config = ConfigurationManager.AppSettings;
        TenantId = config["TenantId"] ?? string.Empty;
        ClientId = config["ClientId"] ?? string.Empty;
        _clientSecret = config["ClientSecret"] ?? string.Empty;
        UserEmail = config["UserEmail"] ?? string.Empty;
        InboxFolder = config["InboxFolder"] ?? "Inbox";
        ServiceIntervalSeconds = int.Parse(config["ServiceIntervalSeconds"] ?? "60");
        ConnectionTestTimeoutSeconds = int.Parse(config["ConnectionTestTimeoutSeconds"] ?? "30");
        LogLevel = config["LogLevel"] ?? "Debug";
        LogRetentionDays = int.Parse(config["LogRetentionDays"] ?? "31");
        RewriteLog = bool.Parse(config["RewriteLog"] ?? "false");
    }

    public void LoadFromConfig(NameValueCollection config)
    {
        TenantId = config["TenantId"] ?? string.Empty;
        ClientId = config["ClientId"] ?? string.Empty;
        _clientSecret = config["ClientSecret"] ?? string.Empty;
        UserEmail = config["UserEmail"] ?? string.Empty;
        InboxFolder = config["InboxFolder"] ?? "Inbox";
        ServiceIntervalSeconds = int.Parse(config["ServiceIntervalSeconds"] ?? "60");
        ConnectionTestTimeoutSeconds = int.Parse(config["ConnectionTestTimeoutSeconds"] ?? "30");
        LogLevel = config["LogLevel"] ?? "Debug";
        LogRetentionDays = int.Parse(config["LogRetentionDays"] ?? "31");
        RewriteLog = bool.Parse(config["RewriteLog"] ?? "false");
    }
}
