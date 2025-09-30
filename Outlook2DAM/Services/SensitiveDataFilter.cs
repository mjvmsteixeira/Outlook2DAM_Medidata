using System.Text.RegularExpressions;

namespace Outlook2DAM.Services;

/// <summary>
/// Filtro para mascarar dados sensíveis em logs e outputs
/// </summary>
public static class SensitiveDataFilter
{
    private static readonly string[] SensitiveKeys = new[]
    {
        "password", "pwd", "secret", "token", "key", "credential",
        "clientsecret", "connectionstring", "apikey", "authorization"
    };

    private static readonly Regex EmailRegex = new(@"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b", RegexOptions.Compiled);
    private static readonly Regex ConnectionStringPasswordRegex = new(@"(password|pwd)\s*=\s*[^;]+", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex TokenRegex = new(@"[A-Za-z0-9_~\-]{20,}", RegexOptions.Compiled);

    /// <summary>
    /// Mascara um valor se a chave for sensível
    /// </summary>
    public static string MaskIfSensitive(string key, string? value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        if (IsSensitiveKey(key))
        {
            return MaskValue(value);
        }

        return value;
    }

    /// <summary>
    /// Verifica se uma chave é sensível
    /// </summary>
    public static bool IsSensitiveKey(string key)
    {
        if (string.IsNullOrEmpty(key))
            return false;

        var lowerKey = key.ToLowerInvariant().Replace("_", "").Replace("-", "");

        return SensitiveKeys.Any(sensitiveKey =>
            lowerKey.Contains(sensitiveKey.Replace("_", "").Replace("-", "")));
    }

    /// <summary>
    /// Mascara um valor sensível
    /// </summary>
    public static string MaskValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        if (value.Length <= 4)
            return "****";

        // Mostra apenas primeiros 4 caracteres
        return $"{value[..4]}{"*".PadRight(Math.Min(value.Length - 4, 20), '*')}";
    }

    /// <summary>
    /// Sanitiza uma connection string removendo passwords
    /// </summary>
    public static string SanitizeConnectionString(string connectionString)
    {
        if (string.IsNullOrEmpty(connectionString))
            return string.Empty;

        // Substituir passwords por asteriscos
        var sanitized = ConnectionStringPasswordRegex.Replace(connectionString, "$1=****");

        return sanitized;
    }

    /// <summary>
    /// Mascara endereços de email (opcional, para GDPR)
    /// </summary>
    public static string MaskEmails(string text, bool preserveDomain = true)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        return EmailRegex.Replace(text, match =>
        {
            var email = match.Value;
            var parts = email.Split('@');

            if (parts.Length != 2)
                return email;

            var localPart = parts[0];
            var domain = parts[1];

            // Mostrar apenas primeiros 2 caracteres do local part
            var maskedLocal = localPart.Length > 2
                ? $"{localPart[..2]}***"
                : "***";

            return preserveDomain
                ? $"{maskedLocal}@{domain}"
                : $"{maskedLocal}@***";
        });
    }

    /// <summary>
    /// Remove tokens e secrets de texto livre
    /// </summary>
    public static string RemoveTokens(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        // Substituir sequências longas que parecem tokens
        return TokenRegex.Replace(text, match =>
        {
            var token = match.Value;

            // Se parece com um token (longo e alfanumérico)
            if (token.Length >= 30 && token.Any(char.IsDigit) && token.Any(char.IsLetter))
            {
                return $"{token[..4]}***[REDACTED]";
            }

            return token;
        });
    }

    /// <summary>
    /// Sanitiza um objeto de configuração para logging
    /// </summary>
    public static Dictionary<string, string> SanitizeConfig(Dictionary<string, string?> config)
    {
        var sanitized = new Dictionary<string, string>();

        foreach (var kvp in config)
        {
            sanitized[kvp.Key] = MaskIfSensitive(kvp.Key, kvp.Value);
        }

        return sanitized;
    }

    /// <summary>
    /// Cria mensagem de log segura
    /// </summary>
    public static string CreateSafeLogMessage(string message, params object[] args)
    {
        try
        {
            var formatted = string.Format(message, args);

            // Sanitizar connection strings
            formatted = SanitizeConnectionString(formatted);

            // Remover tokens
            formatted = RemoveTokens(formatted);

            return formatted;
        }
        catch
        {
            return "[Erro ao formatar mensagem de log]";
        }
    }
}