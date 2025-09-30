using System.Text.RegularExpressions;

namespace Outlook2DAM.Services;

/// <summary>
/// Validador de inputs para prevenir configurações inválidas e vulnerabilidades de segurança
/// </summary>
public static class InputValidator
{
    private static readonly Regex EmailRegex = new(
        @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex PathTraversalRegex = new(
        @"\.\.|[\\/]{2,}|[<>:""|?*]",
        RegexOptions.Compiled);

    /// <summary>
    /// Valida formato de endereço de email
    /// </summary>
    public static bool IsValidEmail(string email)
    {
        if (string.IsNullOrWhiteSpace(email))
            return false;

        // Verificar tamanho máximo (RFC 5321)
        if (email.Length > 254)
            return false;

        return EmailRegex.IsMatch(email);
    }

    /// <summary>
    /// Valida múltiplos emails separados por ponto-e-vírgula
    /// </summary>
    public static ValidationResult ValidateEmails(string emails, string fieldName = "Email")
    {
        if (string.IsNullOrWhiteSpace(emails))
            return ValidationResult.Fail($"{fieldName} não pode estar vazio");

        var emailList = emails.Split(';', StringSplitOptions.RemoveEmptyEntries)
                              .Select(e => e.Trim())
                              .ToList();

        if (emailList.Count == 0)
            return ValidationResult.Fail($"{fieldName} não contém endereços válidos");

        var invalidEmails = emailList.Where(e => !IsValidEmail(e)).ToList();

        if (invalidEmails.Any())
        {
            return ValidationResult.Fail(
                $"{fieldName} contém endereços inválidos: {string.Join(", ", invalidEmails)}");
        }

        return ValidationResult.Success();
    }

    /// <summary>
    /// Valida path prevenindo path traversal attacks
    /// </summary>
    public static ValidationResult ValidatePath(string path, string fieldName = "Path", bool mustExist = false)
    {
        if (string.IsNullOrWhiteSpace(path))
            return ValidationResult.Fail($"{fieldName} não pode estar vazio");

        // Verificar se é UNC path válido ou path local
        var isUncPath = path.StartsWith(@"\\") || path.StartsWith("//");
        var isLocalPath = Path.IsPathRooted(path) && !isUncPath;

        if (!isUncPath && !isLocalPath)
        {
            return ValidationResult.Fail(
                $"{fieldName} deve ser um caminho absoluto válido (local ou UNC)");
        }

        // Para UNC paths, remover o prefixo \\ antes de validar
        // Para paths locais, validar normalmente
        var pathToValidate = isUncPath ? path.Substring(2) : path;

        // Verificar caracteres suspeitos e path traversal (exceto caracteres válidos em paths)
        if (pathToValidate.Contains("..") || pathToValidate.Contains("<") ||
            pathToValidate.Contains(">") || pathToValidate.Contains("\"") ||
            pathToValidate.Contains("|") || pathToValidate.Contains("?") ||
            pathToValidate.Contains("*"))
        {
            return ValidationResult.Fail(
                $"{fieldName} contém caracteres inválidos ou tentativa de path traversal");
        }

        // Verificar se existe (opcional)
        if (mustExist)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    return ValidationResult.Fail($"{fieldName} não existe: {path}");
                }
            }
            catch (Exception ex)
            {
                return ValidationResult.Fail($"{fieldName} não pode ser acessado: {ex.Message}");
            }
        }

        return ValidationResult.Success();
    }

    /// <summary>
    /// Valida intervalo numérico positivo
    /// </summary>
    public static ValidationResult ValidatePositiveInteger(int value, string fieldName, int minValue = 1, int? maxValue = null)
    {
        if (value < minValue)
        {
            return ValidationResult.Fail($"{fieldName} deve ser maior ou igual a {minValue}");
        }

        if (maxValue.HasValue && value > maxValue.Value)
        {
            return ValidationResult.Fail($"{fieldName} deve ser menor ou igual a {maxValue.Value}");
        }

        return ValidationResult.Success();
    }

    /// <summary>
    /// Valida toda a configuração da aplicação
    /// </summary>
    public static ValidationResult ValidateConfiguration(AppConfiguration config)
    {
        var errors = new List<string>();

        // Validar TenantId (GUID)
        if (string.IsNullOrWhiteSpace(config.TenantId) || !Guid.TryParse(config.TenantId, out _))
        {
            errors.Add("TenantId deve ser um GUID válido");
        }

        // Validar ClientId (GUID)
        if (string.IsNullOrWhiteSpace(config.ClientId) || !Guid.TryParse(config.ClientId, out _))
        {
            errors.Add("ClientId deve ser um GUID válido");
        }

        // Validar ClientSecret (não pode estar vazio)
        if (string.IsNullOrWhiteSpace(config.ClientSecret) || config.ClientSecret.Length < 10)
        {
            errors.Add("ClientSecret deve ter pelo menos 10 caracteres");
        }

        // Validar UserEmail
        var emailResult = ValidateEmails(config.UserEmail, "UserEmail");
        if (!emailResult.IsValid)
        {
            errors.Add(emailResult.ErrorMessage!);
        }

        // Validar TempFolder
        var pathResult = ValidatePath(config.TempFolder, "TempFolder", mustExist: false);
        if (!pathResult.IsValid)
        {
            errors.Add(pathResult.ErrorMessage!);
        }

        // Validar ServiceIntervalSeconds
        var intervalResult = ValidatePositiveInteger(
            config.ServiceIntervalSeconds,
            "ServiceIntervalSeconds",
            minValue: 10,
            maxValue: 3600);
        if (!intervalResult.IsValid)
        {
            errors.Add(intervalResult.ErrorMessage!);
        }

        // Validar ConnectionTestTimeoutSeconds
        var timeoutResult = ValidatePositiveInteger(
            config.ConnectionTestTimeoutSeconds,
            "ConnectionTestTimeoutSeconds",
            minValue: 5,
            maxValue: 300);
        if (!timeoutResult.IsValid)
        {
            errors.Add(timeoutResult.ErrorMessage!);
        }

        // Validar EmailsPerCycle
        var emailsResult = ValidatePositiveInteger(
            config.EmailsPerCycle,
            "EmailsPerCycle",
            minValue: 1,
            maxValue: 100);
        if (!emailsResult.IsValid)
        {
            errors.Add(emailsResult.ErrorMessage!);
        }

        // Validar MaxRetries
        var retriesResult = ValidatePositiveInteger(
            config.MaxRetries,
            "MaxRetries",
            minValue: 0,
            maxValue: 10);
        if (!retriesResult.IsValid)
        {
            errors.Add(retriesResult.ErrorMessage!);
        }

        // Validar LogRetentionDays
        var retentionResult = ValidatePositiveInteger(
            config.LogRetentionDays,
            "LogRetentionDays",
            minValue: 1,
            maxValue: 365);
        if (!retentionResult.IsValid)
        {
            errors.Add(retentionResult.ErrorMessage!);
        }

        if (errors.Any())
        {
            return ValidationResult.Fail(string.Join("; ", errors));
        }

        return ValidationResult.Success();
    }
}

/// <summary>
/// Resultado de validação
/// </summary>
public class ValidationResult
{
    public bool IsValid { get; private set; }
    public string? ErrorMessage { get; private set; }

    private ValidationResult(bool isValid, string? errorMessage = null)
    {
        IsValid = isValid;
        ErrorMessage = errorMessage;
    }

    public static ValidationResult Success() => new(true);
    public static ValidationResult Fail(string errorMessage) => new(false, errorMessage);
}

/// <summary>
/// Classe para encapsular configuração da aplicação
/// </summary>
public class AppConfiguration
{
    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string ClientSecret { get; set; } = string.Empty;
    public string UserEmail { get; set; } = string.Empty;
    public string TempFolder { get; set; } = string.Empty;
    public int ServiceIntervalSeconds { get; set; }
    public int ConnectionTestTimeoutSeconds { get; set; }
    public int EmailsPerCycle { get; set; }
    public int MaxRetries { get; set; }
    public int LogRetentionDays { get; set; }
}