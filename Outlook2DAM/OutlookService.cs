using Microsoft.Graph;
using Microsoft.Graph.Models;
using Serilog;
using System.Configuration;
using System.Timers;
using Timer = System.Timers.Timer;
using Message = Microsoft.Graph.Models.Message;
using Outlook2DAM.Services;
using System.Threading;

namespace Outlook2DAM.Services;

public class OutlookService
{
    private readonly ILogger _logger;
    private readonly Timer _timer;
    private readonly EmailProcessor _emailProcessor;
    private readonly Dictionary<string, string> _processedFolderIds;
    private readonly Dictionary<string, string> _inboxFolderIds;
    private readonly Dictionary<string, string> _inboxFolderConfig; // email -> pasta
    private readonly GraphServiceClient _graphClient;
    private List<string> _userEmail = new();
    private string _processedFolder = "Processados";
    private string _defaultInboxFolder = "Inbox";
    private bool _timerEnabled;
    private bool _isRunning;

    public event Action<int>? UnreadEmailCountChanged;

    public OutlookService(GraphServiceClient graphClient)
    {
        _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        _logger = LoggerService.GetLogger<OutlookService>();
        _timer = new Timer();
        _processedFolderIds = new Dictionary<string, string>();
        _inboxFolderIds = new Dictionary<string, string>();
        _inboxFolderConfig = new Dictionary<string, string>();
        _emailProcessor = new EmailProcessor(graphClient);
        _timerEnabled = false;
        _isRunning = false;
        
        LoadSettings();
        
        _timer.Elapsed += Timer_Elapsed;
    }

    private void LoadSettings()
    {
        // Carregar configurações do App.config
        var userEmail = ConfigurationManager.AppSettings["UserEmail"];
        if (string.IsNullOrEmpty(userEmail))
        {
            throw new InvalidOperationException("UserEmail não configurado");
        }

        _userEmail = userEmail.Split(';')
            .Select(e => e.Trim())
            .Where(e => !string.IsNullOrEmpty(e))
            .ToList();

        if (!_userEmail.Any())
        {
            throw new InvalidOperationException("Pelo menos um email deve ser configurado em UserEmail");
        }

        _processedFolder = ConfigurationManager.AppSettings["ProcessedFolder"] ?? "Processados";

        // Processar configuração de InboxFolder
        var inboxFolderConfig = ConfigurationManager.AppSettings["InboxFolder"] ?? "Inbox";
        ParseInboxFolderConfig(inboxFolderConfig);

        var intervalSeconds = int.Parse(ConfigurationManager.AppSettings["ServiceIntervalSeconds"] ?? "60");
        _timer.Interval = intervalSeconds * 1000;

        // Log da configuração
        var configDetails = _inboxFolderConfig.Any()
            ? string.Join(", ", _inboxFolderConfig.Select(kv => $"{kv.Key}→{kv.Value}"))
            : $"Todos→{_defaultInboxFolder}";

        _logger.Information("Outlook2DAM inicializado. Intervalo: {Interval}s, Pastas entrada: [{InboxConfig}], Pasta processados: {ProcessedFolder}, Emails: {Emails}",
            intervalSeconds, configDetails, _processedFolder, string.Join(", ", _userEmail));
    }

    private void ParseInboxFolderConfig(string config)
    {
        if (string.IsNullOrWhiteSpace(config) || config.Equals("inbox", StringComparison.OrdinalIgnoreCase))
        {
            _defaultInboxFolder = "Inbox";
            return;
        }

        // Verificar se é formato email:pasta
        if (config.Contains(':'))
        {
            // Formato: email1:pasta1;email2:pasta2
            var entries = config.Split(';', StringSplitOptions.RemoveEmptyEntries);
            foreach (var entry in entries)
            {
                var parts = entry.Split(':', 2);
                if (parts.Length == 2)
                {
                    var email = parts[0].Trim().ToLowerInvariant();
                    var folder = parts[1].Trim();

                    if (!string.IsNullOrEmpty(email) && !string.IsNullOrEmpty(folder))
                    {
                        _inboxFolderConfig[email] = folder;
                        _logger.Debug("Configurada pasta '{Folder}' para {Email}", folder, email);
                    }
                }
            }
        }
        else
        {
            // Formato simples: apenas nome da pasta (usa para todos)
            _defaultInboxFolder = config.Trim();
        }
    }

    /// <summary>
    /// Valida se todas as pastas configuradas existem (opcional, para uso no editor de configurações)
    /// </summary>
    public async Task ValidateInboxFolders()
    {
        _logger.Information("╔════════════════════════════════════════════════════════════════");
        _logger.Information("║ Validação de Pastas de Entrada");
        _logger.Information("╠════════════════════════════════════════════════════════════════");

        var allValid = true;

        foreach (var email in _userEmail)
        {
            try
            {
                var emailKey = email.ToLowerInvariant();
                var folderName = _inboxFolderConfig.ContainsKey(emailKey)
                    ? _inboxFolderConfig[emailKey]
                    : _defaultInboxFolder;

                _logger.Information("║ Validando: {Email} → pasta '{FolderName}'", email, folderName);

                // Tentar obter o ID da pasta (isso já valida se existe)
                var folderId = await GetInboxFolderId(email);

                if (!string.IsNullOrEmpty(folderId))
                {
                    _logger.Information("║   ✓ Pasta encontrada e acessível", email, folderName);
                }
            }
            catch (Exception ex)
            {
                allValid = false;
                _logger.Error("║   ✗ ERRO: {Message}", ex.Message);
            }
        }

        _logger.Information("╚════════════════════════════════════════════════════════════════");

        if (!allValid)
        {
            throw new InvalidOperationException("Uma ou mais pastas configuradas não foram encontradas. Verifique os logs acima para detalhes.");
        }

        _logger.Information("✓ Todas as pastas de entrada foram validadas com sucesso!");
    }

    /// <summary>
    /// Lista todas as pastas disponíveis para um email específico
    /// </summary>
    public async Task ListAvailableFolders(string userEmail)
    {
        try
        {
            _logger.Warning("A listar todas as pastas disponíveis para {Email}...", userEmail);

            var allFolders = await _graphClient.Users[userEmail].MailFolders
                .GetAsync(q => q.QueryParameters.Top = 50);

            if (allFolders?.Value != null && allFolders.Value.Any())
            {
                _logger.Warning("╔════════════════════════════════════════════════════════════════");
                _logger.Warning("║ Pastas disponíveis em {Email}:", userEmail);
                _logger.Warning("╠════════════════════════════════════════════════════════════════");

                foreach (var folder in allFolders.Value.OrderBy(f => f.DisplayName))
                {
                    var unreadCount = folder.UnreadItemCount ?? 0;
                    var totalCount = folder.TotalItemCount ?? 0;
                    _logger.Warning("║ • {FolderName} (não lidos: {Unread}, total: {Total})",
                        folder.DisplayName, unreadCount, totalCount);
                }

                _logger.Warning("╚════════════════════════════════════════════════════════════════");
                _logger.Warning("Configure InboxFolder com um dos nomes acima (ex: InboxFolder=\"{Example}\")",
                    allFolders.Value.First().DisplayName);
            }
            else
            {
                _logger.Warning("Nenhuma pasta encontrada para {Email}", userEmail);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao listar pastas disponíveis para {Email}", userEmail);
        }
    }

    private async Task<string> GetInboxFolderId(string userEmail)
    {
        try
        {
            if (_inboxFolderIds.ContainsKey(userEmail))
                return _inboxFolderIds[userEmail];

            // Determinar qual pasta usar para este email
            var emailKey = userEmail.ToLowerInvariant();
            var folderName = _inboxFolderConfig.ContainsKey(emailKey)
                ? _inboxFolderConfig[emailKey]
                : _defaultInboxFolder;

            // Se for "Inbox" padrão, usar wellKnownFolder
            if (folderName.Equals("Inbox", StringComparison.OrdinalIgnoreCase))
            {
                var inbox = await _graphClient.Users[userEmail].MailFolders["inbox"].GetAsync();
                if (inbox?.Id != null)
                {
                    _inboxFolderIds[userEmail] = inbox.Id;
                    _logger.Debug("A usar pasta Inbox padrão para {Email}", userEmail);
                    return inbox.Id;
                }
            }
            else
            {
                // Procurar pasta personalizada
                var folders = await _graphClient.Users[userEmail].MailFolders
                    .GetAsync(q => q.QueryParameters.Filter = $"displayName eq '{folderName}'");

                var customFolder = folders?.Value?.FirstOrDefault();
                if (customFolder?.Id != null)
                {
                    _inboxFolderIds[userEmail] = customFolder.Id;
                    _logger.Information("✓ Pasta personalizada '{FolderName}' encontrada e validada para {Email}", folderName, userEmail);
                    return customFolder.Id;
                }
                else
                {
                    // Pasta não encontrada - listar todas as pastas disponíveis
                    _logger.Error("✗ Pasta '{FolderName}' não encontrada para {Email}", folderName, userEmail);
                    await ListAvailableFolders(userEmail);
                    throw new InvalidOperationException($"Pasta '{folderName}' não encontrada para {userEmail}. Verifique o log para ver as pastas disponíveis.");
                }
            }

            throw new InvalidOperationException($"Não foi possível obter ID da pasta de entrada para {userEmail}");
        }
        catch (Exception ex)
        {
            var emailKey = userEmail.ToLowerInvariant();
            var folderName = _inboxFolderConfig.ContainsKey(emailKey)
                ? _inboxFolderConfig[emailKey]
                : _defaultInboxFolder;
            _logger.Error(ex, "Erro ao obter ID da pasta de entrada '{FolderName}' para {Email}", folderName, userEmail);
            throw;
        }
    }

    private async Task EnsureProcessedFolderExists(string userEmail)
    {
        try
        {
            if (_processedFolderIds.ContainsKey(userEmail))
                return;

            var folders = await _graphClient.Users[userEmail].MailFolders
                .GetAsync(q => q.QueryParameters.Filter = $"displayName eq '{_processedFolder}'");

            var processedFolder = folders?.Value?.FirstOrDefault();
            if (processedFolder == null)
            {
                _logger.Information("A criar pasta: {Folder} para {Email}...", _processedFolder, userEmail);
                var newFolder = new MailFolder
                {
                    DisplayName = _processedFolder,
                };
                processedFolder = await _graphClient.Users[userEmail].MailFolders
                    .PostAsync(newFolder);

                _logger.Information("A pasta {Folder} criada com sucesso para {Email}", _processedFolder, userEmail);
            }
            else
            {
                _logger.Debug("A pasta {Folder} já existe para {Email}", _processedFolder, userEmail);
            }

            if (processedFolder?.Id != null)
            {
                _processedFolderIds[userEmail] = processedFolder.Id;
            }
            else
            {
                throw new InvalidOperationException($"Não foi possível criar/encontrar a pasta {_processedFolder} para {userEmail}");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao criar/verificar pasta {Folder} para {Email}", _processedFolder, userEmail);
            throw;
        }
    }

    private async void Timer_Elapsed(object? sender, ElapsedEventArgs e)
    {
        if (!_timerEnabled)
        {
            return;
        }

        try
        {
            _timer.Stop(); // Para o timer durante o processamento
            await CheckEmails();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro durante o processamento de emails");
        }
        finally
        {
            if (_timerEnabled) // Só reinicia se ainda estiver habilitado
            {
                _timer.Start();
            }
        }
    }

    private async Task CheckEmails()
    {
        if (_isRunning)
        {
            _logger.Debug("Processamento anterior ainda em execução, aguardando próximo ciclo");
            return;
        }

        _isRunning = true;
        try
        {
            var totalUnread = 0;
            foreach (var userEmail in _userEmail)
            {
                try
                {
                    // Verificar quantidade de emails não lidos
                    var unreadCount = await GetUnreadEmailCount(userEmail);
                    totalUnread += unreadCount;
                    
                    if (unreadCount == 0)
                    {
                        _logger.Debug("Nenhum email não lido para processar em {Email}", userEmail);
                        continue;
                    }

                    // Garantir que a pasta processados existe
                    await EnsureProcessedFolderExists(userEmail);

                    var maxEmails = int.TryParse(ConfigurationManager.AppSettings["EmailsPerCycle"], out var max) ? max : 1;
                    _logger.Information("Encontrados {Count} emails não lidos em {Email}. Limite por ciclo: {Max}", 
                        unreadCount, userEmail, maxEmails);

                    var emailsToProcess = Math.Min(unreadCount, maxEmails);
                    for (int i = 0; i < emailsToProcess; i++)
                    {
                        await ProcessNextUnreadEmail(userEmail);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Erro ao processar emails para {Email}", userEmail);
                }
            }
            UnreadEmailCountChanged?.Invoke(totalUnread);
        }
        finally
        {
            _isRunning = false;
        }
    }

    private async Task ProcessNextUnreadEmail(string userEmail)
    {
        try
        {
            // Obter ID da pasta de entrada configurada
            var inboxFolderId = await GetInboxFolderId(userEmail);

            // Verificar se deve processar apenas não lidos
            var processOnlyUnread = bool.Parse(ConfigurationManager.AppSettings["ProcessOnlyUnread"] ?? "true");

            var messages = await _graphClient.Users[userEmail].MailFolders[inboxFolderId].Messages
                .GetAsync(requestConfig =>
                {
                    if (processOnlyUnread)
                    {
                        requestConfig.QueryParameters.Filter = "isRead eq false";
                    }
                    requestConfig.QueryParameters.Orderby = new[] { "receivedDateTime desc" };
                    requestConfig.QueryParameters.Top = 1;
                });

            if (messages?.Value == null || !messages.Value.Any())
            {
                var emailKey = userEmail.ToLowerInvariant();
                var folderName = _inboxFolderConfig.ContainsKey(emailKey) ? _inboxFolderConfig[emailKey] : _defaultInboxFolder;
                var filterText = processOnlyUnread ? "não lido" : "";
                _logger.Debug("Nenhum email {FilterText} encontrado na pasta '{FolderName}' para {Email}", filterText, folderName, userEmail);
                return;
            }

            var message = messages.Value.First();
            await _emailProcessor.ProcessEmail(message, userEmail);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao processar próximo email para {Email}", userEmail);
            throw;
        }
    }

    private async Task<int> GetUnreadEmailCount(string userEmail)
    {
        try
        {
            // Obter ID da pasta de entrada configurada
            var inboxFolderId = await GetInboxFolderId(userEmail);

            // Verificar se deve processar apenas não lidos
            var processOnlyUnread = bool.Parse(ConfigurationManager.AppSettings["ProcessOnlyUnread"] ?? "true");

            _logger.Debug("GetUnreadEmailCount: ProcessOnlyUnread={ProcessOnlyUnread}, FolderId={FolderId}", processOnlyUnread, inboxFolderId);

            var messages = await _graphClient.Users[userEmail].MailFolders[inboxFolderId].Messages
                .GetAsync(requestConfiguration => {
                    if (processOnlyUnread)
                    {
                        requestConfiguration.QueryParameters.Filter = "isRead eq false";
                        _logger.Debug("Aplicado filtro: isRead eq false");
                    }
                    else
                    {
                        _logger.Debug("SEM filtro - buscando todos os emails");
                    }
                    requestConfiguration.QueryParameters.Count = true;
                    requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                    requestConfiguration.QueryParameters.Top = 10; // Pegar alguns emails para debug
                });

            var count = messages?.Value?.Count ?? 0;
            var emailKey = userEmail.ToLowerInvariant();
            var folderName = _inboxFolderConfig.ContainsKey(emailKey) ? _inboxFolderConfig[emailKey] : _defaultInboxFolder;
            var filterText = processOnlyUnread ? "não lidos" : "totais";

            _logger.Information("Encontrados {Count} emails {FilterText} na pasta '{FolderName}' de {Email}", count, filterText, folderName, userEmail);

            // Log detalhado se não encontrar emails
            if (count == 0)
            {
                _logger.Warning("⚠️ ATENÇÃO: 0 emails encontrados!");
                _logger.Warning("  → Pasta: '{FolderName}' (ID: {FolderId})", folderName, inboxFolderId);
                _logger.Warning("  → Filtro: ProcessOnlyUnread={ProcessOnlyUnread}", processOnlyUnread);
                _logger.Warning("  → Verifique se a pasta realmente contém emails {FilterText}", filterText);

                // Tentar buscar sem filtro para confirmar
                if (processOnlyUnread)
                {
                    _logger.Warning("  → Tentando buscar TODOS os emails (sem filtro) para diagnóstico...");
                    var allMessages = await _graphClient.Users[userEmail].MailFolders[inboxFolderId].Messages
                        .GetAsync(rc => {
                            rc.QueryParameters.Count = true;
                            rc.Headers.Add("ConsistencyLevel", "eventual");
                            rc.QueryParameters.Top = 10;
                        });
                    var totalCount = allMessages?.Value?.Count ?? 0;
                    _logger.Warning("  → Total de emails na pasta (sem filtro): {TotalCount}", totalCount);

                    if (totalCount > 0 && allMessages?.Value != null)
                    {
                        _logger.Warning("  → Primeiros emails encontrados:");
                        foreach (var msg in allMessages.Value.Take(3))
                        {
                            _logger.Warning("     • Assunto: '{Subject}', Lido: {IsRead}, Data: {ReceivedDateTime}",
                                msg.Subject ?? "(sem assunto)",
                                msg.IsRead ?? false,
                                msg.ReceivedDateTime);
                        }
                    }
                }
            }

            return count;
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao obter contagem de emails para {Email}", userEmail);
            return 0;
        }
    }

    public void StartService(CancellationToken cancellationToken = default)
    {
        if (_isRunning)
        {
            _logger.Warning("Serviço já está em execução");
            return;
        }

        _logger.Information("Iniciando serviço...");

        try
        {
            // Carregar configurações
            var intervalSeconds = int.TryParse(ConfigurationManager.AppSettings["ServiceIntervalSeconds"], out var interval) ? interval : 60;
            _timer.Interval = intervalSeconds * 1000; // Converter para milissegundos

            _logger.Information("Intervalo do serviço configurado para {Interval} segundos", intervalSeconds);

            // Registrar o callback de cancelamento
            cancellationToken.Register(() =>
            {
                _logger.Information("Solicitação de cancelamento recebida");
                StopService();
            });

            // Iniciar o timer
            _timerEnabled = true;
            _timer.Start();

            // Executar primeira verificação imediatamente
            _ = CheckEmails();

            _logger.Information("Serviço iniciado com sucesso");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao iniciar o serviço");
            throw;
        }
    }

    public void StopService()
    {
        if (!_isRunning)
        {
            _logger.Warning("Serviço já está parado");
            return;
        }

        _logger.Information("Parando serviço...");
        
        try
        {
            _timerEnabled = false;
            _timer.Stop();
            _isRunning = false;
            
            _logger.Information("Serviço parado com sucesso");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao parar o serviço");
            throw;
        }
    }

    public bool IsRunning => _timerEnabled;
}
