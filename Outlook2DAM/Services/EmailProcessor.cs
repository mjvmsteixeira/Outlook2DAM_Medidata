using Microsoft.Graph.Models;
using System.Data.OleDb;
using Serilog;
using System.Configuration;
using System.Text;
using System.Xml.Linq;
using Message = Microsoft.Graph.Models.Message;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Font;
using iText.Layout.Properties;
using iText.Html2pdf;
using iText.IO.Font.Constants;
using Microsoft.Graph;

namespace Outlook2DAM.Services;

public class EmailProcessor
{
    private readonly string _connectionString;
    private readonly string _databaseProvider;
    private readonly string _tempFolder;
    private readonly string _errorFolder;
    private readonly int _maxRetries;
    private readonly ILogger _logger;
    private readonly GraphServiceClient _graphServiceClient;

    public EmailProcessor(GraphServiceClient graphServiceClient)
    {
        _graphServiceClient = graphServiceClient;
        _connectionString = ConfigurationManager.ConnectionStrings["Outlook2DAM"].ConnectionString;

        // Detectar provider do connection string (Oracle, SQL Server, etc)
        _databaseProvider = DetectDatabaseProvider(_connectionString);
        _logger = Log.ForContext<EmailProcessor>();
        _logger.Information("Database provider detectado: {Provider}", _databaseProvider);

        _tempFolder = ConfigurationManager.AppSettings["TempFolder"] ?? throw new InvalidOperationException("TempFolder não configurado");
        _errorFolder = Path.Combine(_tempFolder, ConfigurationManager.AppSettings["ErrorFolder"] ?? "Errors");
        _maxRetries = int.TryParse(ConfigurationManager.AppSettings["MaxRetries"], out var retries) ? retries : 3;
    }

    private string DetectDatabaseProvider(string connectionString)
    {
        var lowerConnectionString = connectionString.ToLowerInvariant();

        if (lowerConnectionString.Contains("provider=oraoledb") || lowerConnectionString.Contains("oracle"))
        {
            return "Oracle";
        }
        else if (lowerConnectionString.Contains("provider=sqloledb") || lowerConnectionString.Contains("provider=sqlncli") ||
                 lowerConnectionString.Contains("provider=msoledbsql") || lowerConnectionString.Contains("sql server"))
        {
            return "SqlServer";
        }
        else if (lowerConnectionString.Contains("provider=microsoft.ace.oledb") || lowerConnectionString.Contains("provider=microsoft.jet.oledb"))
        {
            return "Access";
        }

        // Default para Oracle para compatibilidade
        var sanitizedCs = SensitiveDataFilter.SanitizeConnectionString(connectionString);
        _logger?.Warning("Provider não detectado automaticamente. Usando Oracle como padrão. Connection string: {ConnectionString}",
            sanitizedCs.Substring(0, Math.Min(50, sanitizedCs.Length)) + "...");
        return "Oracle";
    }

    private async Task<bool> ValidateFileCreation(string filePath, int maxRetries = 3, int delayMs = 500)
    {
        for (int i = 0; i < maxRetries; i++)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    // Try to open the file to ensure it's not locked
                    using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        _logger.Debug("Arquivo validado com sucesso: {FilePath}", filePath);
                        return true;
                    }
                }
                catch (IOException)
                {
                    _logger.Warning("Arquivo ainda está bloqueado, tentativa {Attempt} de {MaxRetries}: {FilePath}", 
                        i + 1, maxRetries, filePath);
                }
            }
            else
            {
                _logger.Warning("Arquivo não encontrado, tentativa {Attempt} de {MaxRetries}: {FilePath}", 
                    i + 1, maxRetries, filePath);
            }

            if (i < maxRetries - 1)
            {
                await Task.Delay(delayMs);
            }
        }

        _logger.Error("Falha ao validar arquivo após {MaxRetries} tentativas: {FilePath}", maxRetries, filePath);
        return false;
    }

    private async Task EnsureDirectoryExists(string? directoryPath)
    {
        if (string.IsNullOrWhiteSpace(directoryPath))
        {
            throw new ArgumentNullException(nameof(directoryPath), "O caminho do diretório não pode ser nulo ou vazio");
        }

        try
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
                await Task.Delay(100); // Small delay to ensure directory creation is complete
            }

            if (!Directory.Exists(directoryPath))
            {
                throw new DirectoryNotFoundException($"Não foi possível criar o diretório: {directoryPath}");
            }

            _logger.Debug("Diretório criado/verificado com sucesso: {DirectoryPath}", directoryPath);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao criar/verificar diretório: {DirectoryPath}", directoryPath);
            throw;
        }
    }

    private async Task CreateEmailBodyPdf(Message message, string pdfPath)
    {
        try
        {
            if (string.IsNullOrEmpty(message.Body?.Content))
            {
                _logger.Warning("Email sem conteúdo");
                return;
            }

            var directory = Path.GetDirectoryName(pdfPath);
            if (directory != null)
            {
                await EnsureDirectoryExists(directory);
            }

            using var writer = new PdfWriter(pdfPath);
            using var pdf = new PdfDocument(writer);
            using var document = new Document(pdf);

            // Configurar fonte padrão
            var font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
            document.SetFont(font);

            if (message.Body.ContentType == Microsoft.Graph.Models.BodyType.Html)
            {
                // Preparar o HTML para conversão
                var htmlContent = message.Body.Content;
                
                // Adicionar CSS básico para melhorar a formatação
                var styledHtml = $@"
                    <html>
                    <head>
                        <style>
                            body {{ font-family: Arial, sans-serif; font-size: 11pt; }}
                            p {{ margin: 0 0 10px 0; }}
                        </style>
                    </head>
                    <body>
                        {htmlContent}
                    </body>
                    </html>";

                using var htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(styledHtml));
                var converterProperties = new ConverterProperties();
                HtmlConverter.ConvertToPdf(htmlStream, pdf, converterProperties);
            }
            else
            {
                // Para conteúdo em texto simples
                document.Add(new Paragraph(message.Body.Content)
                    .SetFontSize(11));
            }

            if (await ValidateFileCreation(pdfPath))
            {
                _logger.Debug("PDF do corpo do email criado e validado em: {PdfPath}", pdfPath);
            }
            else
            {
                throw new IOException($"Falha ao criar ou validar o arquivo PDF: {pdfPath}");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao criar PDF do corpo do email: {PdfPath}", pdfPath);
            throw;
        }
    }

    public async Task SaveEmlFile(Message message, Stream? mimeStream)
    {
        if (message == null) throw new ArgumentNullException(nameof(message));
        if (mimeStream == null) return;

        try
        {
            var emailId = message.Id ?? throw new ArgumentException("Email ID não pode ser nulo", nameof(message));
            var timestamp = message.ReceivedDateTime?.DateTime.ToString("yyMMddHHmm") ?? DateTime.Now.ToString("yyMMddHHmm");
            var shortId = $"{(emailId.Length > 8 ? emailId.Substring(0, 8) : emailId)}_{timestamp}";
            var emailFolder = Path.Combine(_tempFolder, shortId);
            
            await EnsureDirectoryExists(emailFolder);

            var emlPath = Path.Combine(emailFolder, $"{shortId}.eml");
            using var fileStream = File.Create(emlPath);
            await mimeStream.CopyToAsync(fileStream);
            
            if (await ValidateFileCreation(emlPath))
            {
                _logger.Debug("EML (MIME original) criado e validado em: {EmlPath}", emlPath);
            }
            else
            {
                throw new IOException($"Falha ao criar ou validar o arquivo EML: {emlPath}");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao salvar arquivo EML");
            throw;
        }
    }

    private async Task<string> CreateXmlFile(Message message, string xmlPath, string pdfPath, List<string> anexos, string userEmail)
    {
        try
        {
            var directory = Path.GetDirectoryName(xmlPath);
            if (directory == null)
            {
                throw new ArgumentException("Caminho XML inválido", nameof(xmlPath));
            }

            await EnsureDirectoryExists(directory);

            var receivedTime = message.ReceivedDateTime?.DateTime ?? DateTime.Now;
            var totalSeconds = receivedTime.Hour * 3600 + receivedTime.Minute * 60 + receivedTime.Second;

            var directoryPath = Path.GetDirectoryName(xmlPath);
            if (string.IsNullOrEmpty(directoryPath))
            {
                throw new ArgumentException("Caminho do diretório inválido", nameof(xmlPath));
            }

            // Garantir que o caminho use \\ para separadores
            var normalizedPath = directoryPath.Replace("\\", "\\\\");
            if (!normalizedPath.EndsWith("\\\\"))
            {
                normalizedPath += "\\\\";
            }

            // Criar lista de anexos incluindo o EML se existir
            var allAttachments = new List<string>();
            var emlFile = Path.ChangeExtension(Path.GetFileName(xmlPath), ".eml");
            if (File.Exists(Path.Combine(directoryPath, emlFile)))
            {
                allAttachments.Add(emlFile);
            }
            allAttachments.AddRange(anexos);

            // Filtrar destinatários: manter apenas os emails configurados em UserEmail
            var configuredEmails = ConfigurationManager.AppSettings["UserEmail"]?
                .Split(';')
                .Select(e => e.Trim().ToLowerInvariant())
                .Where(e => !string.IsNullOrEmpty(e))
                .ToHashSet() ?? new HashSet<string>();

            var filteredRecipients = message.ToRecipients?
                .Select(r => r.EmailAddress?.Address)
                .Where(a => !string.IsNullOrEmpty(a) && configuredEmails.Contains(a!.ToLowerInvariant()))
                .Cast<string>()
                .ToList() ?? new List<string>();

            // Se não houver destinatários filtrados, usar o userEmail do processamento
            var toValue = filteredRecipients.Any()
                ? string.Join(";", filteredRecipients)
                : userEmail;

            var xml = new XDocument(
                new XElement("correspondencia",
                    new XElement("via", "E"),
                    new XElement("data", receivedTime.ToString("yyyy-MM-ddTHH:mm:ss")),
                    new XElement("hora", totalSeconds),
                    new XElement("assunto", message.Subject ?? "Sem assunto"),
                    new XElement("pasta", normalizedPath),
                    new XElement("ficheiro", Path.GetFileName(pdfPath)),
                    new XElement("anexos",
                        allAttachments.Select(a => new XElement("anexo", a))
                    ),
                    new XElement("from", message.From?.EmailAddress?.Address ?? ""),
                    new XElement("to", toValue),
                    new XElement("ver", "0")
                )
            );

            _logger.Debug("Gerando XML com os seguintes dados:");
            _logger.Debug("Data: {Data}", receivedTime.ToString("yyyy-MM-ddTHH:mm:ss"));
            _logger.Debug("Hora: {Hora}", totalSeconds);
            _logger.Debug("Assunto: {Assunto}", message.Subject);
            _logger.Debug("Pasta: {Pasta}", normalizedPath);
            _logger.Debug("Ficheiro: {Ficheiro}", Path.GetFileName(pdfPath));
            _logger.Debug("Anexos: {Anexos}", string.Join(", ", allAttachments));
            _logger.Debug("From: {From}", message.From?.EmailAddress?.Address);
            _logger.Debug("To (filtrado): {To}", toValue);
            _logger.Debug("Destinatários originais: {OriginalTo}", string.Join(";", message.ToRecipients?.Select(r => r.EmailAddress?.Address ?? "") ?? Array.Empty<string>()));

            await File.WriteAllTextAsync(xmlPath, xml.ToString(SaveOptions.None));

            if (await ValidateFileCreation(xmlPath))
            {
                var xmlContent = await File.ReadAllTextAsync(xmlPath);
                _logger.Debug("XML criado com sucesso. Conteúdo:\n{XmlContent}", xmlContent);
                return xmlPath;
            }
            else
            {
                throw new IOException($"Falha ao criar ou validar o arquivo XML: {xmlPath}");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao criar arquivo XML");
            throw;
        }
    }

    private async Task InsertEmailRecord(Message message, string xmlPath)
    {
        var receivedTime = message.ReceivedDateTime?.DateTime ?? DateTime.Now;
        var totalSeconds = receivedTime.Hour * 3600 + receivedTime.Minute * 60 + receivedTime.Second;
        var recipients = string.Join(";", message.ToRecipients?.Select(r => r.EmailAddress?.Address) ?? Array.Empty<string>());

        using var connection = new OleDbConnection(_connectionString);
        await connection.OpenAsync();

        var sql = @"
            INSERT INTO outlook (
                chave,
                remetente,
                data,
                hora,
                destinatario,
                assunto,
                caminho_ficheiro,
                processado,
                tipodoc,
                chavedoc,
                observacoes
            ) VALUES (
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?,
                ?
            )";

        using var cmd = new OleDbCommand(sql, connection);
        cmd.Parameters.Add(new OleDbParameter("chave", message.Id));
        cmd.Parameters.Add(new OleDbParameter("remetente", message.From?.EmailAddress?.Address));
        cmd.Parameters.Add(new OleDbParameter("data", receivedTime.Date));
        cmd.Parameters.Add(new OleDbParameter("hora", totalSeconds));
        cmd.Parameters.Add(new OleDbParameter("destinatario", recipients));
        cmd.Parameters.Add(new OleDbParameter("assunto", message.Subject));
        cmd.Parameters.Add(new OleDbParameter("caminho_ficheiro", xmlPath));
        cmd.Parameters.Add(new OleDbParameter("processado", "0"));
        cmd.Parameters.Add(new OleDbParameter("tipodoc", ""));
        cmd.Parameters.Add(new OleDbParameter("chavedoc", ""));
        cmd.Parameters.Add(new OleDbParameter("observacoes", ""));

        await cmd.ExecuteNonQueryAsync();
        _logger.Debug("Email inserida Base de Dados com sucesso!");
    }

    public async Task ProcessEmail(Message message, string userEmail)
    {
        if (message == null)
            throw new ArgumentNullException(nameof(message));

        var retryCount = 0;
        while (retryCount < _maxRetries)
        {
            try
            {
                _logger.Information("Processando email: {Subject} de {From} (Tentativa {Attempt}/{MaxRetries})", 
                    message.Subject, 
                    message.From?.EmailAddress?.Address,
                    retryCount + 1,
                    _maxRetries);

                var timestamp = message.ReceivedDateTime?.DateTime ?? DateTime.Now;
                var emailFolder = Path.Combine(
                    _tempFolder,
                    timestamp.ToString("yyyyMMdd_HHmmss") + "_" + (message.Id?.Substring(0, 8) ?? "unknown")
                );

                // Criar pasta única para o email
                await EnsureDirectoryExists(emailFolder);

                // Salvar EML se configurado
                string? emlPath = null;
                var saveMimeContent = bool.TryParse(ConfigurationManager.AppSettings["SaveMimeContent"], out var saveSetting) && saveSetting;
                if (saveMimeContent && !string.IsNullOrEmpty(message.Id))
                {
                    _logger.Debug("SaveMimeContent está ativado, salvando EML...");
                    emlPath = Path.Combine(emailFolder, "email.eml");
                    await SaveEmailAsEml(userEmail, message.Id, emlPath);
                }

                // Processar anexos
                var anexos = await ProcessarAnexos(message, emailFolder, userEmail);

                // Gerar PDF do corpo do email
                var pdfPath = Path.Combine(emailFolder, "email.pdf");
                await CreateEmailBodyPdf(message, pdfPath);

                // Gerar XML
                var xmlPath = Path.Combine(emailFolder, "email.xml");
                await CreateXmlFile(message, xmlPath, pdfPath, anexos, userEmail);

                // Salvar no banco de dados
                await SaveToDatabase(message, xmlPath, message.Id ?? Guid.NewGuid().ToString(), 
                    timestamp.Hour * 3600 + timestamp.Minute * 60 + timestamp.Second);

                // Mover para pasta processados
                await MoveToProcessedFolder(userEmail, message.Id);

                _logger.Information("Email processado com sucesso: {Subject}", message.Subject);
                return; // Sucesso, sair do loop
            }
            catch (Exception ex)
            {
                retryCount++;
                _logger.Error(ex, "Erro ao processar email (Tentativa {Attempt}/{MaxRetries}): {Subject}", 
                    retryCount, _maxRetries, message.Subject);

                if (retryCount >= _maxRetries)
                {
                    _logger.Error("Número máximo de tentativas atingido. Movendo email para pasta de erros.");
                    await MoveToErrorFolder(userEmail, message.Id);
                    throw; // Propagar o erro após mover para pasta de erros
                }

                await Task.Delay(500 * retryCount); // Delay progressivo entre tentativas
            }
        }
    }

    private async Task<List<string>> ProcessarAnexos(Message message, string emailFolder, string userEmail)
    {
        if (string.IsNullOrEmpty(emailFolder))
        {
            throw new ArgumentException("Pasta do email não pode ser nula ou vazia", nameof(emailFolder));
        }

        var attachmentNames = new List<string>();

        if (message.HasAttachments == true && !string.IsNullOrEmpty(message.Id))
        {
            _logger.Debug("Email tem anexos, processando...");
            
            try
            {
                _logger.Debug("Obtendo anexos do email {EmailId}", message.Id);
                var attachments = await _graphServiceClient.Users[userEmail]
                    .Messages[message.Id]
                    .Attachments
                    .GetAsync();

                if (attachments?.Value != null)
                {
                    foreach (var attachment in attachments.Value)
                    {
                        if (attachment is FileAttachment fileAttachment && !string.IsNullOrEmpty(attachment.Name))
                        {
                            try 
                            {
                                var attachmentPath = Path.Combine(emailFolder, attachment.Name);
                                _logger.Debug("Baixando conteúdo do anexo: {Name}", attachment.Name);
                                
                                if (!string.IsNullOrEmpty(attachment.Id))
                                {
                                    // Fazer uma chamada separada para obter o conteúdo do anexo
                                    var attachmentContent = await _graphServiceClient.Users[userEmail]
                                        .Messages[message.Id]
                                        .Attachments[attachment.Id]
                                        .GetAsync();

                                    if (attachmentContent is FileAttachment fullAttachment && 
                                        fullAttachment.ContentBytes != null)
                                    {
                                        await File.WriteAllBytesAsync(attachmentPath, fullAttachment.ContentBytes);
                                        
                                        if (await ValidateFileCreation(attachmentPath))
                                        {
                                            attachmentNames.Add(attachment.Name);
                                            _logger.Information("Anexo salvo com sucesso: {Name}", attachment.Name);
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.Error(ex, "Erro ao processar anexo: {Name}", attachment.Name);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Erro ao processar anexos do email {EmailId}", message.Id);
            }
        }
        else
        {
            _logger.Debug("Email não tem anexos");
        }

        return attachmentNames;
    }

    private async Task SaveToDatabase(Message message, string xmlPath, string chave, int hora)
    {
        var receivedTime = message.ReceivedDateTime?.DateTime ?? DateTime.Now;
        var totalSeconds = receivedTime.Hour * 3600 + receivedTime.Minute * 60 + receivedTime.Second;

        // Filtrar destinatários: manter apenas os emails configurados em UserEmail
        var configuredEmails = ConfigurationManager.AppSettings["UserEmail"]?
            .Split(';')
            .Select(e => e.Trim().ToLowerInvariant())
            .Where(e => !string.IsNullOrEmpty(e))
            .ToHashSet() ?? new HashSet<string>();

        var filteredRecipients = message.ToRecipients?
            .Select(r => r.EmailAddress?.Address)
            .Where(a => !string.IsNullOrEmpty(a) && configuredEmails.Contains(a!.ToLowerInvariant()))
            .Cast<string>()
            .ToList() ?? new List<string>();

        var recipients = filteredRecipients.Any() ? string.Join(";", filteredRecipients) : string.Empty;

        using var connection = new OleDbConnection(_connectionString);
        await connection.OpenAsync();

        // Adaptar SQL para diferentes providers
        string sql;
        if (_databaseProvider == "SqlServer")
        {
            // SQL Server usa tipos de dados específicos
            sql = @"
                INSERT INTO outlook (
                    chave,
                    remetente,
                    data,
                    hora,
                    destinatario,
                    assunto,
                    caminho_ficheiro,
                    processado,
                    tipodoc,
                    chavedoc,
                    observacoes
                ) VALUES (
                    ?,
                    ?,
                    CONVERT(DATE, ?),
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?
                )";
        }
        else
        {
            // Oracle e outros
            sql = @"
                INSERT INTO outlook (
                    chave,
                    remetente,
                    data,
                    hora,
                    destinatario,
                    assunto,
                    caminho_ficheiro,
                    processado,
                    tipodoc,
                    chavedoc,
                    observacoes
                ) VALUES (
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?,
                    ?
                )";
        }

        using var cmd = new OleDbCommand(sql, connection);
        cmd.Parameters.Add(new OleDbParameter("chave", message.Id));
        cmd.Parameters.Add(new OleDbParameter("remetente", message.From?.EmailAddress?.Address ?? string.Empty));
        cmd.Parameters.Add(new OleDbParameter("data", receivedTime.Date));
        cmd.Parameters.Add(new OleDbParameter("hora", totalSeconds));
        cmd.Parameters.Add(new OleDbParameter("destinatario", recipients));
        cmd.Parameters.Add(new OleDbParameter("assunto", message.Subject ?? string.Empty));
        cmd.Parameters.Add(new OleDbParameter("caminho_ficheiro", xmlPath));
        cmd.Parameters.Add(new OleDbParameter("processado", "0"));
        cmd.Parameters.Add(new OleDbParameter("tipodoc", string.Empty));
        cmd.Parameters.Add(new OleDbParameter("chavedoc", string.Empty));
        cmd.Parameters.Add(new OleDbParameter("observacoes", string.Empty));

        await cmd.ExecuteNonQueryAsync();
        _logger.Information("Email inserido na Base de Dados ({Provider}) com sucesso! Destinatários filtrados: {Recipients}",
            _databaseProvider, recipients);
    }

    private async Task MoveToProcessedFolder(string userEmail, string? messageId)
    {
        if (string.IsNullOrEmpty(messageId))
        {
            _logger.Warning("ID da mensagem está vazio, não é possível mover");
            return;
        }

        try
        {
            var processedFolder = ConfigurationManager.AppSettings["ProcessedFolder"] ?? "Processados";
            
            // Primeiro, marcar como lido
            await _graphServiceClient.Users[userEmail].Messages[messageId]
                .PatchAsync(new Message { IsRead = true });

            // Verificar/criar pasta processados
            var folders = await _graphServiceClient.Users[userEmail].MailFolders
                .GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.Filter = $"displayName eq '{processedFolder}'";
                });

            string folderId;
            if (folders?.Value == null || !folders.Value.Any())
            {
                // Criar pasta se não existir
                var newFolder = await _graphServiceClient.Users[userEmail].MailFolders
                    .PostAsync(new MailFolder
                    {
                        DisplayName = processedFolder,
                        IsHidden = false
                    });

                folderId = newFolder?.Id ?? throw new InvalidOperationException($"Falha ao criar pasta {processedFolder}");
            }
            else
            {
                folderId = folders.Value.First().Id ?? throw new InvalidOperationException($"ID da pasta {processedFolder} é nulo");
            }

            // Mover a mensagem
            await _graphServiceClient.Users[userEmail].Messages[messageId].Move
                .PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                {
                    DestinationId = folderId
                });

            _logger.Information("Email movido para pasta {Folder}", processedFolder);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao mover email para pasta processados");
            throw;
        }
    }

    private async Task SaveEmailAsEml(string userEmail, string messageId, string emlPath)
    {
        try
        {
            _logger.Debug("Obtendo conteúdo MIME para email {MessageId}...", messageId);
            
            var mimeStream = await _graphServiceClient.Users[userEmail].Messages[messageId].Content
                .GetAsync();

            if (mimeStream == null)
            {
                _logger.Warning("Conteúdo MIME não disponível para {MessageId}", messageId);
                return;
            }

            using (var fileStream = File.Create(emlPath))
            {
                await mimeStream.CopyToAsync(fileStream);
            }

            // Validar se o arquivo foi criado corretamente
            if (await ValidateFileCreation(emlPath))
            {
                _logger.Information("Arquivo EML salvo com sucesso: {Path}", emlPath);
            }
            else
            {
                _logger.Error("Falha ao validar arquivo EML: {Path}", emlPath);
                throw new IOException($"Falha ao validar arquivo EML: {emlPath}");
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao salvar arquivo EML: {Path}", emlPath);
            throw;
        }
    }

    private async Task MoveToErrorFolder(string userEmail, string? messageId)
    {
        if (string.IsNullOrEmpty(messageId))
        {
            _logger.Warning("ID da mensagem está vazio, não é possível mover para pasta de erros");
            return;
        }

        try
        {
            // Primeiro, marcar como lido
            await _graphServiceClient.Users[userEmail].Messages[messageId]
                .PatchAsync(new Message { IsRead = true });

            // Verificar/criar pasta de erros
            var folders = await _graphServiceClient.Users[userEmail].MailFolders
                .GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.Filter = $"displayName eq 'Errors'";
                });

            string folderId;
            if (folders?.Value == null || !folders.Value.Any())
            {
                // Criar pasta se não existir
                var newFolder = await _graphServiceClient.Users[userEmail].MailFolders
                    .PostAsync(new MailFolder
                    {
                        DisplayName = "Errors",
                        IsHidden = false
                    });

                folderId = newFolder?.Id ?? throw new InvalidOperationException("Falha ao criar pasta de erros");
                _logger.Information("Pasta de erros criada com sucesso");
            }
            else
            {
                var firstFolder = folders.Value.First();
                folderId = firstFolder.Id ?? throw new InvalidOperationException("ID da pasta de erros é nulo");
            }

            // Mover a mensagem
            await _graphServiceClient.Users[userEmail].Messages[messageId].Move
                .PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                {
                    DestinationId = folderId
                });

            _logger.Information("Email movido para pasta de erros");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao mover email para pasta de erros");
            // Não propagar o erro para não entrar em loop infinito
        }
    }
}
