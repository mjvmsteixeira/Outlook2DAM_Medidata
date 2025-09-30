using Microsoft.Graph;
using Serilog;
using System.Configuration;
using System.Xml.Linq;

namespace Outlook2DAM;

public class ConfigEditorForm : Form
{
    private readonly ILogger _logger;
    private readonly GraphServiceClient _graphClient;

    private TabControl tabControl = null!;
    private Button btnSave = null!;
    private Button btnCancel = null!;

    // Azure AD
    private TextBox txtTenantId = null!;
    private TextBox txtClientId = null!;
    private TextBox txtClientSecret = null!;

    // Email Configuration - Nova estrutura
    private ListBox lstEmails = null!;
    private TextBox txtNewEmail = null!;
    private ComboBox cboEmailFolder = null!;
    private Button btnAddEmail = null!;
    private Button btnRemoveEmail = null!;
    private Button btnLoadFolders = null!;
    private Button btnTestFolder = null!;
    private Label lblSelectedEmail = null!;
    private Label lblFolderStats = null!;
    private Dictionary<string, string> emailInboxFolders = new(); // email -> pasta

    // Service Settings
    private NumericUpDown numServiceInterval = null!;
    private NumericUpDown numEmailsPerCycle = null!;
    private NumericUpDown numMaxRetries = null!;
    private NumericUpDown numConnectionTimeout = null!;
    private CheckBox chkProcessOnlyUnread = null!;

    // Folders
    private TextBox txtTempFolder = null!;
    private TextBox txtProcessedFolder = null!;
    private TextBox txtErrorFolder = null!;
    private CheckBox chkSaveMimeContent = null!;

    // Logs
    private ComboBox cboLogLevel = null!;
    private TextBox txtLogPath = null!;
    private NumericUpDown numLogRetention = null!;
    private CheckBox chkRewriteLog = null!;

    // Database
    private TextBox txtConnectionString = null!;

    public ConfigEditorForm(GraphServiceClient graphClient)
    {
        _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        _logger = LoggerService.GetLogger<ConfigEditorForm>();

        InitializeComponents();
        LoadConfiguration();
    }

    private void InitializeComponents()
    {
        Text = "Editor de Configurações - Outlook2DAM";
        Size = new Size(800, 650);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        tabControl = new TabControl
        {
            Location = new Point(10, 10),
            Size = new Size(760, 540),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };

        // Create tabs
        CreateAzureTab();
        CreateEmailTab();
        CreateServiceTab();
        CreateFoldersTab();
        CreateLogsTab();
        CreateDatabaseTab();

        Controls.Add(tabControl);

        // Buttons
        btnSave = new Button
        {
            Text = "💾 Guardar",
            Location = new Point(570, 560),
            Size = new Size(90, 30),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        btnSave.Click += BtnSave_Click;
        Controls.Add(btnSave);

        btnCancel = new Button
        {
            Text = "Cancelar",
            Location = new Point(670, 560),
            Size = new Size(90, 30),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            DialogResult = DialogResult.Cancel
        };
        Controls.Add(btnCancel);

        CancelButton = btnCancel;
    }

    private void CreateAzureTab()
    {
        var tab = new TabPage("Azure AD");
        tabControl.TabPages.Add(tab);

        var y = 20;

        // TenantId
        tab.Controls.Add(new Label { Text = "Tenant ID:", Location = new Point(20, y), AutoSize = true });
        txtTenantId = new TextBox { Location = new Point(20, y + 20), Size = new Size(700, 25) };
        tab.Controls.Add(txtTenantId);
        y += 60;

        // ClientId
        tab.Controls.Add(new Label { Text = "Client ID:", Location = new Point(20, y), AutoSize = true });
        txtClientId = new TextBox { Location = new Point(20, y + 20), Size = new Size(700, 25) };
        tab.Controls.Add(txtClientId);
        y += 60;

        // ClientSecret
        tab.Controls.Add(new Label { Text = "Client Secret:", Location = new Point(20, y), AutoSize = true });
        txtClientSecret = new TextBox { Location = new Point(20, y + 20), Size = new Size(700, 25), UseSystemPasswordChar = true };
        tab.Controls.Add(txtClientSecret);
        y += 60;

        var lblInfo = new Label
        {
            Text = "ℹ️ Obtenha estas credenciais no Azure Portal:\nhttps://portal.azure.com → Azure Active Directory → App registrations",
            Location = new Point(20, y),
            Size = new Size(700, 60),
            ForeColor = Color.Blue
        };
        tab.Controls.Add(lblInfo);
    }

    private void CreateEmailTab()
    {
        var tab = new TabPage("📧 Emails & Pastas");
        tabControl.TabPages.Add(tab);

        // Left side - Email list
        var lblEmails = new Label
        {
            Text = "Emails configurados:",
            Location = new Point(20, 20),
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        };
        tab.Controls.Add(lblEmails);

        lstEmails = new ListBox
        {
            Location = new Point(20, 45),
            Size = new Size(340, 200),
            SelectionMode = SelectionMode.One
        };
        lstEmails.SelectedIndexChanged += LstEmails_SelectedIndexChanged;
        tab.Controls.Add(lstEmails);

        // Add email section
        var lblAddEmail = new Label
        {
            Text = "Adicionar novo email:",
            Location = new Point(20, 255),
            AutoSize = true
        };
        tab.Controls.Add(lblAddEmail);

        txtNewEmail = new TextBox
        {
            Location = new Point(20, 275),
            Size = new Size(260, 25),
            PlaceholderText = "email@dominio.com"
        };
        tab.Controls.Add(txtNewEmail);

        btnAddEmail = new Button
        {
            Text = "➕ Adicionar",
            Location = new Point(290, 275),
            Size = new Size(70, 25)
        };
        btnAddEmail.Click += BtnAddEmail_Click;
        tab.Controls.Add(btnAddEmail);

        btnRemoveEmail = new Button
        {
            Text = "❌ Remover",
            Location = new Point(20, 310),
            Size = new Size(100, 25)
        };
        btnRemoveEmail.Click += BtnRemoveEmail_Click;
        tab.Controls.Add(btnRemoveEmail);

        // Right side - Folder configuration
        var lblFolderConfig = new Label
        {
            Text = "Pasta de entrada para o email selecionado:",
            Location = new Point(380, 20),
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        };
        tab.Controls.Add(lblFolderConfig);

        lblSelectedEmail = new Label
        {
            Text = "(Selecione um email da lista)",
            Location = new Point(380, 45),
            Size = new Size(340, 20),
            ForeColor = Color.Gray,
            Font = new Font(Font, FontStyle.Italic)
        };
        tab.Controls.Add(lblSelectedEmail);

        var lblFolder = new Label
        {
            Text = "Pasta a monitorizar (InboxFolder):",
            Location = new Point(380, 75),
            AutoSize = true
        };
        tab.Controls.Add(lblFolder);

        cboEmailFolder = new ComboBox
        {
            Location = new Point(380, 95),
            Size = new Size(260, 25),
            DropDownStyle = ComboBoxStyle.DropDown
        };
        cboEmailFolder.Items.AddRange(new[] { "Inbox", "Caixa de Entrada", "Processados", "Rascunhos" });
        cboEmailFolder.TextChanged += CboEmailFolder_TextChanged;
        tab.Controls.Add(cboEmailFolder);

        btnLoadFolders = new Button
        {
            Text = "🔄 Listar Pastas",
            Location = new Point(650, 95),
            Size = new Size(70, 25)
        };
        btnLoadFolders.Click += BtnLoadFolders_Click;
        tab.Controls.Add(btnLoadFolders);

        // Info panel
        var panelInfo = new Panel
        {
            Location = new Point(380, 140),
            Size = new Size(340, 195),
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.FromArgb(240, 248, 255)
        };

        var lblInfoTitle = new Label
        {
            Text = "💡 Como funciona:",
            Location = new Point(10, 10),
            AutoSize = true,
            Font = new Font(Font, FontStyle.Bold)
        };
        panelInfo.Controls.Add(lblInfoTitle);

        var lblInfo = new Label
        {
            Text = "1. Adicione um ou mais emails à lista\n\n" +
                   "2. Selecione cada email e defina sua\n   pasta de entrada\n\n" +
                   "3. Clique em 'Listar Pastas' para carregar\n   as pastas disponíveis do Outlook\n\n" +
                   "4. Deixe 'Inbox' para usar a Caixa de\n   Entrada padrão\n\n" +
                   "5. As configurações serão guardadas no\n   formato: email:pasta;email:pasta",
            Location = new Point(10, 35),
            Size = new Size(320, 150),
            ForeColor = Color.DarkBlue
        };
        panelInfo.Controls.Add(lblInfo);

        tab.Controls.Add(panelInfo);

        // Bottom info
        var lblBottomInfo = new Label
        {
            Text = "⚠️ Emails sem pasta definida usarão 'Inbox' por padrão",
            Location = new Point(20, 350),
            Size = new Size(700, 20),
            ForeColor = Color.OrangeRed,
            Font = new Font(Font, FontStyle.Italic)
        };
        tab.Controls.Add(lblBottomInfo);
    }

    private void CreateServiceTab()
    {
        var tab = new TabPage("⚙️ Serviço");
        tabControl.TabPages.Add(tab);

        var y = 20;

        // ServiceIntervalSeconds
        tab.Controls.Add(new Label { Text = "Intervalo entre verificações (segundos):", Location = new Point(20, y), AutoSize = true });
        numServiceInterval = new NumericUpDown { Location = new Point(20, y + 20), Size = new Size(120, 25), Minimum = 10, Maximum = 3600, Value = 60 };
        tab.Controls.Add(numServiceInterval);
        y += 60;

        // EmailsPerCycle
        tab.Controls.Add(new Label { Text = "Emails por ciclo:", Location = new Point(20, y), AutoSize = true });
        numEmailsPerCycle = new NumericUpDown { Location = new Point(20, y + 20), Size = new Size(120, 25), Minimum = 1, Maximum = 100, Value = 1 };
        tab.Controls.Add(numEmailsPerCycle);
        y += 60;

        // MaxRetries
        tab.Controls.Add(new Label { Text = "Máximo de tentativas:", Location = new Point(20, y), AutoSize = true });
        numMaxRetries = new NumericUpDown { Location = new Point(20, y + 20), Size = new Size(120, 25), Minimum = 0, Maximum = 10, Value = 3 };
        tab.Controls.Add(numMaxRetries);
        y += 60;

        // ConnectionTestTimeoutSeconds
        tab.Controls.Add(new Label { Text = "Timeout para testes de conexão (segundos):", Location = new Point(20, y), AutoSize = true });
        numConnectionTimeout = new NumericUpDown { Location = new Point(20, y + 20), Size = new Size(120, 25), Minimum = 5, Maximum = 300, Value = 30 };
        tab.Controls.Add(numConnectionTimeout);
        y += 60;

        // ProcessOnlyUnread
        chkProcessOnlyUnread = new CheckBox
        {
            Text = "Processar apenas emails não lidos",
            Location = new Point(20, y),
            AutoSize = true,
            Checked = true
        };
        tab.Controls.Add(chkProcessOnlyUnread);
        y += 40;

        var lblWarning = new Label
        {
            Text = "⚠️ Desmarcar esta opção irá processar TODOS os emails da pasta,\nincluindo os já lidos (cuidado com reprocessamento!)",
            Location = new Point(20, y),
            Size = new Size(600, 40),
            ForeColor = Color.OrangeRed,
            Font = new Font(Font, FontStyle.Italic)
        };
        tab.Controls.Add(lblWarning);
    }

    private void CreateFoldersTab()
    {
        var tab = new TabPage("📁 Pastas");
        tabControl.TabPages.Add(tab);

        var y = 20;

        // TempFolder
        tab.Controls.Add(new Label { Text = "Pasta temporária (TempFolder):", Location = new Point(20, y), AutoSize = true });
        txtTempFolder = new TextBox { Location = new Point(20, y + 20), Size = new Size(700, 25) };
        tab.Controls.Add(txtTempFolder);
        y += 60;

        // ProcessedFolder
        tab.Controls.Add(new Label { Text = "Pasta de emails processados:", Location = new Point(20, y), AutoSize = true });
        txtProcessedFolder = new TextBox { Location = new Point(20, y + 20), Size = new Size(700, 25) };
        tab.Controls.Add(txtProcessedFolder);
        y += 60;

        // ErrorFolder
        tab.Controls.Add(new Label { Text = "Pasta de emails com erro:", Location = new Point(20, y), AutoSize = true });
        txtErrorFolder = new TextBox { Location = new Point(20, y + 20), Size = new Size(700, 25) };
        tab.Controls.Add(txtErrorFolder);
        y += 60;

        // SaveMimeContent
        chkSaveMimeContent = new CheckBox { Text = "Guardar arquivo .eml original", Location = new Point(20, y), AutoSize = true };
        tab.Controls.Add(chkSaveMimeContent);
    }

    private void CreateLogsTab()
    {
        var tab = new TabPage("📋 Logs");
        tabControl.TabPages.Add(tab);

        var y = 20;

        // LogLevel
        tab.Controls.Add(new Label { Text = "Nível de log:", Location = new Point(20, y), AutoSize = true });
        cboLogLevel = new ComboBox
        {
            Location = new Point(20, y + 20),
            Size = new Size(200, 25),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        cboLogLevel.Items.AddRange(new[] { "Verbose", "Debug", "Information", "Warning", "Error", "Fatal" });
        tab.Controls.Add(cboLogLevel);
        y += 60;

        // LogPath
        tab.Controls.Add(new Label { Text = "Caminho dos logs:", Location = new Point(20, y), AutoSize = true });
        txtLogPath = new TextBox { Location = new Point(20, y + 20), Size = new Size(700, 25) };
        tab.Controls.Add(txtLogPath);
        y += 60;

        // LogRetentionDays
        tab.Controls.Add(new Label { Text = "Dias de retenção de logs:", Location = new Point(20, y), AutoSize = true });
        numLogRetention = new NumericUpDown { Location = new Point(20, y + 20), Size = new Size(120, 25), Minimum = 1, Maximum = 365, Value = 31 };
        tab.Controls.Add(numLogRetention);
        y += 60;

        // RewriteLog
        chkRewriteLog = new CheckBox { Text = "Reutilizar arquivo de log", Location = new Point(20, y), AutoSize = true };
        tab.Controls.Add(chkRewriteLog);
    }

    private void CreateDatabaseTab()
    {
        var tab = new TabPage("💾 Base de Dados");
        tabControl.TabPages.Add(tab);

        var y = 20;

        tab.Controls.Add(new Label { Text = "Connection String:", Location = new Point(20, y), AutoSize = true });
        txtConnectionString = new TextBox
        {
            Location = new Point(20, y + 20),
            Size = new Size(700, 100),
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };
        tab.Controls.Add(txtConnectionString);
        y += 140;

        var lblInfo = new Label
        {
            Text = "💡 Exemplos:\n\n" +
                   "SQL Server:\nProvider=SQLOLEDB;Data Source=servidor;Initial Catalog=bd;User ID=user;Password=pwd;\n\n" +
                   "Oracle:\nProvider=OraOLEDB.Oracle;Data Source=tns;User ID=user;Password=pwd;",
            Location = new Point(20, y),
            Size = new Size(700, 150),
            ForeColor = Color.DarkBlue
        };
        tab.Controls.Add(lblInfo);
    }

    private void LoadConfiguration()
    {
        try
        {
            var config = ConfigurationManager.AppSettings;

            txtTenantId.Text = config["TenantId"] ?? "";
            txtClientId.Text = config["ClientId"] ?? "";
            txtClientSecret.Text = config["ClientSecret"] ?? "";

            // Parse UserEmail and InboxFolder
            var userEmails = config["UserEmail"] ?? "";
            var inboxFolderConfig = config["InboxFolder"] ?? "Inbox";

            ParseEmailConfiguration(userEmails, inboxFolderConfig);

            numServiceInterval.Value = int.Parse(config["ServiceIntervalSeconds"] ?? "60");
            numEmailsPerCycle.Value = int.Parse(config["EmailsPerCycle"] ?? "1");
            numMaxRetries.Value = int.Parse(config["MaxRetries"] ?? "3");
            numConnectionTimeout.Value = int.Parse(config["ConnectionTestTimeoutSeconds"] ?? "30");
            chkProcessOnlyUnread.Checked = bool.Parse(config["ProcessOnlyUnread"] ?? "true");

            txtTempFolder.Text = config["TempFolder"] ?? "";
            txtProcessedFolder.Text = config["ProcessedFolder"] ?? "Processados";
            txtErrorFolder.Text = config["ErrorFolder"] ?? "Errors";
            chkSaveMimeContent.Checked = bool.Parse(config["SaveMimeContent"] ?? "true");

            cboLogLevel.SelectedItem = config["LogLevel"] ?? "Information";
            txtLogPath.Text = config["LogPath"] ?? "logs";
            numLogRetention.Value = int.Parse(config["LogRetentionDays"] ?? "31");
            chkRewriteLog.Checked = bool.Parse(config["RewriteLog"] ?? "false");

            txtConnectionString.Text = ConfigurationManager.ConnectionStrings["Outlook2DAM"]?.ConnectionString ?? "";
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao carregar configuração");
            MessageBox.Show($"Erro ao carregar configuração: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ParseEmailConfiguration(string userEmails, string inboxFolderConfig)
    {
        emailInboxFolders.Clear();
        lstEmails.Items.Clear();

        var emails = userEmails.Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(e => e.Trim())
            .Where(e => !string.IsNullOrEmpty(e))
            .ToList();

        // Parse InboxFolder configuration
        if (inboxFolderConfig.Contains(':'))
        {
            // Format: email:folder;email:folder
            var entries = inboxFolderConfig.Split(';', StringSplitOptions.RemoveEmptyEntries);
            foreach (var entry in entries)
            {
                var parts = entry.Split(':', 2);
                if (parts.Length == 2)
                {
                    var email = parts[0].Trim().ToLowerInvariant();
                    var folder = parts[1].Trim();
                    emailInboxFolders[email] = folder;
                }
            }
        }
        else if (!string.IsNullOrWhiteSpace(inboxFolderConfig))
        {
            // Simple format: single folder for all
            var defaultFolder = inboxFolderConfig.Trim();
            foreach (var email in emails)
            {
                emailInboxFolders[email.ToLowerInvariant()] = defaultFolder;
            }
        }

        // Add emails to list
        foreach (var email in emails)
        {
            var emailKey = email.ToLowerInvariant();
            var folder = emailInboxFolders.ContainsKey(emailKey) ? emailInboxFolders[emailKey] : "Inbox";
            lstEmails.Items.Add($"{email} → {folder}");
        }
    }

    private void LstEmails_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (lstEmails.SelectedIndex >= 0)
        {
            var selectedText = lstEmails.SelectedItem?.ToString() ?? "";
            var email = selectedText.Split('→')[0].Trim();
            var emailKey = email.ToLowerInvariant();

            lblSelectedEmail.Text = $"Configurando: {email}";
            lblSelectedEmail.ForeColor = Color.Black;
            lblSelectedEmail.Font = new Font(lblSelectedEmail.Font, FontStyle.Bold);

            cboEmailFolder.Text = emailInboxFolders.ContainsKey(emailKey) ? emailInboxFolders[emailKey] : "Inbox";
            cboEmailFolder.Enabled = true;
            btnLoadFolders.Enabled = true;
        }
        else
        {
            lblSelectedEmail.Text = "(Selecione um email da lista)";
            lblSelectedEmail.ForeColor = Color.Gray;
            lblSelectedEmail.Font = new Font(lblSelectedEmail.Font, FontStyle.Italic);
            cboEmailFolder.Enabled = false;
            btnLoadFolders.Enabled = false;
        }
    }

    private void BtnAddEmail_Click(object? sender, EventArgs e)
    {
        var email = txtNewEmail.Text.Trim();
        if (string.IsNullOrEmpty(email))
        {
            MessageBox.Show("Digite um endereço de email.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Simple email validation
        if (!email.Contains('@') || !email.Contains('.'))
        {
            MessageBox.Show("Email inválido.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var emailKey = email.ToLowerInvariant();
        if (emailInboxFolders.ContainsKey(emailKey))
        {
            MessageBox.Show("Este email já está na lista.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        emailInboxFolders[emailKey] = "Inbox";
        lstEmails.Items.Add($"{email} → Inbox");
        txtNewEmail.Clear();
    }

    private void BtnRemoveEmail_Click(object? sender, EventArgs e)
    {
        if (lstEmails.SelectedIndex >= 0)
        {
            var selectedText = lstEmails.SelectedItem?.ToString() ?? "";
            var email = selectedText.Split('→')[0].Trim();
            var emailKey = email.ToLowerInvariant();

            emailInboxFolders.Remove(emailKey);
            lstEmails.Items.RemoveAt(lstEmails.SelectedIndex);
        }
    }

    private void CboEmailFolder_TextChanged(object? sender, EventArgs e)
    {
        if (lstEmails.SelectedIndex >= 0)
        {
            var selectedText = lstEmails.SelectedItem?.ToString() ?? "";
            var email = selectedText.Split('→')[0].Trim();
            var emailKey = email.ToLowerInvariant();

            emailInboxFolders[emailKey] = cboEmailFolder.Text;
            lstEmails.Items[lstEmails.SelectedIndex] = $"{email} → {cboEmailFolder.Text}";
        }
    }

    private async void BtnLoadFolders_Click(object? sender, EventArgs e)
    {
        if (lstEmails.SelectedIndex < 0)
            return;

        try
        {
            btnLoadFolders.Enabled = false;
            btnLoadFolders.Text = "⏳";

            var selectedText = lstEmails.SelectedItem?.ToString() ?? "";
            var email = selectedText.Split('→')[0].Trim();

            var folders = await _graphClient.Users[email].MailFolders
                .GetAsync(q => q.QueryParameters.Top = 50);

            if (folders?.Value != null && folders.Value.Any())
            {
                cboEmailFolder.Items.Clear();
                foreach (var folder in folders.Value.OrderBy(f => f.DisplayName))
                {
                    if (!string.IsNullOrEmpty(folder.DisplayName))
                    {
                        cboEmailFolder.Items.Add(folder.DisplayName);
                    }
                }

                MessageBox.Show($"Carregadas {folders.Value.Count} pastas de {email}", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao carregar pastas");
            MessageBox.Show($"Erro ao carregar pastas: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            btnLoadFolders.Enabled = true;
            btnLoadFolders.Text = "🔄";
        }
    }

    private void BtnSave_Click(object? sender, EventArgs e)
    {
        try
        {
            // Build UserEmail and InboxFolder strings
            var userEmailList = new List<string>();
            var inboxFolderParts = new List<string>();

            foreach (var kvp in emailInboxFolders)
            {
                userEmailList.Add(kvp.Key);
                if (kvp.Value != "Inbox")
                {
                    inboxFolderParts.Add($"{kvp.Key}:{kvp.Value}");
                }
            }

            var userEmailStr = string.Join(";", userEmailList);
            var inboxFolderStr = inboxFolderParts.Any() ? string.Join(";", inboxFolderParts) : "Inbox";

            var configPath = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath;
            var doc = XDocument.Load(configPath);

            var appSettings = doc.Root?.Element("appSettings");
            var connectionStrings = doc.Root?.Element("connectionStrings");

            if (appSettings == null || connectionStrings == null)
            {
                throw new Exception("Estrutura do App.config inválida");
            }

            // Update appSettings
            UpdateSetting(appSettings, "TenantId", txtTenantId.Text);
            UpdateSetting(appSettings, "ClientId", txtClientId.Text);
            UpdateSetting(appSettings, "ClientSecret", txtClientSecret.Text);
            UpdateSetting(appSettings, "UserEmail", userEmailStr);
            UpdateSetting(appSettings, "InboxFolder", inboxFolderStr);
            UpdateSetting(appSettings, "ServiceIntervalSeconds", numServiceInterval.Value.ToString());
            UpdateSetting(appSettings, "EmailsPerCycle", numEmailsPerCycle.Value.ToString());
            UpdateSetting(appSettings, "MaxRetries", numMaxRetries.Value.ToString());
            UpdateSetting(appSettings, "ConnectionTestTimeoutSeconds", numConnectionTimeout.Value.ToString());
            UpdateSetting(appSettings, "ProcessOnlyUnread", chkProcessOnlyUnread.Checked.ToString().ToLower());
            UpdateSetting(appSettings, "TempFolder", txtTempFolder.Text);
            UpdateSetting(appSettings, "ProcessedFolder", txtProcessedFolder.Text);
            UpdateSetting(appSettings, "ErrorFolder", txtErrorFolder.Text);
            UpdateSetting(appSettings, "SaveMimeContent", chkSaveMimeContent.Checked.ToString().ToLower());
            UpdateSetting(appSettings, "LogLevel", cboLogLevel.SelectedItem?.ToString() ?? "Information");
            UpdateSetting(appSettings, "LogPath", txtLogPath.Text);
            UpdateSetting(appSettings, "LogRetentionDays", numLogRetention.Value.ToString());
            UpdateSetting(appSettings, "RewriteLog", chkRewriteLog.Checked.ToString().ToLower());

            // Update connectionString
            var connStringElement = connectionStrings.Elements("add").FirstOrDefault(e => e.Attribute("name")?.Value == "Outlook2DAM");
            if (connStringElement != null)
            {
                connStringElement.SetAttributeValue("connectionString", txtConnectionString.Text);
            }

            doc.Save(configPath);
            ConfigurationManager.RefreshSection("appSettings");
            ConfigurationManager.RefreshSection("connectionStrings");

            _logger.Information("Configuração guardada com sucesso");
            MessageBox.Show("Configuração guardada com sucesso!\n\nReinicie a aplicação para aplicar as alterações.", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao guardar configuração");
            MessageBox.Show($"Erro ao guardar configuração: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void UpdateSetting(XElement appSettings, string key, string value)
    {
        var setting = appSettings.Elements("add").FirstOrDefault(e => e.Attribute("key")?.Value == key);
        if (setting != null)
        {
            setting.SetAttributeValue("value", value);
        }
        else
        {
            appSettings.Add(new XElement("add",
                new XAttribute("key", key),
                new XAttribute("value", value)));
        }
    }
}