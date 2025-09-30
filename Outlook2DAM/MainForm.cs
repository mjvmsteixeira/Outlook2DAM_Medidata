using Outlook2DAM.Services;
using Microsoft.Graph;
using Serilog;
using System.Configuration;
using System.Diagnostics;

namespace Outlook2DAM;

public partial class MainForm : Form
{
    private readonly ILogger _logger;
    private readonly OutlookService _outlookService;
    private readonly GraphServiceClient _graphClient;

    private Label lblStatus = null!;
    private Button btnStartStop = null!;
    private Button btnOpenLogs = null!;
    private Button btnEditConfig = null!;
    private PropertyGrid configGrid = null!;

    public MainForm(OutlookService outlookService, GraphServiceClient graphClient)
    {
        _outlookService = outlookService ?? throw new ArgumentNullException(nameof(outlookService));
        _graphClient = graphClient ?? throw new ArgumentNullException(nameof(graphClient));
        _logger = LoggerService.GetLogger<MainForm>();

        InitializeControls();

        // Configurar controles
        btnStartStop.Text = "Iniciar";
        lblStatus.Text = "Serviço parado";
        lblStatus.ForeColor = Color.Red;

        _logger.Debug("Os controles do formulário foram configurados corretamente!");

        // Configurar eventos
        btnStartStop.Click += btnStartStop_Click;
        btnOpenLogs.Click += BtnOpenLogs_Click;
        btnEditConfig.Click += BtnEditConfig_Click;
        FormClosing += MainForm_FormClosing;
        _outlookService.UnreadEmailCountChanged += UpdateStatus;
    }

    private void InitializeControls()
    {
        Text = "Outlook2DAM";
        Size = new Size(800, 600);

        // Status Label
        lblStatus = new Label
        {
            AutoSize = true,
            Location = new Point(10, 10)
        };
        Controls.Add(lblStatus);

        // Start/Stop Button
        btnStartStop = new Button
        {
            Size = new Size(120, 30),
            Location = new Point(10, 40)
        };
        Controls.Add(btnStartStop);

        // Open Logs Button
        btnOpenLogs = new Button
        {
            Text = "Abrir Logs",
            Size = new Size(120, 30),
            Location = new Point(140, 40)
        };
        Controls.Add(btnOpenLogs);

        // Edit Config Button
        btnEditConfig = new Button
        {
            Text = "⚙️ Configurações",
            Size = new Size(140, 30),
            Location = new Point(270, 40)
        };
        Controls.Add(btnEditConfig);

        // Configuration Grid (read-only view)
        configGrid = new PropertyGrid
        {
            SelectedObject = new ConfigSettings(),
            Location = new Point(10, 80),
            Size = new Size(760, 470),
            Dock = DockStyle.Bottom,
            HelpVisible = false,
            ToolbarVisible = false,
            PropertySort = PropertySort.Categorized
        };
        Controls.Add(configGrid);
    }

    private void btnStartStop_Click(object? sender, EventArgs e)
    {
        try
        {
            if (_outlookService.IsRunning)
            {
                _outlookService.StopService();
                btnStartStop.Text = "Iniciar";
                btnStartStop.BackColor = Color.LightGreen;
            }
            else
            {
                _outlookService.StartService();
                btnStartStop.Text = "Parar";
                btnStartStop.BackColor = Color.LightCoral;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Erro ao iniciar/parar o serviço: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void UpdateStatus(int emailCount)
    {
        if (InvokeRequired)
        {
            Invoke(new Action<int>(UpdateStatus), emailCount);
            return;
        }

        if (_outlookService.IsRunning)
        {
            lblStatus.Text = $"Serviço em execução - {emailCount} emails não lidos na Caixa de Entrada";
        }
    }

    private void BtnOpenLogs_Click(object? sender, EventArgs e)
    {
        try
        {
            var logsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            if (Directory.Exists(logsPath))
            {
                Process.Start("explorer.exe", logsPath);
            }
            else
            {
                MessageBox.Show("Pasta de logs não encontrada", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao abrir pasta de logs");
            MessageBox.Show($"Erro ao abrir pasta de logs: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnEditConfig_Click(object? sender, EventArgs e)
    {
        try
        {
            if (_outlookService.IsRunning)
            {
                var result = MessageBox.Show(
                    "O serviço está em execução. Deseja pará-lo para editar as configurações?",
                    "Serviço em Execução",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    _outlookService.StopService();
                }
                else
                {
                    return;
                }
            }

            using var configForm = new ConfigEditorForm(_graphClient);
            var dialogResult = configForm.ShowDialog(this);

            if (dialogResult == DialogResult.OK)
            {
                // Reload configuration in PropertyGrid
                configGrid.SelectedObject = new ConfigSettings();
                configGrid.Refresh();

                MessageBox.Show(
                    "Configuração atualizada!\n\nReinicie a aplicação para aplicar todas as alterações.",
                    "Configuração Atualizada",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Erro ao abrir editor de configurações");
            MessageBox.Show($"Erro ao abrir editor: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_outlookService.IsRunning)
        {
            _outlookService.StopService();
        }
    }
}
