# Outlook2DAM - Conector de Email para Medidata SIDAM

## Descrição

O **Outlook2DAM** é um serviço Windows desenvolvido em C# (.NET 9.0) que automatiza o processamento de emails através do Microsoft Graph API. O sistema monitora emails não lidos, extrai informações, processa anexos, gera documentos PDF e XML estruturados, e armazena metadados em banco de dados Oracle para integração com sistemas Medidata SIDAM.

## Índice

- [Recursos Principais](#recursos-principais)
- [Requisitos do Sistema](#requisitos-do-sistema)
- [Arquitetura](#arquitetura)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Configuração](#configuração)
- [Instalação](#instalação)
- [Processamento de Emails](#processamento-de-emails)
- [Sistema de Logging](#sistema-de-logging)
- [Segurança](#segurança)
- [Resolução de Problemas](#resolução-de-problemas)
- [Documentação Técnica](#documentação-técnica)

---

## Recursos Principais

### 📧 Email & Integração
- ✅ Monitoramento automático de emails não lidos via **Microsoft Graph API**
- ✅ **Suporte completo para Shared Mailboxes** (caixas de correio partilhadas)
- ✅ Suporte a múltiplas contas de email simultâneas
- ✅ **Pastas de entrada personalizadas por email** (InboxFolder configurável)
- ✅ **Filtragem inteligente de destinatários** (apenas emails configurados no XML)
- ✅ Gestão automática de pastas (Processados/Erros) criadas no Outlook
- ✅ Validação proativa de pastas configuradas com listagem automática

### 📁 Processamento & Armazenamento
- ✅ Processamento robusto de anexos com validação de integridade
- ✅ Geração automática de PDFs do corpo do email (HTML/texto)
- ✅ Criação de XML estruturado para integração com DAM
- ✅ Suporte a **paths UNC** (caminhos de rede `\\servidor\share\`)
- ✅ Mecanismo de retry para operações de I/O e rede
- ✅ Salvamento opcional de arquivo .eml original

### 💾 Base de Dados
- ✅ **Suporte multi-database**: Oracle, SQL Server, MS Access
- ✅ Detecção automática de provider (OraOLEDB, SQLOLEDB, MSOLEDBSQL, ACE)
- ✅ Queries SQL adaptadas automaticamente por provider
- ✅ Health check de conectividade com fallback inteligente

### 🖥️ Interface & Experiência
- ✅ **Editor gráfico completo de configurações** (GUI com tabs)
- ✅ **Dropdown inteligente para pastas** (carrega do Outlook via API)
- ✅ Interface gráfica para modo debug (Windows Forms)
- ✅ Modo CLI para execução como serviço Windows
- ✅ PropertyGrid read-only para visualização rápida

### 🔒 Segurança & Logs
- ✅ Sistema de logging detalhado com **Serilog**
- ✅ **Validação de segurança**: prevenção de path traversal
- ✅ Mascaramento de dados sensíveis nos logs
- ✅ Rotação automática de logs com compressão
- ✅ Health checks completos (Graph API, BD, TempFolder)

---

## Requisitos do Sistema

### Software
- **Windows 10** ou superior
- **.NET 9.0 SDK**
  ```powershell
  winget install Microsoft.DotNet.SDK.9
  ```
- **Visual Studio 2022** (recomendado para desenvolvimento)
- **Oracle Database** (com driver OLEDB)
- **Conta Microsoft Azure AD** com aplicação registrada

### Permissões Microsoft Graph
A aplicação Azure AD requer as seguintes permissões (tipo **Application**):
- `Mail.Read` - Ler emails
- `Mail.ReadWrite` - Mover emails entre pastas
- `MailboxSettings.Read` - Configurações de caixa de correio

**📧 Suporte para Shared Mailboxes:**
- Shared mailboxes são totalmente suportadas com as mesmas permissões
- Não requer permissões adicionais
- Configure o email da shared mailbox diretamente em `UserEmail`
- Exemplo: `<add key="UserEmail" value="shared@dominio.com" />`

### Hardware Mínimo
- **CPU**: 2 cores
- **RAM**: 2GB
- **Disco**: 100MB para instalação + espaço para armazenamento de emails

---

## Arquitetura

### Fluxo de Processamento

```
┌─────────────────────────────────────────────────────────────┐
│                      OUTLOOK2DAM                             │
├─────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌──────────────┐    ┌────────────────────────────────┐    │
│  │   Program    │───▶│   OutlookService (Timer)       │    │
│  │ (CLI/GUI)    │    │   - Verifica emails não lidos  │    │
│  └──────────────┘    │   - Gerencia pastas            │    │
│         │             └─────────────┬──────────────────┘    │
│         │                           │                        │
│         ▼                           ▼                        │
│  ┌──────────────┐    ┌────────────────────────────────┐    │
│  │ TokenProvider│    │    EmailProcessor              │    │
│  │  (OAuth2)    │    │   - Download anexos            │    │
│  └──────────────┘    │   - Geração PDF                │    │
│         │             │   - Criação XML                │    │
│         │             │   - Insert Oracle              │    │
│         │             └─────────────┬──────────────────┘    │
│         │                           │                        │
│         ▼                           ▼                        │
│  ┌──────────────┐    ┌────────────────────────────────┐    │
│  │ Graph API    │    │    LoggerService               │    │
│  │ (Microsoft)  │    │   - Logs estruturados          │    │
│  └──────────────┘    │   - Rotação diária             │    │
│                       └────────────────────────────────┘    │
│                                                               │
└─────────────────────────────────────────────────────────────┘
           │                          │
           ▼                          ▼
    ┌────────────┐          ┌──────────────────┐
    │   Oracle   │          │  File System     │
    │  Database  │          │  (Emails + XML)  │
    └────────────┘          └──────────────────┘
```

### Camadas e Componentes

| Camada | Componente | Responsabilidade |
|--------|------------|------------------|
| **Apresentação** | `Program.cs`, `MainForm.cs` | Ponto de entrada, interface gráfica |
| **Serviços** | `OutlookService.cs` | Orquestração, timer, gestão de pastas |
| | `EmailProcessor.cs` | Processamento de emails e anexos |
| | `LoggerService.cs` | Sistema de logging |
| | `ConnectionTester.cs` | Testes de conectividade |
| **Autenticação** | `TokenProvider.cs` | Autenticação OAuth2 com MSAL |
| **Configuração** | `ConfigSettings.cs` | Gestão de configurações |
| **Modelos** | `Correspondencia.cs` | Modelo XML |
| | `OutlookEmail.cs` | Modelo Oracle |

---

## Estrutura do Projeto

```
Outlook2DAM/
├── Program.cs                    # Ponto de entrada (CLI/GUI)
├── MainForm.cs                   # Interface gráfica (modo debug)
├── ConfigSettings.cs             # Gerenciamento de configurações
├── TokenProvider.cs              # Autenticação Microsoft Graph
├── Outlook2DAM.csproj           # Projeto .NET
│
├── Services/
│   ├── OutlookService.cs        # Orquestrador principal (timer, pastas)
│   ├── EmailProcessor.cs        # Processamento de emails e anexos
│   ├── LoggerService.cs         # Sistema de logs com Serilog
│   └── ConnectionTester.cs      # Testes de conectividade
│
├── Models/
│   ├── Correspondencia.cs       # Modelo para estrutura XML
│   └── OutlookEmail.cs          # Modelo de dados Oracle
│
├── Configuração/
│   ├── App-default.config       # 📝 Template (versionado)
│   ├── App.config               # 🔒 Config real (NÃO versionado)
│   └── CONFIGURACAO.md          # 📖 Guia de configuração
│
├── .gitignore                   # Exclui App.config e credenciais
├── README.md                    # Documentação principal
│
└── logs/                        # Logs gerados automaticamente
    └── outlook2dam-YYYYMMDD.log
```

### 📁 Ficheiros de Configuração

| Ficheiro | Versionado | Descrição |
|----------|------------|-----------|
| **App-default.config** | ✅ Sim | Template com todas as opções, sem credenciais |
| **App.config** | ❌ Não | Configuração real com credenciais (gitignore) |
| **CONFIGURACAO.md** | ✅ Sim | Guia detalhado de configuração |

---

## Configuração

### 📝 Editor Gráfico de Configurações (NOVO!)

O Outlook2DAM agora inclui um **editor gráfico completo** para todas as configurações:

1. **Inicie a aplicação em modo GUI** (duplo clique no executável)
2. Clique no botão **"⚙️ Configurações"**
3. Navegue pelas abas:
   - **Azure AD**: Credenciais do Azure (TenantId, ClientId, ClientSecret)
   - **📧 Emails & Pastas**: Gestão de múltiplos emails com pastas personalizadas
     - ➕ Adicionar/remover emails
     - 🔄 Botão "Listar Pastas" carrega pastas do Outlook em tempo real
     - Visualização: `email@domain.com → NomeDaPasta`
   - **⚙️ Serviço**: Intervalos, timeouts, retries
   - **📁 Pastas**: TempFolder, ProcessedFolder, ErrorFolder
   - **📋 Logs**: Níveis, retenção, paths
   - **💾 Base de Dados**: Connection string com exemplos
4. Clique **"💾 Guardar"** para atualizar o App.config
5. Reinicie a aplicação para aplicar as alterações

**Vantagens do Editor:**
- ✅ Interface intuitiva com validação em tempo real
- ✅ Carregamento dinâmico de pastas do Outlook via Graph API
- ✅ Suporte visual para configuração email→pasta
- ✅ Guarda diretamente no XML (App.config)
- ✅ Instruções inline em cada aba

---

### 1. Registrar Aplicação no Azure AD

1. Aceda ao [Azure Portal](https://portal.azure.com)
2. Navegue até **Azure Active Directory** > **App registrations**
3. Clique em **New registration**
4. Configure:
   - **Name**: Outlook2DAM
   - **Supported account types**: Single tenant
5. Após criação, anote:
   - **Application (client) ID**
   - **Directory (tenant) ID**
6. Vá em **Certificates & secrets** > **New client secret**
   - Anote o **Value** (Client Secret)
7. Em **API permissions**, adicione:
   - Microsoft Graph > Application permissions:
     - `Mail.Read`
     - `Mail.ReadWrite`
     - `MailboxSettings.Read`
   - Clique em **Grant admin consent**

### 2. Configurar App.config

⚠️ **IMPORTANTE**:
- Use [App-default.config](Outlook2DAM/App-default.config) como template
- Copie para `App.config` e preencha com suas credenciais
- `App.config` está no `.gitignore` e não será versionado
- Ver [CONFIGURACAO.md](Outlook2DAM/CONFIGURACAO.md) para guia detalhado

```powershell
# Criar App.config a partir do template
Copy-Item Outlook2DAM\App-default.config Outlook2DAM\App.config
```

Edite o arquivo `App.config` no diretório da aplicação:

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <!-- ========================================
         AUTENTICAÇÃO MICROSOFT GRAPH
         ======================================== -->
    <add key="TenantId" value="00000000-0000-0000-0000-000000000000" />
    <add key="ClientId" value="00000000-0000-0000-0000-000000000000" />
    <add key="ClientSecret" value="seu_client_secret_aqui" />

    <!-- Suporta múltiplos emails separados por ponto e vírgula -->
    <add key="UserEmail" value="email1@dominio.com;email2@dominio.com" />

    <!-- ========================================
         CONFIGURAÇÕES DE PASTAS
         ======================================== -->
    <!-- Nome da pasta de entrada a monitorizar (opcional, padrão: "Inbox") -->

    <!-- OPÇÃO 1: Pasta única para todos os emails -->
    <add key="InboxFolder" value="Inbox" />

    <!-- OPÇÃO 2: Pasta diferente por email -->
    <!-- Formato: email:pasta;email:pasta -->
    <!-- Exemplo: -->
    <!-- <add key="InboxFolder" value="email1@domain.com:Contraordenações;email2@domain.com:Processos" /> -->
    <!-- Emails não listados usam "Inbox" por padrão -->

    <!-- ========================================
         CONFIGURAÇÕES DO SERVIÇO
         ======================================== -->
    <!-- Intervalo entre verificações (segundos) -->
    <add key="ServiceIntervalSeconds" value="60" />

    <!-- Número máximo de emails processados por ciclo -->
    <add key="EmailsPerCycle" value="1" />

    <!-- Número de tentativas em caso de erro -->
    <add key="MaxRetries" value="3" />

    <!-- Timeout para testes de conexão (segundos) -->
    <add key="ConnectionTestTimeoutSeconds" value="30" />

    <!-- ========================================
         CONFIGURAÇÕES DE PASTAS
         ======================================== -->
    <!-- Pasta base para armazenamento (UNC ou local) -->
    <add key="TempFolder" value="\\servidor\share\Outlook2DAM" />

    <!-- Nome da pasta para emails processados (criada no Outlook) -->
    <add key="ProcessedFolder" value="Processados" />

    <!-- Nome da pasta para emails com erro (criada no Outlook) -->
    <add key="ErrorFolder" value="Errors" />

    <!-- Salvar arquivo .eml original (true/false) -->
    <add key="SaveMimeContent" value="true" />

    <!-- ========================================
         CONFIGURAÇÕES DE LOG
         ======================================== -->
    <!-- Níveis: Verbose, Debug, Information, Warning, Error, Fatal -->
    <add key="LogLevel" value="Debug" />

    <!-- Caminho dos logs (relativo ao executável) -->
    <add key="LogPath" value="logs" />

    <!-- Número de dias para manter logs antigos -->
    <add key="LogRetentionDays" value="31" />

    <!-- Se true, reutiliza arquivo de log. Se false, comprime o anterior -->
    <add key="RewriteLog" value="false" />
  </appSettings>

  <connectionStrings>
    <!-- ========================================
         ORACLE DATABASE
         ======================================== -->
    <add name="Outlook2DAM"
         connectionString="Provider=OraOLEDB.Oracle;Data Source=SEU_DATASOURCE;Password=SUA_SENHA;User ID=SEU_USUARIO;Persist Security Info=True" />

    <!-- ========================================
         SQL SERVER (Alternativa)
         ======================================== -->
    <!-- Descomente para usar SQL Server
    <add name="Outlook2DAM"
         connectionString="Provider=SQLOLEDB;Data Source=SEU_SERVIDOR;Initial Catalog=SUA_DATABASE;User ID=SEU_USUARIO;Password=SUA_SENHA;" />
    -->

    <!-- ========================================
         SQL SERVER (Native Client 11.0)
         ======================================== -->
    <!-- Descomente para usar SQL Server Native Client
    <add name="Outlook2DAM"
         connectionString="Provider=SQLNCLI11;Server=SEU_SERVIDOR;Database=SUA_DATABASE;Uid=SEU_USUARIO;Pwd=SUA_SENHA;" />
    -->

    <!-- ========================================
         SQL SERVER (Microsoft OLE DB Driver)
         ======================================== -->
    <!-- Descomente para usar Microsoft OLE DB Driver for SQL Server
    <add name="Outlook2DAM"
         connectionString="Provider=MSOLEDBSQL;Server=SEU_SERVIDOR;Database=SUA_DATABASE;UID=SEU_USUARIO;PWD=SUA_SENHA;" />
    -->
  </connectionStrings>
</configuration>
```

**NOTA**: O sistema detecta automaticamente o provider baseado na connection string e adapta as queries SQL de acordo.

---

## Configurações Avançadas

### 📧 Shared Mailboxes (Caixas de Correio Partilhadas)

O Outlook2DAM suporta **nativamente shared mailboxes** sem configuração adicional:

```xml
<!-- Configurar shared mailbox igual a mailbox normal -->
<add key="UserEmail" value="shared@dominio.com" />

<!-- Ou misturar mailboxes normais com shared -->
<add key="UserEmail" value="user@dominio.com;shared@dominio.com;outro-shared@dominio.com" />
```

**Requisitos:**
- ✅ Mesmas permissões (Mail.Read, Mail.ReadWrite, MailboxSettings.Read)
- ✅ Não requer configuração especial no Azure AD
- ✅ Health check adaptado automaticamente
- ✅ Funciona com pastas personalizadas por email

**Como funciona:**
- A App Registration acede à shared mailbox através do endpoint `Users[shared@domain.com]`
- O Graph API permite acesso a `.Messages` e `.MailFolders` de shared mailboxes
- A aplicação detecta automaticamente se é shared ou mailbox normal

---

### 📁 Pastas de Entrada Personalizadas (InboxFolder)

Configure pastas diferentes para cada email monitorizado:

#### **Modo 1: Pasta única para todos**
```xml
<!-- Todos os emails monitorizam a mesma pasta -->
<add key="InboxFolder" value="Contraordenações" />
```

#### **Modo 2: Pasta diferente por email**
```xml
<!-- Cada email tem sua própria pasta -->
<add key="InboxFolder" value="email1@domain.com:Contraordenações;email2@domain.com:Processos;email3@domain.com:Inbox" />
```

#### **Modo 3: Misto (alguns personalizados, outros padrão)**
```xml
<!-- Emails não listados usam "Inbox" por padrão -->
<add key="UserEmail" value="email1@domain.com;email2@domain.com;email3@domain.com" />
<add key="InboxFolder" value="email1@domain.com:Contraordenações" />
<!-- email2 e email3 usarão "Inbox" automaticamente -->
```

**💡 Dicas:**
- Use o **Editor Gráfico** para configurar visualmente (botão "🔄 Listar Pastas")
- A pasta deve existir na caixa de correio antes de iniciar o serviço
- Nomes de pastas são **case-sensitive**
- Se a pasta não for encontrada, a aplicação lista todas as pastas disponíveis nos logs

**Validação:**
```powershell
# Ao iniciar, os logs mostram:
[INF] Outlook2DAM inicializado. Pastas entrada: [email1→Teste1, email2→Teste2], ...
```

---

### 🌐 Caminhos UNC (Network Paths)

Suporte completo para pastas de rede:

```xml
<!-- Caminho UNC válido -->
<add key="TempFolder" value="\\servidor\share\Outlook2DAM\" />

```

**Requisitos:**
- ✅ Conta de serviço deve ter permissões de leitura/escrita na share
- ✅ Share deve estar acessível pela rede
- ✅ Validação automática de segurança (previne path traversal)

---

### 3. Estrutura da Base de Dados

#### Oracle Database

Crie a tabela `outlook` no Oracle:

```sql
CREATE TABLE outlook (
    chave           VARCHAR2(255) PRIMARY KEY,
    remetente       VARCHAR2(255),
    data            DATE,
    hora            NUMBER,
    destinatario    VARCHAR2(1000),
    assunto         VARCHAR2(500),
    caminho_ficheiro VARCHAR2(1000),
    processado      VARCHAR2(1),
    tipodoc         VARCHAR2(50),
    chavedoc        VARCHAR2(255),
    observacoes     CLOB
);
```

#### SQL Server

Crie a tabela `outlook` no SQL Server:

```sql
CREATE TABLE outlook (
    chave           NVARCHAR(255) PRIMARY KEY,
    remetente       NVARCHAR(255),
    data            DATE,
    hora            INT,
    destinatario    NVARCHAR(1000),
    assunto         NVARCHAR(500),
    caminho_ficheiro NVARCHAR(1000),
    processado      NVARCHAR(1),
    tipodoc         NVARCHAR(50),
    chavedoc        NVARCHAR(255),
    observacoes     NVARCHAR(MAX)
);
```

---

## Instalação

### Desenvolvimento (Modo Debug - GUI)

1. Clone o repositório ou extraia os arquivos
2. Abra `Outlook2DAM.sln` no Visual Studio 2022
3. Restaure os pacotes NuGet:
   ```powershell
   dotnet restore
   ```
4. Configure `App.config` conforme seção anterior
5. Pressione `F5` para executar em modo debug
6. Use a interface gráfica para:
   - Iniciar/Parar o serviço
   - Visualizar logs em tempo real
   - Editar configurações
   - Abrir pasta de logs

### Produção (Modo CLI - Serviço Windows)

#### 1. Compilar a Aplicação

```powershell
# Compilar para Release
dotnet publish -c Release -o C:\Outlook2DAM

# Copiar arquivo de configuração
copy App.config C:\Outlook2DAM\
```

#### 2. Criar Serviço Windows

```powershell
# Abrir PowerShell como Administrador

# Criar o serviço
sc.exe create "Outlook2DAM" `
    binpath= "C:\Outlook2DAM\Outlook2DAM.exe --cli" `
    start= auto `
    DisplayName= "Outlook2DAM - Email Processor"

# Adicionar descrição
sc.exe description "Outlook2DAM" "Processamento automatizado de emails via Microsoft Graph API"

# Configurar recuperação automática em caso de falha
sc.exe failure "Outlook2DAM" reset= 86400 actions= restart/60000/restart/120000/restart/300000
```

#### 3. Gerenciar o Serviço

```powershell
# Iniciar serviço
sc.exe start "Outlook2DAM"

# Verificar status
sc.exe query "Outlook2DAM"

# Parar serviço
sc.exe stop "Outlook2DAM"

# Ver logs em tempo real
Get-Content C:\Outlook2DAM\logs\outlook2dam-*.log -Wait -Tail 50

# Remover serviço (se necessário)
sc.exe delete "Outlook2DAM"
```

#### 4. Configurar Permissões de Pasta

```powershell
# Garantir que o serviço tem acesso à pasta de armazenamento
icacls "\\servidor\share\Outlook2DAM" /grant "NETWORK SERVICE:(OI)(CI)F"
```

---

## Processamento de Emails

### Fluxo Detalhado

```
1. Timer dispara (intervalo configurável)
   ↓
2. Para cada conta de email configurada:
   ├─ Verifica emails não lidos na Inbox
   ├─ Garante que pasta "Processados" existe
   └─ Processa até N emails (EmailsPerCycle)
   ↓
3. Para cada email:
   ├─ Cria pasta única: YYYYMMDD_HHMMSS_[ID]
   ├─ Salva .eml original (se SaveMimeContent=true)
   ├─ Baixa todos os anexos
   ├─ Gera PDF do corpo (HTML ou texto)
   ├─ Cria arquivo XML com metadados
   ├─ Insere registro no Oracle
   ├─ Marca email como lido
   └─ Move para pasta "Processados"
   ↓
4. Em caso de erro:
   ├─ Retry automático (até MaxRetries)
   ├─ Delay progressivo (500ms × tentativa)
   └─ Após max retries: move para pasta "Errors"
```

### Estrutura de Pastas Gerada

Cada email processado cria uma pasta única:

```
\\servidor\share\Outlook2DAM\
└── 20250130_143522_a1b2c3d4\
    ├── 20250130_143522_a1b2c3d4.eml  ← Email original (se SaveMimeContent=true)
    ├── email.pdf                      ← Corpo do email em PDF
    ├── email.xml                      ← Metadados estruturados
    ├── documento1.pdf                 ← Anexo 1
    ├── foto.jpg                       ← Anexo 2
    └── relatorio.xlsx                 ← Anexo N
```

### Estrutura do XML Gerado

```xml
<correspondencia>
    <!-- Via de correspondência (E=Email) -->
    <via>E</via>

    <!-- Data/hora de recepção -->
    <data>2025-01-30T14:35:22</data>
    <hora>52522</hora> <!-- Hora em segundos: 14*3600 + 35*60 + 22 -->

    <!-- Informações do email -->
    <assunto>Proposta Comercial Q1 2025</assunto>
    <from>remetente@empresa.com</from>
    <to>geral@empresa.pt</to> <!-- Apenas emails configurados em UserEmail -->

    <!-- Localização dos arquivos -->
    <pasta>\\servidor\share\Outlook2DAM\20250130_143522_a1b2c3d4\\</pasta>
    <ficheiro>email.pdf</ficheiro>

    <!-- Lista de anexos (inclui .eml se SaveMimeContent=true) -->
    <anexos>
        <anexo>20250130_143522_a1b2c3d4.eml</anexo>
        <anexo>documento1.pdf</anexo>
        <anexo>foto.jpg</anexo>
        <anexo>relatorio.xlsx</anexo>
    </anexos>

    <!-- Versão do processamento -->
    <ver>0</ver>
</correspondencia>
```

#### Filtragem de Destinatários no Campo `<to>`

⚠️ **IMPORTANTE**: O campo `<to>` é **filtrado automaticamente**:

- **Filtra apenas emails configurados em `UserEmail`** (App.config)
- Remove destinatários externos ou não monitorados
- Útil para emails enviados para múltiplos destinatários

**Exemplo**:
- **Email original**: `To: geral@empresa.pt, outro@empresa.com, externo@gmail.com`
- **UserEmail configurado**: `geral@empresa.pt`
- **XML gerado**: `<to>geral@empresa.pt</to>`

Se nenhum destinatário corresponder aos emails configurados, usa o email da conta que processou a mensagem.

### Registro na Base de Dados Oracle

| Campo | Tipo | Descrição | Exemplo |
|-------|------|-----------|---------|
| `chave` | VARCHAR2(255) | ID único do email (Graph API) | `AAMkAGI2...` |
| `remetente` | VARCHAR2(255) | Email do remetente | `remetente@empresa.com` |
| `data` | DATE | Data de recepção | `30-JAN-2025` |
| `hora` | NUMBER | Hora em segundos | `52522` |
| `destinatario` | VARCHAR2(1000) | Destinatários (separados por `;`) | `dest1@email.com;dest2@email.com` |
| `assunto` | VARCHAR2(500) | Assunto do email | `Proposta Comercial` |
| `caminho_ficheiro` | VARCHAR2(1000) | Path do XML | `\\servidor\...\email.xml` |
| `processado` | VARCHAR2(1) | Flag de processamento | `0` |
| `tipodoc` | VARCHAR2(50) | Tipo de documento | `` |
| `chavedoc` | VARCHAR2(255) | Chave externa | `` |
| `observacoes` | CLOB | Observações | `` |

---

## Sistema de Logging

### Localização dos Logs

```
[DiretórioInstalação]\logs\
└── outlook2dam-20250130.log  ← Rotação diária automática
```

### Níveis de Log

| Nível | Uso | Exemplo |
|-------|-----|---------|
| **Debug** | Detalhes de configuração, queries, conteúdo XML | `XML criado com sucesso. Conteúdo: <correspondencia>...` |
| **Information** | Eventos importantes do sistema | `Email processado com sucesso: Proposta Comercial` |
| **Warning** | Situações anormais não-críticas | `Arquivo ainda está bloqueado, tentativa 2 de 3` |
| **Error** | Falhas de processamento | `Erro ao processar email (Tentativa 3/3): Timeout` |
| **Fatal** | Erros críticos do sistema | `Erro fatal na aplicação` |

### Exemplo de Log Completo

```
2025-01-30 14:35:22.123 +00:00 [INF] Iniciando serviço em modo CLI...
2025-01-30 14:35:22.456 +00:00 [INF] Intervalo do serviço configurado para 60 segundos
2025-01-30 14:35:22.789 +00:00 [DBG] A iniciar o TokenProvider...
2025-01-30 14:35:23.012 +00:00 [INF] Token obtido com sucesso
2025-01-30 14:35:23.234 +00:00 [INF] Serviço iniciado com sucesso
2025-01-30 14:35:23.567 +00:00 [INF] Encontrados 5 emails não lidos em user@empresa.com. Limite por ciclo: 1
2025-01-30 14:35:23.890 +00:00 [INF] Processando email: Proposta Comercial de cliente@example.com (Tentativa 1/3)
2025-01-30 14:35:24.123 +00:00 [DBG] SaveMimeContent está ativado, salvando EML...
2025-01-30 14:35:24.456 +00:00 [INF] Arquivo EML salvo com sucesso: \\servidor\...\email.eml
2025-01-30 14:35:24.789 +00:00 [DBG] Email tem anexos, processando...
2025-01-30 14:35:25.012 +00:00 [INF] Anexo salvo com sucesso: documento1.pdf
2025-01-30 14:35:25.234 +00:00 [INF] Anexo salvo com sucesso: foto.jpg
2025-01-30 14:35:25.567 +00:00 [DBG] PDF do corpo do email criado e validado em: \\servidor\...\email.pdf
2025-01-30 14:35:25.890 +00:00 [DBG] XML criado com sucesso. Conteúdo:
<correspondencia><via>E</via>...
2025-01-30 14:35:26.123 +00:00 [DBG] Email inserido na Base de Dados com sucesso!
2025-01-30 14:35:26.456 +00:00 [INF] Email movido para pasta Processados
2025-01-30 14:35:26.789 +00:00 [INF] Email processado com sucesso: Proposta Comercial
```

### Visualizar Logs em Tempo Real

```powershell
# PowerShell
Get-Content C:\Outlook2DAM\logs\outlook2dam-*.log -Wait -Tail 50

# CMD
powershell -Command "Get-Content C:\Outlook2DAM\logs\outlook2dam-*.log -Wait -Tail 50"
```

---

## Segurança

### Boas Práticas

#### 1. Proteção de Credenciais

⚠️ **IMPORTANTE**: O arquivo `App.config` contém credenciais sensíveis.

```powershell
# Definir permissões NTFS (somente Administradores)
icacls "C:\Outlook2DAM\App.config" /inheritance:r
icacls "C:\Outlook2DAM\App.config" /grant:r "Administrators:(R)"
icacls "C:\Outlook2DAM\App.config" /grant:r "SYSTEM:(R)"
icacls "C:\Outlook2DAM\App.config" /grant:r "NETWORK SERVICE:(R)"
```

#### 2. Rotação de Secrets

- Configure **expiração automática** do Client Secret no Azure (máximo 24 meses)
- Renove secrets 30 dias antes da expiração
- Mantenha histórico de secrets para rollback

#### 3. Princípio do Menor Privilégio

**Permissões Azure AD**:
- Use apenas permissões necessárias (`Mail.Read`, `Mail.ReadWrite`)
- Evite permissões delegadas; use application permissions

**Conta Oracle**:
- Crie usuário específico com permissões mínimas:
  ```sql
  CREATE USER outlook2dam IDENTIFIED BY senha_forte;
  GRANT CONNECT, RESOURCE TO outlook2dam;
  GRANT INSERT ON outlook TO outlook2dam;
  ```

---

## Resolução de Problemas

### 1. Erro de Autenticação Microsoft Graph

**Sintoma**:
```
[ERR] Erro do serviço MSAL ao obter token. Código: invalid_client
```

**Soluções**:
- ✅ Verifique se `TenantId`, `ClientId` e `ClientSecret` estão corretos
- ✅ Confirme que o Client Secret não expirou (Azure Portal)
- ✅ Verifique se as permissões foram concedidas (Grant admin consent)

### 2. Erro de Conexão Oracle

**Sintoma**:
```
[ERR] Database connection test failed
System.Data.OleDb.OleDbException: ORA-12154: TNS:could not resolve the connect identifier
```

**Soluções**:
- ✅ Verifique a string de conexão no `App.config`
- ✅ Confirme que o driver Oracle OLEDB está instalado
- ✅ Verifique se `tnsnames.ora` está configurado (se usar TNS)

### 3. Falha ao Criar/Acessar Pastas

**Sintoma**:
```
[ERR] Erro ao criar/verificar diretório: \\servidor\share\Outlook2DAM
System.UnauthorizedAccessException: Access to the path is denied
```

**Soluções**:
- ✅ Verifique permissões NTFS/SMB da pasta
- ✅ Garanta que a conta do serviço (NETWORK SERVICE) tem permissões

---

## Documentação Técnica

### Componentes Principais

#### 1. **Program.cs**
- **Responsabilidade**: Ponto de entrada da aplicação
- **Métodos-chave**:
  - `Main()` - Inicialização, configuração DI, detecção modo CLI/GUI
  - `RunCliMode()` - Execução em modo serviço com CancellationToken

#### 2. **TokenProvider.cs**
- **Responsabilidade**: Autenticação OAuth2 com Microsoft Identity
- **Fluxo**:
  ```
  ConfidentialClientApplication → AcquireTokenForClient → Access Token
  ```
- **Implementa**: `IAccessTokenProvider` (Kiota)

#### 3. **OutlookService.cs**
- **Responsabilidade**: Orquestração principal, timer, gestão de pastas
- **Métodos-chave**:
  - `CheckEmails()` - Ciclo de verificação periódica
  - `ProcessNextUnreadEmail()` - Busca próximo email não lido
  - `EnsureProcessedFolderExists()` - Cria/valida pasta no Outlook

#### 4. **EmailProcessor.cs**
- **Responsabilidade**: Processamento completo de emails
- **Métodos-chave**:
  - `ProcessEmail()` - Loop principal com retry (linhas 345-417)
  - `ProcessarAnexos()` - Download de anexos (linhas 419-491)
  - `CreateEmailBodyPdf()` - Geração PDF com iText7 (linhas 104-172)
  - `CreateXmlFile()` - Criação XML estruturado (linhas 208-290)
  - `SaveToDatabase()` - Insert no Oracle (linhas 493-544)
  - `MoveToProcessedFolder()` - Movimentação via Graph API (linhas 546-601)

### Dependências NuGet

| Pacote | Versão | Uso |
|--------|--------|-----|
| `Microsoft.Graph` | 5.36.0 | Client SDK para Microsoft Graph API |
| `Microsoft.Identity.Client` | 4.66.2 | MSAL para autenticação OAuth2 |
| `itext7` | 9.0.0 | Geração de documentos PDF |
| `itext7.pdfhtml` | 6.0.0 | Conversão HTML para PDF |
| `Oracle.ManagedDataAccess.Core` | 3.21.120 | Driver Oracle gerenciado |
| `System.Data.OleDb` | 7.0.0 | Acesso a dados via OLEDB |
| `Serilog` | 3.1.1 | Framework de logging estruturado |
| `Serilog.Sinks.Console` | 5.0.1 | Output para console |
| `Serilog.Sinks.File` | 5.0.0 | Output para arquivo com rolling |

### Mecanismos de Resiliência

#### Retry Pattern
```csharp
// EmailProcessor.cs:345-417
while (retryCount < _maxRetries) {
    try {
        // Processar email
        return;
    } catch (Exception ex) {
        retryCount++;
        if (retryCount >= _maxRetries) {
            await MoveToErrorFolder(userEmail, message.Id);
            throw;
        }
        await Task.Delay(500 * retryCount); // Delay progressivo
    }
}
```

#### Validação de Arquivos
```csharp
// EmailProcessor.cs:38-73
private async Task<bool> ValidateFileCreation(string filePath, int maxRetries = 3) {
    // Tenta abrir arquivo com FileShare.None
    // Retry com delay se bloqueado
    // Retorna false se falhar após todas as tentativas
}
```

---

## Changelog

### Versão 1.2.0 (2025-09-30)
- ✨ **NOVO**: Suporte completo para **Shared Mailboxes**
  - Health check adaptado para testar acesso a mensagens diretamente
  - Funciona com mesmas permissões de mailboxes normais
  - Configuração transparente: basta adicionar email da shared mailbox
- ✨ **NOVO**: Suporte a pastas de entrada personalizadas **por email**
  - Nova configuração `InboxFolder` permite especificar pasta customizada
  - **Modo 1**: Pasta única para todos os emails: `<add key="InboxFolder" value="Processos" />`
  - **Modo 2**: Pasta diferente por email: `<add key="InboxFolder" value="email1@domain.com:testes;email2@domain.com:teste1" />`
  - Por padrão usa "Inbox" (Caixa de Entrada) para emails não configurados
  - Detecção automática e cache do ID da pasta para performance
  - Logs informativos mostram mapeamento email→pasta
  - Validação proativa ao iniciar: verifica se todas as pastas configuradas existem
  - Lista automaticamente pastas disponíveis se não encontrar a configurada
- ✨ **NOVO**: Editor de Configurações no GUI
  - Interface completa com abas para todas as configurações
  - **Dropdown inteligente para InboxFolder**: botão "Listar Pastas" carrega pastas disponíveis do Outlook
  - Edição de todas as configurações: Azure AD, Emails, Serviço, Pastas, Logs, Base de Dados
  - Guarda diretamente no App.config (formato XML)
  - Validação antes de guardar
  - PropertyGrid read-only para visualização rápida
- 🐛 **FIX**: Corrigidas 2 warnings de nullability em `EmailProcessor.cs` (linhas 286, 557)
  - Adicionado `.Cast<string>()` após filtro de destinatários
  - Resolve incompatibilidade entre `List<string>` e `List<string?>`
- 🐛 **FIX**: Corrigida validação de paths UNC em `InputValidator.cs`
  - Agora aceita corretamente caminhos de rede como `\\servidor\share\pasta\`
  - Validação de path traversal ajustada para permitir `\\` no início
  - Remove falsos positivos na detecção de caracteres suspeitos
- 🐛 **FIX**: Corrigido health check de base de dados em `HealthCheckService.cs`
  - Query `SELECT 1 FROM DUAL` substituída por detecção automática de provider
  - SQL Server agora usa `SELECT 1` corretamente
  - Elimina erro "invalid object name duas" em SQL Server

### Versão 1.1.0 (2025-01-11)
- ✨ **NOVO**: Filtragem automática de destinatários no campo `<to>` do XML
  - Filtra apenas emails configurados em `UserEmail`
  - Remove destinatários externos automaticamente
  - Ideal para emails com múltiplos destinatários
- ✨ **NOVO**: Suporte multi-database com detecção automática
  - Oracle Database (OraOLEDB)
  - SQL Server (SQLOLEDB, SQLNCLI, MSOLEDBSQL)
  - Microsoft Access (ACE, JET)
  - Detecção automática do provider pela connection string
  - Queries SQL adaptadas automaticamente
- 🔧 Melhorias no logging de destinatários filtrados
- 🔧 Validação de valores null em campos de banco de dados

### Versão 1.0.0 (2024-09-20)
- ✅ Implementação inicial
- ✅ Suporte a Microsoft Graph API
- ✅ Processamento de anexos
- ✅ Geração de PDF e XML
- ✅ Persistência em Oracle
- ✅ Sistema de logging
- ✅ Modo CLI e GUI
- ✅ Mecanismo de retry

---

**Desenvolvido em C# .NET 9.0**
