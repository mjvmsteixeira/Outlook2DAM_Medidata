# Scripts PowerShell - Outlook2DAM

Este documento descreve os scripts PowerShell disponíveis para automação de instalação, configuração e testes do Outlook2DAM.

## 📋 Índice

- [Test-Outlook2DAM.ps1](#test-outlook2damps1) - Script de validação e testes
- [Install-Outlook2DAM.ps1](#install-outlook2damps1) - Script de instalação e configuração

---

## 🧪 Test-Outlook2DAM.ps1

Script de validação que executa 10 verificações críticas antes do deployment.

### Uso

```powershell
# Execução simples (modo padrão)
.\Test-Outlook2DAM.ps1

# Especificar caminho do App.config
.\Test-Outlook2DAM.ps1 -ConfigPath "C:\Custom\Path\App.config"
```

### Verificações Executadas

| # | Teste | Descrição |
|---|-------|-----------|
| 1 | SDK .NET | Verifica se .NET 9.0 SDK está instalado |
| 2 | Ficheiros do projeto | Valida presença de ficheiros essenciais |
| 3 | App.config | Verifica se todas as configurações obrigatórias estão preenchidas |
| 4 | Azure AD (GUIDs) | Valida formato dos GUIDs de TenantId e ClientId |
| 5 | Emails | Valida formato dos emails configurados |
| 6 | TempFolder | Testa permissões de leitura/escrita na pasta temporária |
| 7 | Connection String | Detecta provider de base de dados |
| 8 | Conectividade | Testa acesso a graph.microsoft.com |
| 9 | NuGet | Restaura e valida pacotes |
| 10 | Compilação | Tenta compilar o projeto |

### Saída de Exemplo

```
═══════════════════════════════════════════════
  Outlook2DAM - Teste de Validação
═══════════════════════════════════════════════

A executar testes...

  [1/10] .NET SDK... ✓
  [2/10] Ficheiros do projeto... ✓
  [3/10] Configuração (App.config)... ✓
  [4/10] Azure AD (GUIDs)... ✓
  [5/10] Validação de emails... ✓
  [6/10] TempFolder... ✓
  [7/10] Connection String... ✓
  [8/10] Conectividade Internet... ✓
  [9/10] Pacotes NuGet... ✓
  [10/10] Compilação... ✓

═══════════════════════════════════════════════
  Resumo dos Testes
═══════════════════════════════════════════════

  ✓ SDK .NET: Versão 9.0.1 instalada
  ✓ Ficheiros do projeto: Todos os ficheiros encontrados
  ✓ App.config: Todas as configurações obrigatórias preenchidas
  ✓ Azure AD: TenantId e ClientId são GUIDs válidos
  ✓ Emails: 2 email(s) válido(s)
  ✓ TempFolder: Pasta acessível com permissões de escrita
    → \\servidor\share\Outlook2DAM
  ✓ Connection String: Provider detectado: Oracle
  ✓ Conectividade: Acesso a graph.microsoft.com OK
  ✓ NuGet: Todos os pacotes restaurados
  ✓ Compilação: Projeto compila sem erros

═══════════════════════════════════════════════
  Resultado: 10/10 testes passaram
═══════════════════════════════════════════════

✓ Sistema pronto para deployment!

Próximos passos:
  1. Execute: .\Install-Outlook2DAM.ps1 -Mode Service
  2. Ou execute: .\Install-Outlook2DAM.ps1 para menu interativo
```

### Códigos de Saída

- `0` - Todos os testes passaram
- `1` - Um ou mais testes falharam

---

## ⚙️ Install-Outlook2DAM.ps1

Script completo de instalação e configuração com suporte a modo interativo e linha de comandos.

### Uso

#### Modo Interativo (Menu)

```powershell
# Executar sem parâmetros para ver o menu
.\Install-Outlook2DAM.ps1
```

**Menu disponível:**
```
Selecione uma opção:
  1. Configurar App.config
  2. Compilar projeto
  3. Instalar como Windows Service
  4. Instalação completa (Config + Build + Service)
  5. Desinstalar Windows Service
  6. Health Check
  0. Sair
```

#### Modo Linha de Comandos

```powershell
# Apenas configurar App.config interativamente
.\Install-Outlook2DAM.ps1 -Mode Config

# Compilar para execução standalone
.\Install-Outlook2DAM.ps1 -Mode Standalone

# Instalar como Windows Service (requer Admin)
.\Install-Outlook2DAM.ps1 -Mode Service -InstallPath "C:\Services\Outlook2DAM"

# Desinstalar Windows Service
.\Install-Outlook2DAM.ps1 -Mode Uninstall

# Executar health check
.\Install-Outlook2DAM.ps1 -Mode HealthCheck
```

### Parâmetros

| Parâmetro | Tipo | Default | Descrição |
|-----------|------|---------|-----------|
| `-Mode` | String | `Config` | Modo de operação: `Service`, `Standalone`, `Config`, `Uninstall`, `HealthCheck` |
| `-ServiceName` | String | `Outlook2DAM` | Nome do Windows Service |
| `-InstallPath` | String | `C:\Services\Outlook2DAM` | Diretório de instalação |

### Fluxo de Instalação Completa

**Passo 1: Validação de Pré-requisitos**
- Verifica instalação do .NET SDK 9.0
- Valida presença dos ficheiros do projeto
- Verifica se está no diretório correto

**Passo 2: Configuração Interativa**
- Solicita credenciais Azure AD (TenantId, ClientId, ClientSecret)
- Configura emails a monitorizar
- Define TempFolder (cria se não existir)
- Configura connection string da base de dados
- Suporta Oracle, SQL Server e Access

**Passo 3: Compilação**
- Executa `dotnet restore`
- Compila em modo Release
- Publica aplicação na pasta `publish`

**Passo 4: Instalação do Serviço**
- Copia ficheiros para `InstallPath`
- Cria Windows Service via `sc.exe`
- Configura política de recuperação automática
- Inicia o serviço

### Exemplos de Uso

#### Exemplo 1: Instalação Completa Standalone

```powershell
# 1. Validar sistema primeiro
.\Test-Outlook2DAM.ps1

# 2. Compilar para standalone
.\Install-Outlook2DAM.ps1 -Mode Standalone

# 3. Executar manualmente
.\publish\Outlook2DAM.exe --cli
```

#### Exemplo 2: Instalação como Windows Service

```powershell
# Requer PowerShell como Administrador
.\Install-Outlook2DAM.ps1 -Mode Service -InstallPath "D:\Services\Outlook2DAM"

# Verificar status
sc.exe query Outlook2DAM

# Ver logs
Get-Content D:\Services\Outlook2DAM\logs\outlook2dam-*.log -Wait -Tail 50
```

#### Exemplo 3: Workflow Completo de Desenvolvimento para Produção

```powershell
# 1. Configurar ambiente de desenvolvimento
.\Install-Outlook2DAM.ps1 -Mode Config

# 2. Testar configuração
.\Test-Outlook2DAM.ps1

# 3. Se testes passarem, compilar
.\Install-Outlook2DAM.ps1 -Mode Standalone

# 4. Testar localmente
.\publish\Outlook2DAM.exe --cli
# (Ctrl+C para parar)

# 5. Instalar como serviço em produção
# (Execute como Administrador)
.\Install-Outlook2DAM.ps1 -Mode Service

# 6. Monitorizar logs
Get-Content C:\Services\Outlook2DAM\logs\outlook2dam-*.log -Wait -Tail 50
```

#### Exemplo 4: Atualização de Versão

```powershell
# 1. Parar serviço existente
sc.exe stop Outlook2DAM

# 2. Fazer backup da configuração
Copy-Item "C:\Services\Outlook2DAM\App.config" "C:\Backup\App.config.bak"

# 3. Recompilar
.\Install-Outlook2DAM.ps1 -Mode Standalone

# 4. Copiar novos binários (preservando App.config)
Copy-Item ".\publish\*" "C:\Services\Outlook2DAM\" -Exclude "App.config" -Force

# 5. Reiniciar serviço
sc.exe start Outlook2DAM
```

### Configuração Guiada de Base de Dados

Durante o modo `Config`, o script oferece configuração guiada para diferentes providers:

**Oracle:**
```
Tipo (Oracle/SqlServer/Access): Oracle
  Host: db.empresa.com
  Port (default: 1521): 1521
  Service Name: ORCL
  Username: outlook2dam
  Password: ********
```
Gera: `Provider=OraOLEDB.Oracle;Data Source=...;User Id=...;Password=...`

**SQL Server:**
```
Tipo (Oracle/SqlServer/Access): SqlServer
  Server: sql.empresa.com
  Database: Outlook2DAM
  Username: sa
  Password: ********
```
Gera: `Provider=SQLOLEDB;Data Source=...;Initial Catalog=...;User ID=...;Password=...`

**Access:**
```
Tipo (Oracle/SqlServer/Access): Access
  Caminho do ficheiro .mdb/.accdb: C:\Data\Outlook2DAM.accdb
```
Gera: `Provider=Microsoft.ACE.OLEDB.12.0;Data Source=...`

### Resolução de Problemas

#### Erro: "Script não assinado"
```powershell
# Permitir execução de scripts (Execute como Admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Erro: "Requer privilégios de Administrador"
```powershell
# Executar PowerShell como Admin
Start-Process powershell -Verb runAs -ArgumentList "-File .\Install-Outlook2DAM.ps1 -Mode Service"
```

#### Erro: "dotnet não encontrado"
```powershell
# Instalar .NET 9.0 SDK
winget install Microsoft.DotNet.SDK.9
# Reiniciar PowerShell após instalação
```

#### Serviço não inicia
```powershell
# Verificar logs do Event Viewer
Get-EventLog -LogName Application -Source Outlook2DAM -Newest 10

# Verificar logs da aplicação
Get-Content C:\Services\Outlook2DAM\logs\outlook2dam-*.log -Tail 50

# Testar manualmente
cd C:\Services\Outlook2DAM
.\Outlook2DAM.exe --cli
```

### Segurança

⚠️ **Avisos de Segurança:**

1. **ClientSecret**: Introduzido como `SecureString` para não aparecer em plaintext durante digitação
2. **App.config**: Ficheiro contém credenciais - proteja com permissões NTFS adequadas
3. **Logs**: Credenciais são automaticamente mascaradas nos logs (implementado na v1.2.0)
4. **Connection Strings**: Passwords são sanitizadas antes de logging

**Recomendações:**
```powershell
# Restringir permissões do App.config (apenas SYSTEM e Admins)
$acl = Get-Acl "C:\Services\Outlook2DAM\App.config"
$acl.SetAccessRuleProtection($true, $false)
$acl.Access | ForEach-Object { $acl.RemoveAccessRule($_) }
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM", "FullControl", "Allow")))
$acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("Administrators", "FullControl", "Allow")))
Set-Acl "C:\Services\Outlook2DAM\App.config" $acl
```

---

## 🔄 Workflows Recomendados

### Desenvolvimento
```powershell
# 1. Configurar
.\Install-Outlook2DAM.ps1 -Mode Config

# 2. Testar
.\Test-Outlook2DAM.ps1

# 3. Executar em modo debug via Visual Studio
# (F5 no Visual Studio)
```

### Staging
```powershell
# 1. Validar
.\Test-Outlook2DAM.ps1

# 2. Compilar
.\Install-Outlook2DAM.ps1 -Mode Standalone

# 3. Testar standalone
.\publish\Outlook2DAM.exe --cli
```

### Produção
```powershell
# 1. Validar (em ambiente de staging)
.\Test-Outlook2DAM.ps1

# 2. Instalar como serviço (em produção, como Admin)
.\Install-Outlook2DAM.ps1 -Mode Service -InstallPath "D:\Production\Outlook2DAM"

# 3. Monitorizar
Get-Content D:\Production\Outlook2DAM\logs\*.log -Wait -Tail 50
```

---

## 📚 Referências

- [README.md](README.md) - Documentação principal do projeto
- [CONFIGURACAO.md](Outlook2DAM/CONFIGURACAO.md) - Guia detalhado de configuração
- [App-default.config](Outlook2DAM/App-default.config) - Template de configuração

---

**Nota**: Estes scripts requerem PowerShell 5.1 ou superior e .NET 9.0 SDK instalado.