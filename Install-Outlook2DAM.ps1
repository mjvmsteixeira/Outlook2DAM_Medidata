<#
.SYNOPSIS
    Script de instalação e configuração do Outlook2DAM

.DESCRIPTION
    Este script automatiza a instalação, configuração e deployment do serviço Outlook2DAM.
    Permite instalar como Windows Service ou executar em modo standalone.

.PARAMETER Mode
    Modo de instalação: Service, Standalone, Config, Uninstall

.PARAMETER ServiceName
    Nome do Windows Service (default: Outlook2DAM)

.PARAMETER InstallPath
    Caminho de instalação (default: C:\Services\Outlook2DAM)

.EXAMPLE
    .\Install-Outlook2DAM.ps1 -Mode Config
    Apenas configura o App.config interativamente

.EXAMPLE
    .\Install-Outlook2DAM.ps1 -Mode Service -InstallPath "C:\MyServices\Outlook2DAM"
    Compila, instala como Windows Service no caminho especificado

.EXAMPLE
    .\Install-Outlook2DAM.ps1 -Mode Standalone
    Compila e prepara para execução standalone

.EXAMPLE
    .\Install-Outlook2DAM.ps1 -Mode Uninstall
    Remove o Windows Service
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Service', 'Standalone', 'Config', 'Uninstall', 'HealthCheck')]
    [string]$Mode = 'Config',

    [Parameter(Mandatory=$false)]
    [string]$ServiceName = 'Outlook2DAM',

    [Parameter(Mandatory=$false)]
    [string]$InstallPath = 'C:\Services\Outlook2DAM'
)

# Requer privilégios de administrador para instalação de serviço
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Banner
function Show-Banner {
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "  Outlook2DAM - Instalador e Configurador" -ForegroundColor Cyan
    Write-Host "  Versão 1.2.0" -ForegroundColor Cyan
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""
}

# Verificar pré-requisitos
function Test-Prerequisites {
    Write-Host "[1/4] A verificar pré-requisitos..." -ForegroundColor Yellow

    # Verificar .NET 9.0
    try {
        $dotnetVersion = dotnet --version
        Write-Host "  ✓ .NET SDK instalado: $dotnetVersion" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ .NET SDK 9.0 não encontrado!" -ForegroundColor Red
        Write-Host "    Descarregue em: https://dotnet.microsoft.com/download/dotnet/9.0" -ForegroundColor Red
        return $false
    }

    # Verificar se estamos no diretório correto
    $csprojPath = Join-Path $PSScriptRoot "Outlook2DAM\Outlook2DAM.csproj"
    if (-not (Test-Path $csprojPath)) {
        Write-Host "  ✗ Projeto não encontrado em: $csprojPath" -ForegroundColor Red
        return $false
    }
    Write-Host "  ✓ Projeto encontrado" -ForegroundColor Green

    return $true
}

# Configurar App.config interativamente
function Set-Configuration {
    Write-Host "[2/4] A configurar aplicação..." -ForegroundColor Yellow

    $configPath = Join-Path $PSScriptRoot "Outlook2DAM\App.config"
    $defaultConfigPath = Join-Path $PSScriptRoot "Outlook2DAM\App-default.config"

    # Se não existe App.config, copiar do template
    if (-not (Test-Path $configPath)) {
        if (Test-Path $defaultConfigPath) {
            Copy-Item $defaultConfigPath $configPath
            Write-Host "  ✓ App.config criado a partir do template" -ForegroundColor Green
        }
        else {
            Write-Host "  ✗ Template App-default.config não encontrado!" -ForegroundColor Red
            return $false
        }
    }

    Write-Host ""
    Write-Host "  Configuração da aplicação:" -ForegroundColor Cyan
    Write-Host "  (Pressione Enter para manter o valor atual)" -ForegroundColor Gray
    Write-Host ""

    # Ler App.config atual
    [xml]$config = Get-Content $configPath

    # Azure AD / Microsoft Graph
    $tenantId = Read-Host "  TenantId (Azure AD)"
    if ($tenantId) {
        $config.configuration.appSettings.add | Where-Object { $_.key -eq 'TenantId' } | ForEach-Object { $_.value = $tenantId }
    }

    $clientId = Read-Host "  ClientId (App Registration)"
    if ($clientId) {
        $config.configuration.appSettings.add | Where-Object { $_.key -eq 'ClientId' } | ForEach-Object { $_.value = $clientId }
    }

    $clientSecret = Read-Host "  ClientSecret (App Registration)" -AsSecureString
    if ($clientSecret.Length -gt 0) {
        $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($clientSecret)
        $plainSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
        $config.configuration.appSettings.add | Where-Object { $_.key -eq 'ClientSecret' } | ForEach-Object { $_.value = $plainSecret }
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }

    $userEmail = Read-Host "  UserEmail (emails separados por ;)"
    if ($userEmail) {
        $config.configuration.appSettings.add | Where-Object { $_.key -eq 'UserEmail' } | ForEach-Object { $_.value = $userEmail }
    }

    # Configurações de sistema
    $tempFolder = Read-Host "  TempFolder (caminho temporário)"
    if ($tempFolder) {
        $config.configuration.appSettings.add | Where-Object { $_.key -eq 'TempFolder' } | ForEach-Object { $_.value = $tempFolder }

        # Criar diretório se não existir
        if (-not (Test-Path $tempFolder)) {
            New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
            Write-Host "    ✓ Diretório criado: $tempFolder" -ForegroundColor Green
        }
    }

    # Connection String
    Write-Host ""
    Write-Host "  Base de Dados:" -ForegroundColor Cyan
    $dbType = Read-Host "  Tipo (Oracle/SqlServer/Access) [Enter para não alterar]"

    if ($dbType) {
        $connectionString = ""
        switch ($dbType.ToLower()) {
            "oracle" {
                $host = Read-Host "    Host"
                $port = Read-Host "    Port (default: 1521)"
                if (-not $port) { $port = "1521" }
                $service = Read-Host "    Service Name"
                $user = Read-Host "    Username"
                $password = Read-Host "    Password" -AsSecureString
                $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
                $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
                $connectionString = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$host)(PORT=$port))(CONNECT_DATA=(SERVICE_NAME=$service)));User Id=$user;Password=$plainPassword;"
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            }
            "sqlserver" {
                $server = Read-Host "    Server"
                $database = Read-Host "    Database"
                $user = Read-Host "    Username"
                $password = Read-Host "    Password" -AsSecureString
                $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
                $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
                $connectionString = "Provider=SQLOLEDB;Data Source=$server;Initial Catalog=$database;User ID=$user;Password=$plainPassword;"
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            }
            "access" {
                $dbPath = Read-Host "    Caminho do ficheiro .mdb/.accdb"
                $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
            }
        }

        if ($connectionString) {
            $connectionStringNode = $config.configuration.connectionStrings.add | Where-Object { $_.name -eq 'Outlook2DAM' }
            if ($connectionStringNode) {
                $connectionStringNode.connectionString = $connectionString
            }
        }
    }

    # Guardar configuração
    $config.Save($configPath)
    Write-Host ""
    Write-Host "  ✓ Configuração guardada em: $configPath" -ForegroundColor Green

    return $true
}

# Compilar projeto
function Build-Project {
    Write-Host "[3/4] A compilar projeto..." -ForegroundColor Yellow

    $projectPath = Join-Path $PSScriptRoot "Outlook2DAM"
    Push-Location $projectPath

    try {
        # Restore packages
        Write-Host "  → Restaurando pacotes NuGet..." -ForegroundColor Gray
        dotnet restore
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  ✗ Erro ao restaurar pacotes" -ForegroundColor Red
            return $false
        }

        # Build
        Write-Host "  → Compilando em Release..." -ForegroundColor Gray
        dotnet build -c Release --no-restore
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  ✗ Erro ao compilar" -ForegroundColor Red
            return $false
        }

        # Publish
        Write-Host "  → Publicando aplicação..." -ForegroundColor Gray
        dotnet publish -c Release -o "$PSScriptRoot\publish" --no-build
        if ($LASTEXITCODE -ne 0) {
            Write-Host "  ✗ Erro ao publicar" -ForegroundColor Red
            return $false
        }

        Write-Host "  ✓ Compilação concluída com sucesso" -ForegroundColor Green
        return $true
    }
    finally {
        Pop-Location
    }
}

# Instalar como Windows Service
function Install-WindowsService {
    Write-Host "[4/4] A instalar Windows Service..." -ForegroundColor Yellow

    if (-not (Test-Administrator)) {
        Write-Host "  ✗ Este passo requer privilégios de Administrador!" -ForegroundColor Red
        Write-Host "    Execute: Start-Process powershell -Verb runAs -ArgumentList '-File $PSCommandPath -Mode Service'" -ForegroundColor Yellow
        return $false
    }

    # Criar diretório de instalação
    if (-not (Test-Path $InstallPath)) {
        New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
        Write-Host "  ✓ Diretório de instalação criado: $InstallPath" -ForegroundColor Green
    }

    # Copiar ficheiros publicados
    $publishPath = Join-Path $PSScriptRoot "publish"
    Copy-Item "$publishPath\*" $InstallPath -Recurse -Force
    Write-Host "  ✓ Ficheiros copiados para: $InstallPath" -ForegroundColor Green

    # Verificar se serviço já existe
    $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($existingService) {
        Write-Host "  → Serviço já existe, a parar..." -ForegroundColor Gray
        Stop-Service -Name $ServiceName -Force
        Write-Host "  → A remover serviço existente..." -ForegroundColor Gray
        sc.exe delete $ServiceName | Out-Null
        Start-Sleep -Seconds 2
    }

    # Criar serviço Windows
    $exePath = Join-Path $InstallPath "Outlook2DAM.exe"
    $serviceArgs = "--cli"

    Write-Host "  → A criar Windows Service..." -ForegroundColor Gray
    sc.exe create $ServiceName binPath= "`"$exePath`" $serviceArgs" start= auto DisplayName= "Outlook2DAM Email Processor" | Out-Null

    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ Windows Service criado: $ServiceName" -ForegroundColor Green

        # Configurar recuperação em caso de falha
        sc.exe failure $ServiceName reset= 86400 actions= restart/60000/restart/60000/restart/60000 | Out-Null
        Write-Host "  ✓ Política de recuperação configurada" -ForegroundColor Green

        # Iniciar serviço
        Write-Host "  → A iniciar serviço..." -ForegroundColor Gray
        Start-Service -Name $ServiceName
        Start-Sleep -Seconds 2

        $service = Get-Service -Name $ServiceName
        if ($service.Status -eq 'Running') {
            Write-Host "  ✓ Serviço iniciado com sucesso!" -ForegroundColor Green
        }
        else {
            Write-Host "  ⚠ Serviço criado mas não está em execução" -ForegroundColor Yellow
            Write-Host "    Verifique os logs em: $InstallPath\Logs" -ForegroundColor Yellow
        }

        return $true
    }
    else {
        Write-Host "  ✗ Erro ao criar Windows Service" -ForegroundColor Red
        return $false
    }
}

# Desinstalar Windows Service
function Uninstall-WindowsService {
    Write-Host "[Uninstall] A remover Windows Service..." -ForegroundColor Yellow

    if (-not (Test-Administrator)) {
        Write-Host "  ✗ Este passo requer privilégios de Administrador!" -ForegroundColor Red
        return $false
    }

    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        Write-Host "  → A parar serviço..." -ForegroundColor Gray
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        Write-Host "  → A remover serviço..." -ForegroundColor Gray
        sc.exe delete $ServiceName | Out-Null

        if ($LASTEXITCODE -eq 0) {
            Write-Host "  ✓ Serviço removido com sucesso" -ForegroundColor Green
        }
        else {
            Write-Host "  ✗ Erro ao remover serviço" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "  ⚠ Serviço não encontrado: $ServiceName" -ForegroundColor Yellow
    }

    # Remover ficheiros
    if (Test-Path $InstallPath) {
        $confirm = Read-Host "  Remover ficheiros de $InstallPath? (S/N)"
        if ($confirm -eq 'S' -or $confirm -eq 's') {
            Remove-Item $InstallPath -Recurse -Force
            Write-Host "  ✓ Ficheiros removidos" -ForegroundColor Green
        }
    }

    return $true
}

# Health Check
function Invoke-HealthCheck {
    Write-Host "[Health Check] A verificar sistema..." -ForegroundColor Yellow

    $exePath = Join-Path $PSScriptRoot "publish\Outlook2DAM.exe"
    if (-not (Test-Path $exePath)) {
        # Tentar no InstallPath
        $exePath = Join-Path $InstallPath "Outlook2DAM.exe"
    }

    if (-not (Test-Path $exePath)) {
        Write-Host "  ✗ Executável não encontrado. Execute primeiro o build." -ForegroundColor Red
        return $false
    }

    # Executar aplicação com flag de health check (seria necessário implementar)
    Write-Host "  → A executar verificações..." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Para verificar manualmente:" -ForegroundColor Cyan
    Write-Host "  1. Logs: " -NoNewline; Write-Host "$InstallPath\Logs" -ForegroundColor White
    Write-Host "  2. Config: " -NoNewline; Write-Host "$InstallPath\App.config" -ForegroundColor White
    Write-Host "  3. Serviço: " -NoNewline

    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        Write-Host $service.Status -ForegroundColor $(if ($service.Status -eq 'Running') { 'Green' } else { 'Red' })
    }
    else {
        Write-Host "Não instalado" -ForegroundColor Yellow
    }

    return $true
}

# Menu principal
function Show-Menu {
    Write-Host ""
    Write-Host "Selecione uma opção:" -ForegroundColor Cyan
    Write-Host "  1. Configurar App.config" -ForegroundColor White
    Write-Host "  2. Compilar projeto" -ForegroundColor White
    Write-Host "  3. Instalar como Windows Service" -ForegroundColor White
    Write-Host "  4. Instalação completa (Config + Build + Service)" -ForegroundColor White
    Write-Host "  5. Desinstalar Windows Service" -ForegroundColor White
    Write-Host "  6. Health Check" -ForegroundColor White
    Write-Host "  0. Sair" -ForegroundColor White
    Write-Host ""

    $choice = Read-Host "Opção"
    return $choice
}

# Main
Show-Banner

# Se modo não especificado, mostrar menu
if ($PSBoundParameters.Count -eq 0) {
    do {
        $choice = Show-Menu
        Write-Host ""

        switch ($choice) {
            "1" { Set-Configuration }
            "2" {
                if (Test-Prerequisites) {
                    Build-Project
                }
            }
            "3" {
                if (Test-Administrator) {
                    Install-WindowsService
                }
                else {
                    Write-Host "✗ Requer privilégios de Administrador" -ForegroundColor Red
                }
            }
            "4" {
                if (Test-Prerequisites) {
                    Set-Configuration
                    Build-Project
                    if (Test-Administrator) {
                        Install-WindowsService
                    }
                    else {
                        Write-Host "⚠ Compilação concluída, mas instalação de serviço requer privilégios de Administrador" -ForegroundColor Yellow
                    }
                }
            }
            "5" {
                if (Test-Administrator) {
                    Uninstall-WindowsService
                }
                else {
                    Write-Host "✗ Requer privilégios de Administrador" -ForegroundColor Red
                }
            }
            "6" { Invoke-HealthCheck }
            "0" { Write-Host "Saindo..." -ForegroundColor Gray }
            default { Write-Host "Opção inválida" -ForegroundColor Red }
        }

        if ($choice -ne "0") {
            Write-Host ""
            Read-Host "Pressione Enter para continuar"
        }
    } while ($choice -ne "0")
}
else {
    # Modo especificado via parâmetro
    switch ($Mode) {
        "Config" {
            Set-Configuration
        }
        "Standalone" {
            if (Test-Prerequisites) {
                Build-Project
                Write-Host ""
                Write-Host "✓ Build concluído. Execute com:" -ForegroundColor Green
                Write-Host "  .\publish\Outlook2DAM.exe --cli" -ForegroundColor White
            }
        }
        "Service" {
            if (Test-Prerequisites) {
                Build-Project
                Install-WindowsService
            }
        }
        "Uninstall" {
            Uninstall-WindowsService
        }
        "HealthCheck" {
            Invoke-HealthCheck
        }
    }
}

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan