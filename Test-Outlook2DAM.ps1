<#
.SYNOPSIS
    Script de testes e validação do Outlook2DAM

.DESCRIPTION
    Valida configuração, conectividade e dependências antes de executar o serviço

.EXAMPLE
    .\Test-Outlook2DAM.ps1
    Executa todos os testes de validação
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = ".\Outlook2DAM\App.config"
)

$ErrorActionPreference = "Continue"
$testResults = @()

# Classe para resultado de teste
class TestResult {
    [string]$TestName
    [bool]$Passed
    [string]$Message
    [string]$Details
}

function New-TestResult {
    param(
        [string]$TestName,
        [bool]$Passed,
        [string]$Message,
        [string]$Details = ""
    )

    $result = [TestResult]::new()
    $result.TestName = $TestName
    $result.Passed = $Passed
    $result.Message = $Message
    $result.Details = $Details
    return $result
}

function Write-TestHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan
}

function Write-TestResult {
    param([TestResult]$Result)

    $icon = if ($Result.Passed) { "✓" } else { "✗" }
    $color = if ($Result.Passed) { "Green" } else { "Red" }

    Write-Host "  $icon " -NoNewline -ForegroundColor $color
    Write-Host "$($Result.TestName): " -NoNewline -ForegroundColor White
    Write-Host $Result.Message -ForegroundColor $color

    if ($Result.Details) {
        Write-Host "    → $($Result.Details)" -ForegroundColor Gray
    }
}

# Teste 1: Verificar .NET SDK
function Test-DotNetSDK {
    Write-Host "  [1/10] .NET SDK..." -NoNewline
    try {
        $version = dotnet --version 2>$null
        if ($version -and $version -match "^9\.") {
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "SDK .NET" $true "Versão $version instalada"
        }
        else {
            Write-Host " ✗" -ForegroundColor Red
            return New-TestResult "SDK .NET" $false "Versão 9.x necessária, encontrada: $version"
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "SDK .NET" $false "SDK não encontrado" $_.Exception.Message
    }
}

# Teste 2: Verificar ficheiros de projeto
function Test-ProjectFiles {
    Write-Host "  [2/10] Ficheiros do projeto..." -NoNewline

    $requiredFiles = @(
        ".\Outlook2DAM\Outlook2DAM.csproj",
        ".\Outlook2DAM\Program.cs",
        ".\Outlook2DAM\App-default.config"
    )

    $missingFiles = $requiredFiles | Where-Object { -not (Test-Path $_) }

    if ($missingFiles.Count -eq 0) {
        Write-Host " ✓" -ForegroundColor Green
        return New-TestResult "Ficheiros do projeto" $true "Todos os ficheiros encontrados"
    }
    else {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "Ficheiros do projeto" $false "Ficheiros em falta" ($missingFiles -join ", ")
    }
}

# Teste 3: Verificar App.config
function Test-Configuration {
    Write-Host "  [3/10] Configuração (App.config)..." -NoNewline

    if (-not (Test-Path $ConfigPath)) {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "App.config" $false "Ficheiro não encontrado: $ConfigPath"
    }

    try {
        [xml]$config = Get-Content $ConfigPath
        $appSettings = $config.configuration.appSettings.add

        $requiredSettings = @(
            "TenantId",
            "ClientId",
            "ClientSecret",
            "UserEmail",
            "TempFolder"
        )

        $emptySettings = @()
        foreach ($setting in $requiredSettings) {
            $value = ($appSettings | Where-Object { $_.key -eq $setting }).value
            if ([string]::IsNullOrWhiteSpace($value)) {
                $emptySettings += $setting
            }
        }

        if ($emptySettings.Count -eq 0) {
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "App.config" $true "Todas as configurações obrigatórias preenchidas"
        }
        else {
            Write-Host " ✗" -ForegroundColor Red
            return New-TestResult "App.config" $false "Configurações vazias" ($emptySettings -join ", ")
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "App.config" $false "Erro ao ler ficheiro" $_.Exception.Message
    }
}

# Teste 4: Validar GUIDs
function Test-AzureADConfiguration {
    Write-Host "  [4/10] Azure AD (GUIDs)..." -NoNewline

    try {
        [xml]$config = Get-Content $ConfigPath
        $appSettings = $config.configuration.appSettings.add

        $tenantId = ($appSettings | Where-Object { $_.key -eq 'TenantId' }).value
        $clientId = ($appSettings | Where-Object { $_.key -eq 'ClientId' }).value

        $tenantIdValid = $tenantId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        $clientIdValid = $clientId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

        if ($tenantIdValid -and $clientIdValid) {
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "Azure AD" $true "TenantId e ClientId são GUIDs válidos"
        }
        else {
            Write-Host " ✗" -ForegroundColor Red
            $invalid = @()
            if (-not $tenantIdValid) { $invalid += "TenantId" }
            if (-not $clientIdValid) { $invalid += "ClientId" }
            return New-TestResult "Azure AD" $false "GUIDs inválidos" ($invalid -join ", ")
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "Azure AD" $false "Erro na validação" $_.Exception.Message
    }
}

# Teste 5: Validar emails
function Test-EmailConfiguration {
    Write-Host "  [5/10] Validação de emails..." -NoNewline

    try {
        [xml]$config = Get-Content $ConfigPath
        $userEmail = ($config.configuration.appSettings.add | Where-Object { $_.key -eq 'UserEmail' }).value

        $emails = $userEmail -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        $emailRegex = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

        $invalidEmails = $emails | Where-Object { $_ -notmatch $emailRegex }

        if ($invalidEmails.Count -eq 0) {
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "Emails" $true "$($emails.Count) email(s) válido(s)"
        }
        else {
            Write-Host " ✗" -ForegroundColor Red
            return New-TestResult "Emails" $false "Emails inválidos" ($invalidEmails -join ", ")
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "Emails" $false "Erro na validação" $_.Exception.Message
    }
}

# Teste 6: Verificar TempFolder
function Test-TempFolder {
    Write-Host "  [6/10] TempFolder..." -NoNewline

    try {
        [xml]$config = Get-Content $ConfigPath
        $tempFolder = ($config.configuration.appSettings.add | Where-Object { $_.key -eq 'TempFolder' }).value

        if ([string]::IsNullOrWhiteSpace($tempFolder)) {
            Write-Host " ✗" -ForegroundColor Red
            return New-TestResult "TempFolder" $false "TempFolder não configurado"
        }

        # Tentar criar se não existir
        if (-not (Test-Path $tempFolder)) {
            New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
        }

        # Testar escrita
        $testFile = Join-Path $tempFolder "test_$(Get-Random).tmp"
        "test" | Out-File $testFile -Force
        Remove-Item $testFile -Force

        Write-Host " ✓" -ForegroundColor Green
        return New-TestResult "TempFolder" $true "Pasta acessível com permissões de escrita" $tempFolder
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "TempFolder" $false "Erro de acesso" $_.Exception.Message
    }
}

# Teste 7: Verificar Connection String
function Test-DatabaseConfiguration {
    Write-Host "  [7/10] Connection String..." -NoNewline

    try {
        [xml]$config = Get-Content $ConfigPath
        $connectionString = ($config.configuration.connectionStrings.add | Where-Object { $_.name -eq 'Outlook2DAM' }).connectionString

        if ([string]::IsNullOrWhiteSpace($connectionString)) {
            Write-Host " ⚠" -ForegroundColor Yellow
            return New-TestResult "Connection String" $true "Não configurada (opcional)"
        }

        # Detectar provider
        $provider = ""
        if ($connectionString -match "Provider=OraOLEDB") {
            $provider = "Oracle"
        }
        elseif ($connectionString -match "Provider=SQLOLEDB") {
            $provider = "SQL Server"
        }
        elseif ($connectionString -match "Provider=Microsoft\.ACE\.OLEDB") {
            $provider = "Access"
        }

        if ($provider) {
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "Connection String" $true "Provider detectado: $provider"
        }
        else {
            Write-Host " ⚠" -ForegroundColor Yellow
            return New-TestResult "Connection String" $true "Provider não identificado automaticamente"
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "Connection String" $false "Erro na leitura" $_.Exception.Message
    }
}

# Teste 8: Conectividade Internet
function Test-InternetConnectivity {
    Write-Host "  [8/10] Conectividade Internet..." -NoNewline

    try {
        $response = Invoke-WebRequest -Uri "https://graph.microsoft.com" -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop
        Write-Host " ✓" -ForegroundColor Green
        return New-TestResult "Conectividade" $true "Acesso a graph.microsoft.com OK"
    }
    catch {
        if ($_.Exception.Response.StatusCode -eq 401) {
            # 401 é esperado sem autenticação
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "Conectividade" $true "Acesso a graph.microsoft.com OK (401 esperado)"
        }
        Write-Host " ✗" -ForegroundColor Red
        return New-TestResult "Conectividade" $false "Sem acesso a graph.microsoft.com" $_.Exception.Message
    }
}

# Teste 9: Verificar NuGet packages
function Test-NuGetPackages {
    Write-Host "  [9/10] Pacotes NuGet..." -NoNewline

    try {
        Push-Location ".\Outlook2DAM"
        $restore = dotnet restore --verbosity quiet 2>&1
        Pop-Location

        if ($LASTEXITCODE -eq 0) {
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "NuGet" $true "Todos os pacotes restaurados"
        }
        else {
            Write-Host " ✗" -ForegroundColor Red
            return New-TestResult "NuGet" $false "Erro ao restaurar pacotes" ($restore -join "; ")
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        Pop-Location
        return New-TestResult "NuGet" $false "Erro ao restaurar" $_.Exception.Message
    }
}

# Teste 10: Compilação
function Test-Build {
    Write-Host "  [10/10] Compilação..." -NoNewline

    try {
        Push-Location ".\Outlook2DAM"
        $build = dotnet build -c Release --no-restore --verbosity quiet 2>&1
        Pop-Location

        if ($LASTEXITCODE -eq 0) {
            Write-Host " ✓" -ForegroundColor Green
            return New-TestResult "Compilação" $true "Projeto compila sem erros"
        }
        else {
            Write-Host " ✗" -ForegroundColor Red
            $errors = ($build | Select-String "error" | Select-Object -First 3) -join "; "
            return New-TestResult "Compilação" $false "Erros de compilação" $errors
        }
    }
    catch {
        Write-Host " ✗" -ForegroundColor Red
        Pop-Location
        return New-TestResult "Compilação" $false "Erro ao compilar" $_.Exception.Message
    }
}

# Execução principal
Write-Host ""
Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Outlook2DAM - Teste de Validação" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

Write-Host "A executar testes..." -ForegroundColor Yellow
Write-Host ""

$testResults += Test-DotNetSDK
$testResults += Test-ProjectFiles
$testResults += Test-Configuration
$testResults += Test-AzureADConfiguration
$testResults += Test-EmailConfiguration
$testResults += Test-TempFolder
$testResults += Test-DatabaseConfiguration
$testResults += Test-InternetConnectivity
$testResults += Test-NuGetPackages
$testResults += Test-Build

# Resumo
Write-Host ""
Write-TestHeader "Resumo dos Testes"
Write-Host ""

$passed = ($testResults | Where-Object { $_.Passed }).Count
$failed = ($testResults | Where-Object { -not $_.Passed }).Count
$total = $testResults.Count

foreach ($result in $testResults) {
    Write-TestResult $result
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Resultado: $passed/$total testes passaram" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })
Write-Host "═══════════════════════════════════════════════" -ForegroundColor Cyan

if ($failed -eq 0) {
    Write-Host ""
    Write-Host "✓ Sistema pronto para deployment!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Próximos passos:" -ForegroundColor Cyan
    Write-Host "  1. Execute: .\Install-Outlook2DAM.ps1 -Mode Service" -ForegroundColor White
    Write-Host "  2. Ou execute: .\Install-Outlook2DAM.ps1 para menu interativo" -ForegroundColor White
    exit 0
}
else {
    Write-Host ""
    Write-Host "⚠ Corrija os erros antes de continuar" -ForegroundColor Yellow
    exit 1
}