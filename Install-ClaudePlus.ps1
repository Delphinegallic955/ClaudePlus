# ============================================================================
# ClaudePlus - Installation Script
# Adds a PowerShell hook that enhances claude with Telegram mirror
# Author: Majid - FiscalIQ
# ============================================================================

param(
    [string]$TelegramBotToken = "",
    [string]$TelegramChatId = ""
)

$ErrorActionPreference = "Stop"

# --- 1. Load Telegram config: project folder first, then VS extension ---
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectConfigPath = Join-Path (Split-Path $scriptDir) "telegram-config.json"
$vsExtensionPath = "$env:LOCALAPPDATA\ClaudeCodeExtension\claudecode-settings.json"

# Source 1: telegram-config.json in parent project folder
if ((Test-Path $projectConfigPath) -and [string]::IsNullOrEmpty($TelegramBotToken)) {
    $telegramConfig = Get-Content $projectConfigPath -Raw | ConvertFrom-Json
    if ($telegramConfig.TelegramBotToken) {
        $TelegramBotToken = $telegramConfig.TelegramBotToken
        Write-Host "[OK] Telegram Bot Token recupere depuis telegram-config.json" -ForegroundColor Green
    }
    if ($telegramConfig.TelegramChatId) {
        $TelegramChatId = $telegramConfig.TelegramChatId
        Write-Host "[OK] Telegram Chat ID recupere depuis telegram-config.json" -ForegroundColor Green
    }
}

# Source 2: ClaudeCodeExtension VS settings (fallback)
if ((Test-Path $vsExtensionPath) -and [string]::IsNullOrEmpty($TelegramBotToken)) {
    $existingSettings = Get-Content $vsExtensionPath -Raw | ConvertFrom-Json
    if ($existingSettings.TelegramBotToken) {
        $TelegramBotToken = $existingSettings.TelegramBotToken
        Write-Host "[OK] Telegram Bot Token recupere depuis ClaudeCodeExtension" -ForegroundColor Green
    }
    if ($existingSettings.TelegramChatId) {
        $TelegramChatId = $existingSettings.TelegramChatId
        Write-Host "[OK] Telegram Chat ID recupere depuis ClaudeCodeExtension" -ForegroundColor Green
    }
}

# --- 2. Create config directory and save settings ---
$configDir = "$env:LOCALAPPDATA\ClaudePlus"
if (-not (Test-Path $configDir)) {
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
}

$config = @{
    TelegramBotToken = $TelegramBotToken
    TelegramChatId   = $TelegramChatId
    DangerouslySkipPermissions = $true
    AutoTelegram     = $true
}
$config | ConvertTo-Json | Set-Content "$configDir\config.json" -Encoding UTF8

# --- 3. Copy the module files ---
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = "$configDir\ClaudePlus.psm1"
Copy-Item "$scriptDir\ClaudePlus.psm1" $modulePath -Force
Write-Host "[OK] Module ClaudePlus copie vers $configDir" -ForegroundColor Green

# --- 4. Add to PowerShell profile ---
$profilePath = $PROFILE.CurrentUserAllHosts
if (-not (Test-Path $profilePath)) {
    New-Item -ItemType File -Path $profilePath -Force | Out-Null
}

$profileContent = Get-Content $profilePath -Raw -ErrorAction SilentlyContinue
$importLine = "Import-Module '$modulePath' -Force  # ClaudePlus"

if ($profileContent -notlike "*ClaudePlus*") {
    Add-Content $profilePath "`n$importLine"
    Write-Host "[OK] Module ajoute au profil PowerShell: $profilePath" -ForegroundColor Green
} else {
    Write-Host "[INFO] Module deja present dans le profil PowerShell" -ForegroundColor Yellow
}

# --- 5. Summary ---
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  ClaudePlus installe avec succes !" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Commandes disponibles dans TOUT terminal PowerShell :" -ForegroundColor White
Write-Host "  claudeplus              - Lance Claude Code + Telegram mirror" -ForegroundColor Gray
Write-Host "  claudeplus --no-telegram - Lance sans Telegram" -ForegroundColor Gray
Write-Host "  claudeplus-config       - Affiche/modifie la config" -ForegroundColor Gray
Write-Host ""
if ($TelegramBotToken -and $TelegramChatId) {
    Write-Host "  Telegram: CONFIGURE" -ForegroundColor Green
} else {
    Write-Host "  Telegram: NON CONFIGURE" -ForegroundColor Yellow
    Write-Host "  Pour configurer: claudeplus-config -TelegramBotToken 'TOKEN' -TelegramChatId 'ID'" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "Redemarrez votre terminal pour activer." -ForegroundColor Yellow
