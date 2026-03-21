# ============================================================================
# ClaudePlus PowerShell Module
# Wraps Claude Code with Telegram bidirectional mirror
# Author: Majid - FiscalIQ
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms

$script:ConfigPath = "$env:LOCALAPPDATA\ClaudePlus\config.json"
$script:ProjectTelegramConfig = "$([Environment]::GetFolderPath('MyDocuments'))\Visual Studio 2026\FiscalIQ\Claude Code Extension\telegram-config.json"
$script:VsExtensionConfig = "$env:LOCALAPPDATA\ClaudeCodeExtension\claudecode-settings.json"
$script:ClaudeProcess = $null
$script:ClaudeCmdPid = 0
$script:WshShell = New-Object -ComObject WScript.Shell

# No direct P/Invoke for console write - PowerShell shares its own console
# All WriteConsoleInput calls go through a helper process (see Send-TextToClaude)

# UI Automation for reading console text
try {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
} catch { }

$script:LastConsoleText = ""
$script:ConsoleReaderScript = $null

# ============================================================================
# READ CLAUDE CONSOLE OUTPUT (via helper process + AttachConsole)
# ============================================================================

function Initialize-ConsoleReader {
    # Create a helper PowerShell script that reads a console buffer
    $script:ConsoleReaderScript = "$env:TEMP\claudeplus_reader.ps1"

    $readerCode = @'
param([int]$TargetPid, [string]$OutputFile)
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class ConReader {
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool AttachConsole(int pid);
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool FreeConsole();
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll", SetLastError=true)]
    public static extern bool GetConsoleScreenBufferInfo(IntPtr h, out CSBI info);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern bool ReadConsoleOutputCharacter(IntPtr h, StringBuilder buf, int len, COORD coord, out int read);
    [StructLayout(LayoutKind.Sequential)]
    public struct COORD { public short X; public short Y; }
    [StructLayout(LayoutKind.Sequential)]
    public struct SMALL_RECT { public short Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    public struct CSBI {
        public COORD dwSize;
        public COORD dwCursorPosition;
        public short wAttributes;
        public SMALL_RECT srWindow;
        public COORD dwMaximumWindowSize;
    }
    public static string Read(int pid) {
        FreeConsole();
        if (!AttachConsole(pid)) return null;
        try {
            IntPtr h = GetStdHandle(-11);
            CSBI info;
            if (!GetConsoleScreenBufferInfo(h, out info)) return null;
            int w = info.dwSize.X;
            StringBuilder result = new StringBuilder();
            for (int y = info.srWindow.Top; y <= info.srWindow.Bottom; y++) {
                StringBuilder line = new StringBuilder(w);
                COORD c = new COORD { X = 0, Y = (short)y };
                int nr;
                ReadConsoleOutputCharacter(h, line, w, c, out nr);
                result.AppendLine(line.ToString().TrimEnd());
            }
            return result.ToString();
        } finally { FreeConsole(); }
    }
}
"@
try {
    $text = [ConReader]::Read($TargetPid)
    if ($text) { [System.IO.File]::WriteAllText($OutputFile, $text, [System.Text.Encoding]::UTF8) }
} catch { }
'@

    [System.IO.File]::WriteAllText($script:ConsoleReaderScript, $readerCode, [System.Text.Encoding]::ASCII)
}

function Read-ClaudeConsole {
    param([int]$CmdPid)

    if ($CmdPid -le 0) { return $null }
    if (-not $script:ConsoleReaderScript -or -not (Test-Path $script:ConsoleReaderScript)) { return $null }

    $outputFile = "$env:TEMP\claudeplus_console_$CmdPid.txt"

    # Run helper in hidden window (separate process = separate console)
    try {
        $proc = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$($script:ConsoleReaderScript)`" -TargetPid $CmdPid -OutputFile `"$outputFile`"" -PassThru -WindowStyle Hidden
        $proc.WaitForExit(5000) | Out-Null
        if (-not $proc.HasExited) { $proc.Kill() }
    } catch { return $null }

    if (Test-Path $outputFile) {
        try {
            $text = [System.IO.File]::ReadAllText($outputFile, [System.Text.Encoding]::UTF8)
            Remove-Item $outputFile -ErrorAction SilentlyContinue
            return $text
        } catch { }
    }
    return $null
}

# ============================================================================
# SIMPLE APPROACH: Wait for console to stabilize, extract new lines
# Compare RAW text for stabilization. Minimal cleaning only for Telegram.
# ============================================================================

function Get-ConsoleHash {
    param([string]$Text)
    if (-not $Text) { return "" }
    # Strip last 2 lines (status bar may update dynamically) + normalize
    $lines = $Text -split "`r?`n"
    if ($lines.Count -gt 2) { $lines = $lines[0..($lines.Count - 3)] }
    $normalized = ($lines | ForEach-Object { $_.TrimEnd() } | Where-Object { $_.Length -gt 0 }) -join "`n"
    # Remove control chars
    $normalized = $normalized -replace '[\u0000-\u001F\u007F]', ''
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalized)
    $md5 = [System.Security.Cryptography.MD5]::Create()
    return [System.BitConverter]::ToString($md5.ComputeHash($bytes))
}

function Wait-ClaudeResponse {
    param([int]$CmdPid, [string]$BaselineRaw, [int]$MaxWaitSec = 120)

    $pollInterval = 3
    $stableNeeded = 2

    # NEW STRATEGY: compare CONSECUTIVE reads (hash-based)
    # This works even when total console length stays constant (fixed-size window)
    $baseHash = Get-ConsoleHash $BaselineRaw

    Write-Host "[ClaudePlus] Attente reponse (baseHash=$($baseHash.Substring(0,8))...)..." -ForegroundColor DarkGray

    # Wait for Claude to start responding
    Start-Sleep -Seconds 5

    $previousHash = ""
    $stableCount = 0
    $stableRaw = ""
    $elapsed = 5

    while ($elapsed -lt $MaxWaitSec) {
        $currentRaw = Read-ClaudeConsole -CmdPid $CmdPid
        if (-not $currentRaw) {
            Write-Host "[ClaudePlus] Poll ${elapsed}s: lecture echouee" -ForegroundColor Red
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval
            continue
        }

        $curHash = Get-ConsoleHash $currentRaw
        $changedFromBase = ($curHash -ne $baseHash)
        $sameAsPrevious = ($curHash -eq $previousHash -and $previousHash -ne "")

        Write-Host "[ClaudePlus] Poll ${elapsed}s: changed=$changedFromBase stable=$stableCount hash=$($curHash.Substring(0,8))" -ForegroundColor DarkGray

        if ($sameAsPrevious) {
            $stableCount++
            Write-Host "[ClaudePlus] Stable $stableCount/$stableNeeded" -ForegroundColor DarkYellow
            if ($stableCount -ge $stableNeeded) {
                Write-Host "[ClaudePlus] Reponse stabilisee!" -ForegroundColor Green

                # Debug dump
                try {
                    [System.IO.File]::WriteAllText("$env:TEMP\claudeplus_baseline.txt", $BaselineRaw, [System.Text.Encoding]::UTF8)
                    [System.IO.File]::WriteAllText("$env:TEMP\claudeplus_response.txt", $currentRaw, [System.Text.Encoding]::UTF8)
                } catch { }

                $newContent = $null
                try {
                    $newContent = Extract-NewLines -BaselineRaw $BaselineRaw -CurrentRaw $currentRaw
                } catch {
                    Write-Host "[ClaudePlus] Erreur extraction: $_" -ForegroundColor Red
                }

                if ($newContent) {
                    Write-Host "[ClaudePlus] Extrait: $($newContent.Length) chars" -ForegroundColor Magenta
                } else {
                    Write-Host "[ClaudePlus] Extraction vide, fallback dump" -ForegroundColor Yellow
                    $curLines = ($currentRaw -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 2 })
                    $newContent = $curLines -join "`n"
                }

                $script:LastConsoleText = $currentRaw
                return $newContent
            }
        } else {
            $stableCount = 0
        }

        $previousHash = $curHash
        $stableRaw = $currentRaw
        Start-Sleep -Seconds $pollInterval
        $elapsed += $pollInterval
    }

    Write-Host "[ClaudePlus] Timeout ($MaxWaitSec s)" -ForegroundColor Yellow
    return $null
}

# Extract Claude's actual response from console output
# Compares with baseline, strips TUI chrome, keeps only meaningful content
function Extract-NewLines {
    param([string]$BaselineRaw, [string]$CurrentRaw)

    $baseLines = $BaselineRaw -split "`r?`n" | ForEach-Object { $_.TrimEnd() }
    $currLines = $CurrentRaw -split "`r?`n" | ForEach-Object { $_.TrimEnd() }

    # Build set of baseline lines
    $baseSet = @{}
    foreach ($line in $baseLines) {
        $t = $line.Trim()
        if ($t.Length -gt 0) { $baseSet[$t] = $true }
    }

    # Find new lines and clean them
    $newLines = @()
    foreach ($line in $currLines) {
        $t = $line.Trim()
        if ($t.Length -eq 0) { continue }
        if ($baseSet.ContainsKey($t)) { continue }

        # Skip TUI chrome
        if (Test-IsTuiNoise $t) { continue }

        # Clean the line: remove box-drawing, block chars, stray Unicode artifacts
        $clean = Remove-TuiChars $t
        $clean = $clean.Trim()
        if ($clean.Length -lt 2) { continue }

        # Skip user input echo lines ("> some text")
        if ($clean -match '^>\s+\S') { continue }
        # Skip empty prompt lines (just ">")
        if ($clean -match '^>\s*$') { continue }

        # Remove Claude bullet prefix
        $clean = $clean -replace '^\u25CF\s*', ''
        $clean = $clean.Trim()
        if ($clean.Length -lt 2) { continue }

        $newLines += $clean
    }

    if ($newLines.Count -eq 0) { return $null }

    $result = $newLines -join "`n"
    if ($result.Length -gt 3800) {
        $result = $result.Substring(0, 3800) + "`n[... tronque]"
    }

    return $result.Trim()
}

# Check if a line is TUI noise (should be skipped entirely)
function Test-IsTuiNoise {
    param([string]$Line)

    # Box-drawing chars anywhere (Unicode block: U+2500-U+257F)
    if ($Line -match '[\u2500-\u257F]') {
        # But only skip if the line is MOSTLY box-drawing (>50% special chars)
        $cleaned = $Line -replace '[\u2500-\u257F\u2580-\u259F\s]', ''
        if ($cleaned.Length -lt ($Line.Trim().Length / 3)) { return $true }
    }

    # Block elements (logo) — U+2580-U+259F
    if ($Line -match '[\u2580-\u259F]') {
        $cleaned = $Line -replace '[\u2580-\u259F\u2500-\u257F\s]', ''
        if ($cleaned.Length -lt ($Line.Trim().Length / 3)) { return $true }
    }

    # Status bar, model info, UI elements
    if ($Line -match 'Tips for getting started') { return $true }
    if ($Line -match 'Welcome back') { return $true }
    if ($Line -match 'Run /init') { return $true }
    if ($Line -match 'Recent activity|No recent activity') { return $true }
    if ($Line -match 'Opus.*context|Sonnet.*context|Haiku.*context') { return $true }
    if ($Line -match 'Claude Max|Claude Pro') { return $true }
    if ($Line -match '@\S+\.\S+.*Organization') { return $true }
    if ($Line -match '~\\Documents\\|~\\Desktop\\') { return $true }
    if ($Line -match '/gsd') { return $true }
    if ($Line -match 'bypass permissions') { return $true }
    if ($Line -match 'shift.tab to cycle') { return $true }
    if ($Line -match '^\s*\d+\s*%\s*$') { return $true }

    # Lines that are just prompt chars or separators
    if ($Line -match '^\s*[>]\s*$') { return $true }
    if ($Line -match '^\s*$') { return $true }

    return $false
}

# Remove box-drawing and TUI characters from a line, keep text content
function Remove-TuiChars {
    param([string]$Text)

    # Remove box-drawing chars (U+2500-U+257F)
    $t = $Text -replace '[\u2500-\u257F]', ''
    # Remove block elements (U+2580-U+259F)
    $t = $t -replace '[\u2580-\u259F]', ''
    # Remove arrows (U+2190-U+21FF)
    $t = $t -replace '[\u2190-\u21FF]', ''
    # Remove geometric shapes (U+25A0-U+25FF) but keep bullet
    $t = $t -replace '[\u25A0-\u25CE\u25D0-\u25FF]', ''
    # Remove misc symbols
    $t = $t -replace '[\u2700-\u27BF]', ''
    # Remove play buttons etc
    $t = $t -replace '[\u23E9-\u23FF]', ''
    # Remove CJK/Hangul artifacts (TUI rendering garbage)
    $t = $t -replace '[\uAC00-\uD7AF]', ''
    $t = $t -replace '[\u4E00-\u9FFF]', ''
    $t = $t -replace '[\u3000-\u303F]', ''
    # Remove other common TUI artifacts
    $t = $t -replace '[\uE000-\uF8FF]', ''
    $t = $t -replace '[\u00C0-\u00C6\u00C8-\u00CF\u00D0-\u00D6\u00D8-\u00DF](?=\s*$)', ''
    # Remove control chars
    $t = $t -replace '[\u0000-\u001F\u007F]', ''
    # Remove Latin Extended artifacts (Ǥ, Ø, etc. — not used in French/Dutch)
    $t = $t -replace '[\u0100-\u024F]', ''
    $t = $t -replace '[\u00D8-\u00DF]', ''
    # Remove stray single non-text chars at end of line
    $t = $t -replace '\s*[^\x20-\x7E\u00E0-\u00F6\u00F8-\u00FF]+\s*$', ''
    # Clean up multiple spaces
    $t = $t -replace '\s{2,}', ' '

    return $t.Trim()
}

# ============================================================================
# CONFIG MANAGEMENT
# ============================================================================

function Get-ClaudePlusConfig {
    $config = $null
    if (Test-Path $script:ConfigPath) {
        $config = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
    }
    if (-not $config) {
        $config = @{
            TelegramBotToken = ""
            TelegramChatId = ""
            DangerouslySkipPermissions = $true
            AutoTelegram = $true
        }
    }
    if ([string]::IsNullOrEmpty($config.TelegramBotToken)) {
        $imported = $false
        if (Test-Path $script:ProjectTelegramConfig) {
            try {
                $tj = Get-Content $script:ProjectTelegramConfig -Raw | ConvertFrom-Json
                if ($tj.TelegramBotToken) {
                    $config.TelegramBotToken = $tj.TelegramBotToken
                    $config.TelegramChatId = $tj.TelegramChatId
                    $imported = $true
                }
            } catch { }
        }
        if (-not $imported -and (Test-Path $script:VsExtensionConfig)) {
            try {
                $vj = Get-Content $script:VsExtensionConfig -Raw | ConvertFrom-Json
                if ($vj.TelegramBotToken) {
                    $config.TelegramBotToken = $vj.TelegramBotToken
                    $config.TelegramChatId = $vj.TelegramChatId
                    $imported = $true
                }
            } catch { }
        }
        if ($imported) { Save-ClaudePlusConfig $config }
    }
    return $config
}

function Save-ClaudePlusConfig($config) {
    $dir = Split-Path $script:ConfigPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $config | ConvertTo-Json | Set-Content $script:ConfigPath -Encoding UTF8
}

# ============================================================================
# TELEGRAM
# ============================================================================

function Send-TelegramMessage {
    param([string]$Message, [string]$Token, [string]$ChatId)
    if ([string]::IsNullOrEmpty($Message) -or [string]::IsNullOrEmpty($Token) -or [string]::IsNullOrEmpty($ChatId)) { return }
    if ($Message.Length -gt 4000) { $Message = $Message.Substring(0, 4000) + "`n[...]" }
    try {
        $body = @{ chat_id = $ChatId; text = $Message; disable_web_page_preview = "true" }
        Invoke-RestMethod -Uri "https://api.telegram.org/bot$Token/sendMessage" -Method Post -Body $body -TimeoutSec 10 -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}

function Delete-TelegramWebhook {
    param([string]$Token)
    try { Invoke-RestMethod -Uri "https://api.telegram.org/bot$Token/deleteWebhook" -Method Get -TimeoutSec 10 -ErrorAction SilentlyContinue | Out-Null } catch { }
}

# ============================================================================
# FIND CHILD CMD.EXE PID
# ============================================================================

function Find-ChildCmdPid {
    param([int]$ParentPid)
    $children = Get-CimInstance Win32_Process -Filter "ParentProcessId=$ParentPid" -ErrorAction SilentlyContinue
    if ($children) {
        foreach ($child in $children) {
            if ($child.Name -eq "cmd.exe") { return [int]$child.ProcessId }
        }
        foreach ($child in $children) {
            $gcs = Get-CimInstance Win32_Process -Filter "ParentProcessId=$($child.ProcessId)" -ErrorAction SilentlyContinue
            if ($gcs) {
                foreach ($gc in $gcs) {
                    if ($gc.Name -eq "cmd.exe") { return [int]$gc.ProcessId }
                }
            }
        }
    }
    return 0
}

# ============================================================================
# GET WINDOW HANDLE FOR A PID (tries the process and its parent)
# ============================================================================

function Get-WindowHandleForPid {
    param([int]$Pid1)

    # Try the process itself
    $proc = Get-Process -Id $Pid1 -ErrorAction SilentlyContinue
    if ($proc -and $proc.MainWindowHandle -ne [IntPtr]::Zero) {
        return $proc.MainWindowHandle
    }

    # Try parent (conhost owns the window)
    try {
        $wmi = Get-CimInstance Win32_Process -Filter "ProcessId=$Pid1" -ErrorAction SilentlyContinue
        if ($wmi -and $wmi.ParentProcessId) {
            $parent = Get-Process -Id $wmi.ParentProcessId -ErrorAction SilentlyContinue
            if ($parent -and $parent.MainWindowHandle -ne [IntPtr]::Zero) {
                return $parent.MainWindowHandle
            }
        }
    } catch { }

    return [IntPtr]::Zero
}

# ============================================================================
# SEND TEXT TO CLAUDE WINDOW
# Uses: MainWindowHandle + SetForegroundWindow + WshShell.SendKeys
# ============================================================================

function Send-TextToClaude {
    param([string]$Text)

    # Find cmd PID if not set
    if ($script:ClaudeCmdPid -eq 0 -and $script:ClaudeProcess -and -not $script:ClaudeProcess.HasExited) {
        $script:ClaudeCmdPid = Find-ChildCmdPid -ParentPid $script:ClaudeProcess.Id
        if ($script:ClaudeCmdPid -gt 0) {
            Write-Host "[ClaudePlus] cmd.exe PID: $($script:ClaudeCmdPid)" -ForegroundColor DarkGray
        }
    }

    if ($script:ClaudeCmdPid -le 0) {
        Write-Host "[ClaudePlus] ERREUR: cmd.exe PID introuvable" -ForegroundColor Red
        return $false
    }

    # WriteConsoleInput via helper process (separate process to avoid FreeConsole on our console)
    # C# code in separate .cs file to avoid nested here-string issues
    Write-Host "[ClaudePlus] Envoi via WriteConsoleInput (helper) vers PID $($script:ClaudeCmdPid)..." -ForegroundColor DarkGray

    try {
        # 1. Write C# source file
        $csFile = "$env:TEMP\claudeplus_conwriter.cs"
        $csCode = "using System;" + [Environment]::NewLine
        $csCode += "using System.Runtime.InteropServices;" + [Environment]::NewLine
        $csCode += "using System.Threading;" + [Environment]::NewLine
        $csCode += "public class ConW {" + [Environment]::NewLine
        $csCode += "    [DllImport(`"kernel32.dll`", SetLastError=true)] public static extern bool FreeConsole();" + [Environment]::NewLine
        $csCode += "    [DllImport(`"kernel32.dll`", SetLastError=true)] public static extern bool AttachConsole(int pid);" + [Environment]::NewLine
        $csCode += "    [DllImport(`"kernel32.dll`", SetLastError=true)] public static extern IntPtr GetStdHandle(int h);" + [Environment]::NewLine
        $csCode += "    [DllImport(`"kernel32.dll`", CharSet=CharSet.Unicode, SetLastError=true)] public static extern bool WriteConsoleInput(IntPtr hIn, INPUT_RECORD[] buf, uint len, out uint written);" + [Environment]::NewLine
        $csCode += "    [DllImport(`"kernel32.dll`")] public static extern int GetLastError();" + [Environment]::NewLine
        $csCode += "    public const int STD_INPUT_HANDLE = -10;" + [Environment]::NewLine
        $csCode += "    public const ushort KEY_EVENT = 0x0001;" + [Environment]::NewLine
        $csCode += "    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]" + [Environment]::NewLine
        $csCode += "    public struct KEY_EVENT_RECORD { public bool bKeyDown; public ushort wRepeatCount; public ushort wVirtualKeyCode; public ushort wVirtualScanCode; public char UnicodeChar; public uint dwControlKeyState; }" + [Environment]::NewLine
        $csCode += "    [StructLayout(LayoutKind.Explicit, CharSet=CharSet.Unicode)]" + [Environment]::NewLine
        $csCode += "    public struct INPUT_RECORD { [FieldOffset(0)] public ushort EventType; [FieldOffset(4)] public KEY_EVENT_RECORD KeyEvent; }" + [Environment]::NewLine
        $csCode += "    public static string Write(int pid, string text) {" + [Environment]::NewLine
        $csCode += "        FreeConsole();" + [Environment]::NewLine
        $csCode += "        if (!AttachConsole(pid)) return `"FAIL:Attach err=`" + GetLastError();" + [Environment]::NewLine
        $csCode += "        try {" + [Environment]::NewLine
        $csCode += "            IntPtr h = GetStdHandle(STD_INPUT_HANDLE);" + [Environment]::NewLine
        $csCode += "            if (h == IntPtr.Zero || h == new IntPtr(-1)) return `"FAIL:Handle`";" + [Environment]::NewLine
        $csCode += "            int n = 0;" + [Environment]::NewLine
        $csCode += "            foreach (char c in text) {" + [Environment]::NewLine
        $csCode += "                INPUT_RECORD[] r = new INPUT_RECORD[2];" + [Environment]::NewLine
        $csCode += "                r[0].EventType = KEY_EVENT; r[0].KeyEvent.bKeyDown = true; r[0].KeyEvent.wRepeatCount = 1; r[0].KeyEvent.UnicodeChar = c;" + [Environment]::NewLine
        $csCode += "                r[1] = r[0]; r[1].KeyEvent.bKeyDown = false;" + [Environment]::NewLine
        $csCode += "                uint w; if (!WriteConsoleInput(h, r, 2, out w)) return `"FAIL:Write n=`" + n;" + [Environment]::NewLine
        $csCode += "                n++; Thread.Sleep(15);" + [Environment]::NewLine
        $csCode += "            }" + [Environment]::NewLine
        $csCode += "            Thread.Sleep(100);" + [Environment]::NewLine
        $csCode += "            INPUT_RECORD[] e = new INPUT_RECORD[2];" + [Environment]::NewLine
        $csCode += "            e[0].EventType = KEY_EVENT; e[0].KeyEvent.bKeyDown = true; e[0].KeyEvent.wRepeatCount = 1; e[0].KeyEvent.wVirtualKeyCode = 0x0D; e[0].KeyEvent.UnicodeChar = (char)13;" + [Environment]::NewLine
        $csCode += "            e[1] = e[0]; e[1].KeyEvent.bKeyDown = false;" + [Environment]::NewLine
        $csCode += "            uint ew; WriteConsoleInput(h, e, 2, out ew);" + [Environment]::NewLine
        $csCode += "            return `"OK:sent=`" + n;" + [Environment]::NewLine
        $csCode += "        } finally { FreeConsole(); }" + [Environment]::NewLine
        $csCode += "    }" + [Environment]::NewLine
        $csCode += "}" + [Environment]::NewLine
        [System.IO.File]::WriteAllText($csFile, $csCode, [System.Text.Encoding]::ASCII)

        # 2. Write helper PS1 that uses the .cs file
        $helperScript = "$env:TEMP\claudeplus_writer.ps1"
        $helperLines = @()
        $helperLines += 'param([int]$TargetPid, [string]$TextFile, [string]$CsFile)'
        $helperLines += 'try {'
        $helperLines += '    $text = [System.IO.File]::ReadAllText($TextFile, [System.Text.Encoding]::UTF8)'
        $helperLines += '    Add-Type -Path $CsFile'
        $helperLines += '    $result = [ConW]::Write($TargetPid, $text)'
        $helperLines += '    [System.IO.File]::WriteAllText("$TextFile.log", "Result=$result")'
        $helperLines += '    if ($result.StartsWith("OK")) { [System.IO.File]::WriteAllText("$TextFile.ok", $result) }'
        $helperLines += '} catch {'
        $helperLines += '    [System.IO.File]::WriteAllText("$TextFile.log", "EXCEPTION: $_")'
        $helperLines += '}'
        $helperContent = $helperLines -join "`r`n"
        [System.IO.File]::WriteAllText($helperScript, $helperContent, [System.Text.Encoding]::ASCII)

        # 3. Write text to temp file and run helper
        $textFile = "$env:TEMP\claudeplus_sendtext.txt"
        $okFile = "$textFile.ok"
        $logFile = "$textFile.log"
        if (Test-Path $okFile) { Remove-Item $okFile -ErrorAction SilentlyContinue }
        if (Test-Path $logFile) { Remove-Item $logFile -ErrorAction SilentlyContinue }
        [System.IO.File]::WriteAllText($textFile, $Text, [System.Text.Encoding]::UTF8)

        $proc = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$helperScript`" -TargetPid $($script:ClaudeCmdPid) -TextFile `"$textFile`" -CsFile `"$csFile`"" `
            -PassThru -WindowStyle Hidden
        $proc.WaitForExit(15000) | Out-Null
        if (-not $proc.HasExited) { $proc.Kill() }

        # Read log for debug
        if (Test-Path $logFile) {
            $logContent = [System.IO.File]::ReadAllText($logFile)
            Write-Host "[ClaudePlus] Helper log: $logContent" -ForegroundColor DarkGray
        } else {
            Write-Host "[ClaudePlus] Pas de log helper (process crash?)" -ForegroundColor Red
        }

        if (Test-Path $okFile) {
            $okContent = [System.IO.File]::ReadAllText($okFile)
            Write-Host "[ClaudePlus] Texte envoye! ($okContent)" -ForegroundColor Green
            Remove-Item $okFile -ErrorAction SilentlyContinue
            return $true
        }
        Write-Host "[ClaudePlus] Helper process echoue" -ForegroundColor Red
    } catch {
        Write-Host "[ClaudePlus] Erreur helper: $_" -ForegroundColor Red
    }

    return $false
}

# ============================================================================
# MAIN COMMAND: claudeplus
# ============================================================================

function Invoke-ClaudePlus {
    [CmdletBinding()]
    param(
        [switch]$NoTelegram,
        [switch]$NoDangerouslySkipPermissions,
        [Parameter(ValueFromRemainingArguments=$true)]
        [string[]]$ClaudeArgs
    )

    $config = Get-ClaudePlusConfig

    $claudePath = "$env:USERPROFILE\.local\bin\claude.exe"
    if (-not (Test-Path $claudePath)) {
        $claudePath = (Get-Command claude -ErrorAction SilentlyContinue).Source
    }
    if (-not $claudePath -or -not (Test-Path $claudePath)) {
        Write-Host "[ClaudePlus] ERREUR: claude.exe introuvable." -ForegroundColor Red
        return
    }

    $claudeArgsList = @()
    if ($config.DangerouslySkipPermissions -and -not $NoDangerouslySkipPermissions) {
        $claudeArgsList += "--dangerously-skip-permissions"
    }
    if ($ClaudeArgs) { $claudeArgsList += $ClaudeArgs }

    Write-Host ""
    Write-Host "  +============================================+" -ForegroundColor DarkCyan
    Write-Host "  |         ClaudePlus - FiscalIQ               |" -ForegroundColor DarkCyan
    Write-Host "  |   Claude Code + Telegram Mirror             |" -ForegroundColor DarkCyan
    Write-Host "  +============================================+" -ForegroundColor DarkCyan
    Write-Host ""

    $useTelegram = (-not $NoTelegram -and $config.AutoTelegram -and $config.TelegramBotToken -and $config.TelegramChatId)

    if ($useTelegram) {
        Write-Host "[ClaudePlus] Mode Mirror Telegram actif." -ForegroundColor Green

        Delete-TelegramWebhook -Token $config.TelegramBotToken

        $workDir = (Get-Location).Path
        $allArgs = $claudeArgsList -join " "
        $sessionId = Get-Random -Minimum 10000 -Maximum 99999
        $batPath = "$env:TEMP\claudeplus_$sessionId.bat"
        $batLines = @(
            "@echo off",
            "cd /d `"$workDir`"",
            "cls",
            "`"$claudePath`" $allArgs",
            "pause"
        )
        $batLines -join "`r`n" | Set-Content $batPath -Encoding ASCII

        # Launch via conhost.exe
        $conhost = "$env:SystemRoot\System32\conhost.exe"
        if (Test-Path $conhost) {
            Write-Host "[ClaudePlus] Lancement via conhost.exe..." -ForegroundColor Cyan
            $script:ClaudeProcess = Start-Process -FilePath $conhost -ArgumentList "cmd.exe /c `"$batPath`"" -PassThru
        } else {
            $script:ClaudeProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batPath`"" -PassThru
        }

        Write-Host "[ClaudePlus] Conhost PID: $($script:ClaudeProcess.Id)" -ForegroundColor DarkGray

        # Wait and find child cmd.exe + its window handle
        Write-Host "[ClaudePlus] Recherche de la fenetre Claude..." -ForegroundColor DarkGray
        $found = $false
        for ($i = 0; $i -lt 20; $i++) {
            Start-Sleep -Milliseconds 500

            # Find child cmd PID
            if ($script:ClaudeCmdPid -eq 0) {
                $script:ClaudeCmdPid = Find-ChildCmdPid -ParentPid $script:ClaudeProcess.Id
            }

            # Try to get window handle
            if ($script:ClaudeCmdPid -gt 0) {
                $hwnd = Get-WindowHandleForPid -Pid1 $script:ClaudeCmdPid
                if ($hwnd -ne [IntPtr]::Zero) {
                    $found = $true
                    $script:ClaudeWindowHandle = $hwnd
                    Write-Host "[ClaudePlus] FENETRE TROUVEE! cmd PID=$($script:ClaudeCmdPid), Handle=$hwnd" -ForegroundColor Green
                    break
                }
            }
        }

        if (-not $found) {
            Write-Host "[ClaudePlus] Fenetre non trouvee apres 10s, tentative scan..." -ForegroundColor Yellow
            # Scan all cmd.exe for our child
            Get-Process -Name "cmd" -ErrorAction SilentlyContinue | ForEach-Object {
                Write-Host "  cmd PID=$($_.Id) Handle=$($_.MainWindowHandle) Title='$($_.MainWindowTitle)'" -ForegroundColor DarkGray
                if ($_.MainWindowHandle -ne [IntPtr]::Zero -and -not $found) {
                    try {
                        $wmi = Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue
                        if ($wmi -and [int]$wmi.ParentProcessId -eq $script:ClaudeProcess.Id) {
                            $script:ClaudeCmdPid = $_.Id
                            $found = $true
                            Write-Host "[ClaudePlus] FENETRE TROUVEE via scan! PID=$($_.Id)" -ForegroundColor Green
                        }
                    } catch { }
                }
            }
        }

        if (-not $found) {
            Write-Host "[ClaudePlus] ATTENTION: Pas de handle fenetre. Le mirror peut echouer." -ForegroundColor Red
        }

        Send-TelegramMessage -Message "ClaudePlus demarre! Envoyez vos messages.`n/stop pour arreter." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

        # Initialize console buffer reader
        Initialize-ConsoleReader
        $script:LastConsoleText = ""

        # Do initial baseline read
        if ($script:ClaudeCmdPid -gt 0) {
            $initialText = Read-ClaudeConsole -CmdPid $script:ClaudeCmdPid
            if ($initialText) {
                $script:LastConsoleText = $initialText
                Write-Host "[ClaudePlus] Console baseline OK ($($initialText.Length) chars)" -ForegroundColor Green
            } else {
                Write-Host "[ClaudePlus] WARN: Lecture console echouee (retry plus tard)" -ForegroundColor Yellow
            }
        }

        # Telegram polling loop (response captured after each send)
        Write-Host "[ClaudePlus] Polling Telegram... (Ctrl+C pour arreter)" -ForegroundColor Green
        Write-Host ""

        # Flush old Telegram messages to avoid replaying stale commands
        $lastUpdateId = 0
        try {
            $flushUrl = "https://api.telegram.org/bot$($config.TelegramBotToken)/getUpdates?limit=100&timeout=1"
            $flushResp = Invoke-RestMethod -Uri $flushUrl -Method Get -TimeoutSec 5 -ErrorAction SilentlyContinue
            if ($flushResp.ok -and $flushResp.result -and $flushResp.result.Count -gt 0) {
                $lastUpdateId = ($flushResp.result | Select-Object -Last 1).update_id
                Write-Host "[ClaudePlus] Flush: $($flushResp.result.Count) anciens messages ignores" -ForegroundColor DarkGray
            }
        } catch { }

        $stopRequested = $false
        $waitingForResponse = $false

        try {
            while (-not $stopRequested) {
                # Check if conhost or cmd.exe has exited
                if ($script:ClaudeProcess.HasExited) { break }
                if ($script:ClaudeCmdPid -gt 0) {
                    $cmdProc = Get-Process -Id $script:ClaudeCmdPid -ErrorAction SilentlyContinue
                    if (-not $cmdProc) { Write-Host "[ClaudePlus] cmd.exe termine, arret." -ForegroundColor Yellow; break }
                }

                # --- CHECK TELEGRAM ---
                try {
                    $url = "https://api.telegram.org/bot$($config.TelegramBotToken)/getUpdates?limit=10&timeout=2"
                    if ($lastUpdateId -gt 0) { $url += "&offset=$($lastUpdateId + 1)" }
                    $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 5 -ErrorAction Stop

                    if ($response.ok -and $response.result) {
                        foreach ($update in $response.result) {
                            if ($update.update_id -gt $lastUpdateId) { $lastUpdateId = $update.update_id }
                            $msg = $update.message
                            if (-not $msg -or [string]::IsNullOrEmpty($msg.text)) { continue }
                            if ($msg.chat.id -ne [long]$config.TelegramChatId) { continue }
                            $text = $msg.text.Trim()

                            if ($text -match "^/stop") {
                                Send-TelegramMessage -Message "Mirror arrete." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                $stopRequested = $true
                                break
                            }

                            # Skip if already waiting for a response
                            if ($waitingForResponse) {
                                Send-TelegramMessage -Message "[ATTENTE] Claude est en train de repondre, patientez..." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }

                            Write-Host "[Telegram -> Claude] $text" -ForegroundColor Cyan

                            # Capture baseline BEFORE sending
                            $preBaseline = Read-ClaudeConsole -CmdPid $script:ClaudeCmdPid
                            if (-not $preBaseline) { $preBaseline = $script:LastConsoleText }

                            $sent = Send-TextToClaude -Text $text

                            if ($sent) {
                                Write-Host "[ClaudePlus] Message envoye, attente reponse..." -ForegroundColor DarkGray
                                $waitingForResponse = $true

                                # Wait for Claude to finish responding (blocking)
                                # Pass the PRE-SEND baseline so we detect any change
                                $responseText = Wait-ClaudeResponse -CmdPid $script:ClaudeCmdPid -BaselineRaw $preBaseline -MaxWaitSec 120

                                $waitingForResponse = $false

                                if ($responseText -and $responseText.Length -gt 3) {
                                    # Truncate if too long for Telegram
                                    if ($responseText.Length -gt 3500) {
                                        $responseText = $responseText.Substring(0, 3500) + "`n[... tronque]"
                                    }
                                    Write-Host "[Claude -> Telegram] $($responseText.Length) chars" -ForegroundColor Magenta
                                    Send-TelegramMessage -Message $responseText -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                } else {
                                    Write-Host "[ClaudePlus] Pas de reponse detectee (timeout ou vide)" -ForegroundColor Yellow
                                    Send-TelegramMessage -Message "[Pas de reponse detectee - verifiez le terminal]" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                }
                            } else {
                                Send-TelegramMessage -Message "[ERREUR] Fenetre Claude introuvable" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            }
                        }
                    }
                }
                catch { }

                Start-Sleep -Milliseconds 500
            }
        }
        finally {
            Send-TelegramMessage -Message "ClaudePlus terminee." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
            if (Test-Path $batPath) { Remove-Item $batPath -ErrorAction SilentlyContinue }
            if ($script:ConsoleReaderScript -and (Test-Path $script:ConsoleReaderScript)) { Remove-Item $script:ConsoleReaderScript -ErrorAction SilentlyContinue }
            Write-Host ""
            Write-Host "[ClaudePlus] Session terminee." -ForegroundColor DarkCyan
        }
    }
    else {
        Write-Host "[ClaudePlus] Lancement de Claude Code..." -ForegroundColor Cyan
        Write-Host ""
        try { & $claudePath @claudeArgsList }
        finally {
            Write-Host ""
            Write-Host "[ClaudePlus] Session terminee." -ForegroundColor DarkCyan
        }
    }
}

# ============================================================================
# CONFIG COMMAND
# ============================================================================

function Set-ClaudePlusConfig {
    [CmdletBinding()]
    param(
        [string]$TelegramBotToken,
        [string]$TelegramChatId,
        [Nullable[bool]]$DangerouslySkipPermissions,
        [Nullable[bool]]$AutoTelegram
    )
    $config = Get-ClaudePlusConfig
    $changed = $false
    if ($TelegramBotToken) { $config.TelegramBotToken = $TelegramBotToken; $changed = $true }
    if ($TelegramChatId) { $config.TelegramChatId = $TelegramChatId; $changed = $true }
    if ($null -ne $DangerouslySkipPermissions) { $config.DangerouslySkipPermissions = $DangerouslySkipPermissions; $changed = $true }
    if ($null -ne $AutoTelegram) { $config.AutoTelegram = $AutoTelegram; $changed = $true }
    if ($changed) {
        Save-ClaudePlusConfig $config
        Write-Host "[ClaudePlus] Configuration sauvegardee." -ForegroundColor Green
    }
    Write-Host ""
    Write-Host "  Configuration ClaudePlus:" -ForegroundColor Cyan
    if ($config.TelegramBotToken) {
        $tokenDisplay = $config.TelegramBotToken.Substring(0, [Math]::Min(10, $config.TelegramBotToken.Length)) + "..."
    } else { $tokenDisplay = "(non configure)" }
    Write-Host "  TelegramBotToken : $tokenDisplay" -ForegroundColor White
    if ($config.TelegramChatId) { $chatIdDisplay = $config.TelegramChatId } else { $chatIdDisplay = "(non configure)" }
    Write-Host "  TelegramChatId   : $chatIdDisplay" -ForegroundColor White
    Write-Host "  SkipPermissions  : $($config.DangerouslySkipPermissions)" -ForegroundColor White
    Write-Host "  AutoTelegram     : $($config.AutoTelegram)" -ForegroundColor White
    Write-Host ""
}

# ============================================================================
# EXPORT
# ============================================================================

Set-Alias -Name claudeplus -Value Invoke-ClaudePlus -Scope Global
Set-Alias -Name claudeplus-config -Value Set-ClaudePlusConfig -Scope Global
Export-ModuleMember -Function Invoke-ClaudePlus, Set-ClaudePlusConfig, Get-ClaudePlusConfig -Alias claudeplus, claudeplus-config
