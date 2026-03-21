# ============================================================================
# ClaudePlus PowerShell Module
# Wraps Claude Code with Telegram bidirectional mirror
# Author: Majid - FiscalIQ
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms

$script:ConfigPath = "$env:LOCALAPPDATA\ClaudePlus\config.json"
$script:ProjectTelegramConfig = "$([Environment]::GetFolderPath('MyDocuments'))\Visual Studio 2026\FiscalIQ\Claude Code Extension\telegram-config.json"
$script:VsExtensionConfig = "$env:LOCALAPPDATA\ClaudeCodeExtension\claudecode-settings.json"
$script:SessionRegistryDir = "$env:LOCALAPPDATA\ClaudePlus\sessions"
$script:ClaudeProcess = $null
$script:ClaudeCmdPid = 0
$script:WshShell = New-Object -ComObject WScript.Shell
$script:SessionName = $null  # Set via -Name parameter, used for @prefix routing

# No direct P/Invoke for console write - PowerShell shares its own console
# All WriteConsoleInput calls go through a helper process (see Send-TextToClaude)

$script:LastConsoleText = ""
$script:ReaderExe = $null

# ============================================================================
# READ CLAUDE CONSOLE OUTPUT
# Strategy: Compile a standalone .exe (once) that does:
#   FreeConsole + AttachConsole(pid) + CreateFile("CONOUT$") + ReadConsoleOutputCharacter
# The .exe runs as a separate process with no PS overhead.
# ============================================================================

function Initialize-ConsoleReader {
    $script:ReaderExe = "$env:TEMP\claudeplus_reader.exe"

    # Reuse compiled exe if it exists (fast startup)
    if (Test-Path $script:ReaderExe) {
        Write-Host "[ClaudePlus] Lecteur console OK (exe)" -ForegroundColor DarkGray
        return
    }

    # Find csc.exe (.NET Framework compiler)
    $cscPath = $null
    $fwDirs = @(
        "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319",
        "$env:SystemRoot\Microsoft.NET\Framework\v4.0.30319"
    )
    foreach ($d in $fwDirs) {
        $p = Join-Path $d "csc.exe"
        if (Test-Path $p) { $cscPath = $p; break }
    }
    if (-not $cscPath) {
        Write-Host "[ClaudePlus] WARN: csc.exe introuvable, lecteur desactive" -ForegroundColor Yellow
        $script:ReaderExe = $null
        return
    }

    # Write C# source
    $csFile = "$env:TEMP\claudeplus_reader.cs"
    $csSource = @'
using System;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
class Program {
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool FreeConsole();
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool AttachConsole(int pid);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool GetConsoleScreenBufferInfo(IntPtr h, out CSBI i);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern bool ReadConsoleOutputCharacter(IntPtr h, StringBuilder b, int len, COORD c, out int nr);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern IntPtr CreateFile(string name, uint access, uint share, IntPtr sec, uint disp, uint flags, IntPtr tmpl);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    [DllImport("kernel32.dll")] static extern uint GetLastError();
    [StructLayout(LayoutKind.Sequential)] public struct COORD { public short X, Y; }
    [StructLayout(LayoutKind.Sequential)] public struct SRECT { public short L, T, R, B; }
    [StructLayout(LayoutKind.Sequential)] public struct CSBI { public COORD sz, cur; public short attr; public SRECT win; public COORD maxsz; }
    static int Main(string[] args) {
        if (args.Length < 2) { Console.Error.WriteLine("Usage: reader.exe outputFile pid1 [pid2 ...]"); return 1; }
        string outFile = args[0];
        FreeConsole();
        for (int i = 1; i < args.Length; i++) {
            int pid;
            if (!int.TryParse(args[i], out pid) || pid <= 0) continue;
            Console.Error.Write("PID " + pid + ": ");
            bool attached = AttachConsole(pid);
            if (!attached) { Console.Error.WriteLine("ATTACH_FAIL:" + GetLastError()); continue; }
            IntPtr h = CreateFile("CONOUT$", 0x80000000|0x40000000, 1|2, IntPtr.Zero, 3, 0, IntPtr.Zero);
            if (h == IntPtr.Zero || h == new IntPtr(-1)) { Console.Error.WriteLine("CONOUT_FAIL:" + GetLastError()); FreeConsole(); continue; }
            CSBI info;
            if (!GetConsoleScreenBufferInfo(h, out info)) { Console.Error.WriteLine("CSBI_FAIL:" + GetLastError()); CloseHandle(h); FreeConsole(); continue; }
            int w = info.sz.X;
            var sb = new StringBuilder();
            // Read visible window only (TUI uses alternate screen buffer)
            for (int y = info.win.T; y <= info.win.B; y++) {
                var line = new StringBuilder(w + 2);
                COORD c; c.X = 0; c.Y = (short)y;
                int nr;
                ReadConsoleOutputCharacter(h, line, w, c, out nr);
                sb.AppendLine(line.ToString().TrimEnd());
            }
            CloseHandle(h); FreeConsole();
            string result = sb.ToString();
            if (result.Trim().Length > 20) {
                File.WriteAllText(outFile, result, Encoding.UTF8);
                Console.Error.WriteLine("SUCCESS (" + result.Length + " chars)");
                return 0;
            }
            Console.Error.WriteLine("TOO_SHORT:" + result.Trim().Length);
        }
        Console.Error.WriteLine("All PIDs failed");
        return 1;
    }
}
'@
    [System.IO.File]::WriteAllText($csFile, $csSource, [System.Text.Encoding]::UTF8)

    # Compile
    Write-Host "[ClaudePlus] Compilation lecteur console..." -ForegroundColor DarkGray
    $compileResult = & $cscPath /nologo /optimize /out:$script:ReaderExe /target:exe $csFile 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[ClaudePlus] WARN: Echec compilation: $compileResult" -ForegroundColor Yellow
        $script:ReaderExe = $null
    } else {
        Write-Host "[ClaudePlus] Lecteur compile OK" -ForegroundColor Green
    }
    Remove-Item $csFile -ErrorAction SilentlyContinue
}

function Get-ClaudeChildPids {
    param([int]$CmdPid)
    $pids = @()
    if ($CmdPid -le 0) { return $pids }
    try {
        $children = Get-CimInstance Win32_Process -Filter "ParentProcessId=$CmdPid" -ErrorAction SilentlyContinue
        foreach ($child in $children) {
            $pids += [int]$child.ProcessId
            $gcs = Get-CimInstance Win32_Process -Filter "ParentProcessId=$($child.ProcessId)" -ErrorAction SilentlyContinue
            foreach ($gc in $gcs) { $pids += [int]$gc.ProcessId }
        }
    } catch {}
    return $pids
}

function Read-ClaudeConsole {
    param([int]$CmdPid)

    if ($CmdPid -le 0) { return $null }
    if (-not $script:ReaderExe -or -not (Test-Path $script:ReaderExe)) { return $null }

    $outputFile = "$env:TEMP\claudeplus_console_out.txt"
    Remove-Item $outputFile -ErrorAction SilentlyContinue

    # Build PID list: grandchildren first (claude.exe/node.exe), then cmd.exe
    $allPids = @()
    $childPids = Get-ClaudeChildPids -CmdPid $CmdPid
    $allPids += $childPids
    $allPids += $CmdPid
    $pidArgs = ($allPids | ForEach-Object { $_.ToString() }) -join " "

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $script:ReaderExe
        $psi.Arguments = "`"$outputFile`" $pidArgs"
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardError = $true
        $proc = [System.Diagnostics.Process]::Start($psi)
        $stderr = $proc.StandardError.ReadToEnd()
        $proc.WaitForExit(10000) | Out-Null
        if (-not $proc.HasExited) { $proc.Kill() }

        # Show diagnostics from stderr (always)
        if ($stderr) {
            foreach ($line in ($stderr -split "`r?`n" | Where-Object { $_.Trim().Length -gt 0 })) {
                $col = if ($line -match "SUCCESS") { "Green" } else { "Cyan" }
                Write-Host "[Reader] $line" -ForegroundColor $col
            }
        }
    } catch {
        Write-Host "[Reader] Erreur lancement: $_" -ForegroundColor Red
        return $null
    }

    if (Test-Path $outputFile) {
        try {
            $text = [System.IO.File]::ReadAllText($outputFile, [System.Text.Encoding]::UTF8)
            Remove-Item $outputFile -ErrorAction SilentlyContinue
            if ($text -and $text.Trim().Length -gt 5) { return $text }
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
    $stableNeeded = 3

    # STRATEGY v3: The TUI cursor blink corrupts a few chars on every read,
    # making exact text comparison impossible. Instead we count consecutive
    # polls where a response EXISTS (>10 chars) with SIMILAR length (+-20).
    # After 3 consecutive similar-length detections, return the longest one.

    Write-Host "[ClaudePlus] Attente reponse (v3)..." -ForegroundColor DarkGray

    Start-Sleep -Seconds 5

    $consecutiveHits = 0
    $previousLen = 0
    $bestExtracted = ""
    $elapsed = 5

    while ($elapsed -lt $MaxWaitSec) {
        $currentRaw = Read-ClaudeConsole -CmdPid $CmdPid
        if (-not $currentRaw) {
            Write-Host "[ClaudePlus] Poll ${elapsed}s: lecture echouee" -ForegroundColor Red
            $consecutiveHits = 0
            Start-Sleep -Seconds $pollInterval
            $elapsed += $pollInterval
            continue
        }

        $extracted = $null
        try {
            $extracted = Extract-NewLines -BaselineRaw $BaselineRaw -CurrentRaw $currentRaw
        } catch { }

        if ($extracted -and $extracted.Length -gt 10) {
            $curLen = $extracted.Length

            # Keep the longest extraction
            if ($curLen -gt $bestExtracted.Length) {
                $bestExtracted = $extracted
            }

            # Check if length is similar to previous (within +-20 chars)
            if ($previousLen -gt 0 -and [Math]::Abs($curLen - $previousLen) -le 20) {
                $consecutiveHits++
                Write-Host "[ClaudePlus] Poll ${elapsed}s: STABLE $consecutiveHits/$stableNeeded ($curLen chars, prev=$previousLen)" -ForegroundColor DarkYellow
                if ($consecutiveHits -ge $stableNeeded) {
                    Write-Host "[ClaudePlus] Reponse capturee! ($($bestExtracted.Length) chars)" -ForegroundColor Green
                    $script:LastConsoleText = $currentRaw
                    return $bestExtracted
                }
            } else {
                $consecutiveHits = 0
                Write-Host "[ClaudePlus] Poll ${elapsed}s: reponse en cours ($curLen chars, delta=$([Math]::Abs($curLen - $previousLen)))" -ForegroundColor DarkGray
            }
            $previousLen = $curLen
        } else {
            $consecutiveHits = 0
            $previousLen = 0
            Write-Host "[ClaudePlus] Poll ${elapsed}s: pas encore de reponse" -ForegroundColor DarkGray
        }

        Start-Sleep -Seconds $pollInterval
        $elapsed += $pollInterval
    }

    # Timeout but we have something -- return best effort
    if ($bestExtracted.Length -gt 10) {
        Write-Host "[ClaudePlus] Timeout mais reponse partielle ($($bestExtracted.Length) chars)" -ForegroundColor Yellow
        return $bestExtracted
    }

    Write-Host "[ClaudePlus] Timeout - aucune reponse" -ForegroundColor Yellow
    return $null
}

# Extract Claude's actual response from console output
# Strategy 1: locate the ● response marker and capture what follows
# Strategy 2: diff from baseline as fallback
function Extract-NewLines {
    param([string]$BaselineRaw, [string]$CurrentRaw)

    # Remove raw control chars before processing
    $cleanedCurrent = $CurrentRaw -replace '[\u0000-\u001F\u007F]', ' '

    # ---- STRATEGY 1: find ● (Claude's response marker) ----
    $lines = $cleanedCurrent -split "`r?`n"
    $responseLines = @()
    $capturing = $false
    $emptyLineCount = 0

    foreach ($line in $lines) {
        $t = $line.TrimEnd()

        # Does this line contain the ● marker?
        if ($t -match '[\u25CF]') {
            $capturing = $true
            $emptyLineCount = 0
            # Take everything AFTER the ● (and optional space)
            $after = ($t -replace '^.*[\u25CF]\s*', '').TrimEnd()
            $after = Remove-TuiChars $after
            $after = $after.Trim()
            if ($after.Length -gt 1) { $responseLines += $after }
            continue
        }

        if ($capturing) {
            $trimmed = $t.Trim()

            # Stop capturing on these conditions:
            # Empty/very short line (2 in a row = end of response)
            if ($trimmed.Length -le 1) {
                $emptyLineCount++
                if ($emptyLineCount -ge 2) { $capturing = $false }
                continue
            }
            $emptyLineCount = 0

            # Stop at prompt line
            if ($trimmed -match '^>\s') { $capturing = $false; continue }

            # Stop at TUI status bar keywords
            if ($trimmed -match 'Opus|Sonnet|Haiku|context\)|MCP server|bypass perm|shift.tab|\.PowerShell|\.Documents|\.Desktop') {
                $capturing = $false; continue
            }

            # Stop at TUI noise (box-drawing, block elements)
            if (Test-IsTuiNoise $trimmed) { $capturing = $false; continue }

            $clean = Remove-TuiChars $trimmed
            $clean = $clean.Trim()

            # Skip single-char garbage (x, Ϙ, etc.)
            if ($clean.Length -le 2) { continue }

            # Skip lines that are mostly non-letter chars
            $letterCount = ($clean -replace '[^a-zA-Z\u00C0-\u00FF]', '').Length
            if ($clean.Length -gt 5 -and $letterCount -lt ($clean.Length / 3)) { continue }

            $responseLines += $clean
        }
    }

    if ($responseLines.Count -gt 0) {
        $result = $responseLines -join "`n"
        if ($result.Length -gt 3800) { $result = $result.Substring(0, 3800) + "`n[... tronque]" }
        return $result.Trim()
    }

    # ---- STRATEGY 2: diff-based fallback ----
    $baseLines = $BaselineRaw -split "`r?`n" | ForEach-Object { $_.TrimEnd() }
    $currLines = $cleanedCurrent -split "`r?`n" | ForEach-Object { $_.TrimEnd() }

    $baseSet = @{}
    foreach ($line in $baseLines) {
        $t = $line.Trim()
        if ($t.Length -gt 0) { $baseSet[$t] = $true }
    }

    $newLines = @()
    foreach ($line in $currLines) {
        $t = $line.Trim()
        if ($t.Length -lt 3) { continue }
        if ($baseSet.ContainsKey($t)) { continue }
        if (Test-IsTuiNoise $t) { continue }
        $clean = Remove-TuiChars $t
        $clean = $clean -replace '^\u25CF\s*', ''
        $clean = $clean.Trim()
        if ($clean.Length -lt 3) { continue }
        if ($clean -match '^>\s*') { continue }
        $newLines += $clean
    }

    if ($newLines.Count -gt 0) {
        $result = $newLines -join "`n"
        if ($result.Length -gt 3800) { $result = $result.Substring(0, 3800) + "`n[... tronque]" }
        return $result.Trim()
    }

    return $null
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

    # Block elements (logo) -- U+2580-U+259F
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
    # Remove Latin Extended artifacts (not used in French/Dutch)
    $t = $t -replace '[\u0100-\u024F]', ''
    $t = $t -replace '[\u00D8-\u00DF]', ''
    # Remove Greek/Coptic artifacts
    $t = $t -replace '[\u0370-\u03FF]', ''
    # Remove Cyrillic artifacts
    $t = $t -replace '[\u0400-\u04FF]', ''
    # Remove Modifier Letters (˸ = U+02F8, etc.)
    $t = $t -replace '[\u02B0-\u02FF]', ''
    # Remove Spacing Modifier Letters
    $t = $t -replace '[\u0250-\u02AF]', ''
    # Remove stray single non-text chars at end of line
    $t = $t -replace '\s*[^\x20-\x7E\u00E0-\u00F6\u00F8-\u00FF]+\s*$', ''
    # Clean up multiple spaces
    $t = $t -replace '\s{2,}', ' '

    return $t.Trim()
}

# Final cleanup of extracted response before sending to Telegram.
# Cuts at status bar markers and removes remaining TUI garbage.
# Works on the FULL response string (not per-line).
function Clean-FinalResponse {
    param([string]$Text)

    if (-not $Text) { return "" }

    # Cut everything starting from known TUI status bar patterns
    # These patterns mark the end of Claude's actual response
    $cutPatterns = @(
        '\s*>?\s*Opus\s',
        '\s*>?\s*Sonnet\s',
        '\s*>?\s*Haiku\s',
        '\s*\d+M context',
        '\s*MCP server',
        '\s*bypass perm',
        '\s*shift.tab',
        '\bOption\d',
        '\bOpton\d',
        '\d+%\s*(1\s)?MCP',
        '\bPowerShell\b',
        '\b/mcp\b',
        '\bneeds auth\b'
    )
    foreach ($pat in $cutPatterns) {
        if ($Text -match $pat) {
            $idx = $Text.IndexOf(($Matches[0]))
            if ($idx -gt 5) {
                $Text = $Text.Substring(0, $idx)
            }
        }
    }

    # Remove isolated single characters surrounded by spaces (cursor artifacts: x, ˸, etc.)
    $Text = $Text -replace '(?<=\s)[^\s\w](?=\s)', ''
    $Text = $Text -replace '\s+\w\s+\w\s+\w\s*$', ''

    # Remove stray non-ASCII non-French chars
    $Text = $Text -replace '[\u0100-\u02FF]', ''
    $Text = $Text -replace '[\u0370-\u04FF]', ''

    # Clean up whitespace
    $Text = $Text -replace '\s{2,}', ' '
    $Text = $Text -replace '\s+$', ''

    return $Text.Trim()
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
# VOICE TRANSCRIPTION (faster-whisper, auto language detection, 99+ languages)
# ============================================================================

function Initialize-Transcription {
    # Check Python
    $py = $null
    foreach ($cmd in @("python", "python3", "py")) {
        try {
            $ver = & $cmd --version 2>&1
            if ($ver -match "Python\s+3\.") { $py = $cmd; break }
        } catch { }
    }
    if (-not $py) {
        Write-Host "[ClaudePlus] ERREUR: Python 3.x requis pour la transcription vocale." -ForegroundColor Red
        Write-Host "[ClaudePlus] Installez Python: https://www.python.org/downloads/" -ForegroundColor Yellow
        return $false
    }
    Write-Host "[ClaudePlus] Python OK: $py" -ForegroundColor DarkGray

    # Check/install faster-whisper
    $checkWhisper = (& $py -c "import faster_whisper; print('ok')" 2>$null) | Select-Object -Last 1
    $checkWhisper = "$checkWhisper".Trim()
    Write-Host "[ClaudePlus] Import check faster-whisper: '$checkWhisper'" -ForegroundColor DarkGray
    if ($checkWhisper -ne "ok") {
        Write-Host "[ClaudePlus] Installation de faster-whisper (premiere fois)..." -ForegroundColor Yellow

        # Try with --break-system-packages first (Python 3.11+), then without
        $pipResult = & $py -m pip install faster-whisper --break-system-packages 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Host "[ClaudePlus] Retry sans --break-system-packages..." -ForegroundColor DarkGray
            $pipResult = & $py -m pip install faster-whisper 2>&1
        }

        if ($LASTEXITCODE -ne 0) {
            Write-Host "[ClaudePlus] Pip output: $($pipResult | Out-String)" -ForegroundColor DarkGray
            Write-Host "[ClaudePlus] WARN: faster-whisper echoue, essai openai-whisper..." -ForegroundColor Yellow
            $pipResult = & $py -m pip install openai-whisper 2>&1
        }

        # Verify import
        $checkWhisper = (& $py -c "import faster_whisper; print('ok')" 2>$null) | Select-Object -Last 1
        $checkWhisper = "$checkWhisper".Trim()
        if ($checkWhisper -ne "ok") {
            # Check if openai-whisper was installed instead
            $checkOpenai = (& $py -c "import whisper; print('ok')" 2>$null) | Select-Object -Last 1
            $checkOpenai = "$checkOpenai".Trim()
            if ($checkOpenai -eq "ok") {
                $script:WhisperBackend = "openai"
                Write-Host "[ClaudePlus] openai-whisper OK (fallback)" -ForegroundColor DarkGray
            } else {
                Write-Host "[ClaudePlus] Pip output: $($pipResult | Out-String)" -ForegroundColor Red
                Write-Host "[ClaudePlus] ERREUR: Impossible d'installer whisper. Vocal desactive." -ForegroundColor Red
                Write-Host "[ClaudePlus] Essayez manuellement: $py -m pip install faster-whisper" -ForegroundColor Yellow
                return $false
            }
        } else {
            $script:WhisperBackend = "faster"
        }
    } else {
        $script:WhisperBackend = "faster"
    }
    Write-Host "[ClaudePlus] Whisper OK (backend: $($script:WhisperBackend))" -ForegroundColor DarkGray

    # Check ffmpeg (needed to decode OGG Opus from Telegram)
    $ffmpegOk = $false
    try {
        $ffVer = & ffmpeg -version 2>&1
        if ($ffVer -match "ffmpeg") { $ffmpegOk = $true }
    } catch { }

    if (-not $ffmpegOk) {
        Write-Host "[ClaudePlus] Installation de ffmpeg via winget..." -ForegroundColor Yellow
        try {
            & winget install Gyan.FFmpeg --accept-source-agreements --accept-package-agreements -q 2>&1 | Out-Null
            # Refresh PATH
            $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
            $ffVer = & ffmpeg -version 2>&1
            if ($ffVer -match "ffmpeg") { $ffmpegOk = $true }
        } catch { }
    }

    if (-not $ffmpegOk) {
        Write-Host "[ClaudePlus] WARN: ffmpeg non trouve. Certains formats audio peuvent echouer." -ForegroundColor Yellow
        Write-Host "[ClaudePlus] Installez ffmpeg: winget install Gyan.FFmpeg" -ForegroundColor Yellow
    } else {
        Write-Host "[ClaudePlus] ffmpeg OK" -ForegroundColor DarkGray
    }

    $script:PythonCmd = $py
    $script:TranscriptionReady = $true

    # Create the transcription Python script (supports both backends)
    $script:TranscribeScript = "$env:TEMP\claudeplus_transcribe.py"
    $pyCode = @'
import sys, json, os

audio_path = sys.argv[1]
model_size = sys.argv[2] if len(sys.argv) > 2 else "base"
backend = sys.argv[3] if len(sys.argv) > 3 else "faster"

if backend == "faster":
    from faster_whisper import WhisperModel
    model = WhisperModel(model_size, device="cpu", compute_type="int8")
    segments, info = model.transcribe(audio_path, beam_size=5)
    text = " ".join([s.text.strip() for s in segments])
    result = {"text": text, "language": info.language, "probability": round(info.language_probability, 2)}
else:
    import whisper
    model = whisper.load_model(model_size)
    out = model.transcribe(audio_path)
    result = {"text": out["text"].strip(), "language": out.get("language", "?"), "probability": 0.99}

print(json.dumps(result, ensure_ascii=False))
'@
    [System.IO.File]::WriteAllText($script:TranscribeScript, $pyCode, [System.Text.Encoding]::UTF8)

    Write-Host "[ClaudePlus] Transcription vocale prete (auto-detection langue, 99+ langues)" -ForegroundColor Green
    return $true
}

function Transcribe-TelegramAudio {
    param(
        [string]$FileId,
        [string]$Token,
        [string]$ModelSize = "base"
    )

    if (-not $script:TranscriptionReady) {
        Write-Host "[ClaudePlus] Transcription non initialisee" -ForegroundColor Red
        return $null
    }

    try {
        # Step 1: Get file path from Telegram
        $fileInfo = Invoke-RestMethod -Uri "https://api.telegram.org/bot$Token/getFile?file_id=$FileId" -Method Get -TimeoutSec 10
        if (-not $fileInfo.ok) {
            Write-Host "[ClaudePlus] Erreur getFile Telegram" -ForegroundColor Red
            return $null
        }
        $filePath = $fileInfo.result.file_path

        # Step 2: Download audio file
        $audioFile = "$env:TEMP\claudeplus_voice_$(Get-Random).ogg"
        $downloadUrl = "https://api.telegram.org/file/bot$Token/$filePath"
        Write-Host "[ClaudePlus] Telechargement audio: $filePath" -ForegroundColor DarkGray
        Invoke-WebRequest -Uri $downloadUrl -OutFile $audioFile -TimeoutSec 30 -ErrorAction Stop

        if (-not (Test-Path $audioFile) -or (Get-Item $audioFile).Length -lt 100) {
            Write-Host "[ClaudePlus] Fichier audio invalide" -ForegroundColor Red
            return $null
        }

        $fileSize = [Math]::Round((Get-Item $audioFile).Length / 1024, 1)
        Write-Host "[ClaudePlus] Audio telecharge ($fileSize KB), transcription en cours..." -ForegroundColor DarkCyan

        # Step 3: Transcribe with faster-whisper
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $script:PythonCmd
        $backend = if ($script:WhisperBackend) { $script:WhisperBackend } else { "faster" }
        $psi.Arguments = "`"$($script:TranscribeScript)`" `"$audioFile`" `"$ModelSize`" `"$backend`""
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

        $proc = [System.Diagnostics.Process]::Start($psi)
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        $exited = $proc.WaitForExit(120000)

        if (-not $exited) {
            try { $proc.Kill() } catch { }
            Write-Host "[ClaudePlus] Transcription timeout (2min)" -ForegroundColor Yellow
            return $null
        }

        $stdout = $stdoutTask.Result
        $stderr = $stderrTask.Result

        # Cleanup audio file
        Remove-Item $audioFile -ErrorAction SilentlyContinue

        if ($proc.ExitCode -ne 0 -or [string]::IsNullOrEmpty($stdout)) {
            Write-Host "[ClaudePlus] Erreur transcription: $stderr" -ForegroundColor Red
            return $null
        }

        # Parse JSON result
        $result = $stdout.Trim() | ConvertFrom-Json
        Write-Host "[ClaudePlus] Transcription OK: [$($result.language) $($result.probability * 100)%] $($result.text)" -ForegroundColor Green
        return $result

    } catch {
        Write-Host "[ClaudePlus] Erreur transcription: $_" -ForegroundColor Red
        return $null
    }
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
# ============================================================================
# PIPE MODE with STREAMING: Send message via 'claude -p --output-format stream-json'
# Reads JSON events line by line, sends real-time tool updates to Telegram,
# accumulates text deltas for the final response.
# No TUI, no buffer reading, no cursor artifacts, no size limit.
# Uses --continue to maintain conversation context across messages.
# ============================================================================

function Invoke-ClaudePipe {
    param(
        [string]$Message,
        [string]$ClaudePath,
        [string]$WorkDir,
        [switch]$Continue,
        [switch]$DangerouslySkipPermissions,
        [string]$TelegramToken,
        [string]$TelegramChatId
    )

    $escapedMsg = $Message -replace '"', '\"'
    $argList = @("-p", "`"$escapedMsg`"", "--output-format", "stream-json")
    if ($Continue) { $argList += "--continue" }
    if ($DangerouslySkipPermissions) { $argList += "--dangerously-skip-permissions" }
    $argStr = $argList -join " "

    Write-Host "[ClaudePlus] Stream: claude -p --output-format stream-json $(if($Continue){'--continue '})" -ForegroundColor DarkCyan

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ClaudePath
        $psi.Arguments = $argStr
        $psi.WorkingDirectory = $WorkDir
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

        Write-Host "[ClaudePlus] Stream: lancement..." -ForegroundColor DarkGray
        $proc = [System.Diagnostics.Process]::Start($psi)

        # Read stderr async to prevent deadlock
        $stderrTask = $proc.StandardError.ReadToEndAsync()

        # State tracking
        $textBuilder = New-Object System.Text.StringBuilder
        $toolsUsed = @()
        $lastTelegramUpdate = [DateTime]::MinValue
        $telegramUpdateInterval = 3
        $startTime = Get-Date
        $toolCount = 0
        $lineCount = 0
        $debugLogFile = "$env:TEMP\claudeplus_stream_debug.log"
        Remove-Item $debugLogFile -ErrorAction SilentlyContinue

        # Emoji map for tool types
        $toolEmojis = @{
            "Read" = [char]0x1F4D6; "Write" = [char]0x270F; "Edit" = [char]0x2702
            "Bash" = [char]0x2699; "Grep" = [char]0x1F50D; "Glob" = [char]0x1F4C2
            "Search" = [char]0x1F50E; "Agent" = [char]0x1F916; "Explore" = [char]0x1F9ED
            "TodoWrite" = [char]0x1F4DD; "WebSearch" = [char]0x1F310; "WebFetch" = [char]0x1F310
        }

        # Helper: recursively find tool names and text in any JSON structure
        function Find-InJson {
            param($obj, [string]$path = "")
            if ($null -eq $obj) { return @() }
            $found = @()

            if ($obj -is [System.Management.Automation.PSCustomObject]) {
                $props = $obj.PSObject.Properties
                foreach ($p in $props) {
                    $found += Find-InJson -obj $p.Value -path "$path.$($p.Name)"
                }
            }
            elseif ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) {
                $idx = 0
                foreach ($item in $obj) {
                    $found += Find-InJson -obj $item -path "$path[$idx]"
                    $idx++
                }
            }
            return $found
        }

        # Read stdout line by line (NDJSON stream)
        while (-not $proc.StandardOutput.EndOfStream) {
            $line = $proc.StandardOutput.ReadLine()
            if (-not $line -or $line.Trim().Length -eq 0) { continue }
            $lineCount++

            # Debug: log first 10 raw lines to file
            if ($lineCount -le 10) {
                $debugLine = "LINE $lineCount : $($line.Substring(0, [Math]::Min(500, $line.Length)))"
                Add-Content -Path $debugLogFile -Value $debugLine -Encoding UTF8
                Write-Host "[ClaudePlus] DEBUG L$lineCount : $($line.Substring(0, [Math]::Min(150, $line.Length)))" -ForegroundColor DarkGray
            }

            try {
                $event = $line | ConvertFrom-Json -ErrorAction Stop
            } catch {
                # Not JSON — might be plain text response (fallback)
                [void]$textBuilder.Append($line)
                [void]$textBuilder.Append("`n")
                continue
            }

            # ================================================================
            # STRATEGY: Try all known Claude Code CLI stream-json formats
            # The CLI format may differ from the raw API format
            # ================================================================

            $jsonStr = $line.ToLower()

            # --- DETECT TOOL USE (any format) ---
            # Look for tool_name, name with tool context, tool_use type
            $detectedTool = $null

            # Format A: API-style {"type":"content_block_start","content_block":{"type":"tool_use","name":"Read"}}
            if ($event.content_block -and $event.content_block.type -eq "tool_use" -and $event.content_block.name) {
                $detectedTool = $event.content_block.name
            }
            # Format B: {"type":"tool_use","name":"Read","input":{...}}
            if (-not $detectedTool -and $event.type -eq "tool_use" -and $event.name) {
                $detectedTool = $event.name
            }
            # Format C: Wrapped {"event":{"content_block":{"type":"tool_use","name":"Read"}}}
            if (-not $detectedTool -and $event.event -and $event.event.content_block -and $event.event.content_block.name) {
                $detectedTool = $event.event.content_block.name
            }
            # Format D: {"tool_name":"Read"} or {"tool":"Read"}
            if (-not $detectedTool -and $event.tool_name) { $detectedTool = $event.tool_name }
            if (-not $detectedTool -and $event.tool) { $detectedTool = $event.tool }
            # Format E: {"type":"assistant","tool_use":{"name":"Read"}}
            if (-not $detectedTool -and $event.tool_use -and $event.tool_use.name) { $detectedTool = $event.tool_use.name }
            # Format F: Claude Code subagent {"type":"system","subtype":"tool_use",...,"tool":{"name":"Read"}}
            if (-not $detectedTool -and $event.subtype -eq "tool_use" -and $event.tool -and $event.tool.name) { $detectedTool = $event.tool.name }

            if ($detectedTool) {
                $toolCount++
                $emoji = $toolEmojis[$detectedTool]
                if (-not $emoji) { $emoji = [char]0x26A1 }

                # Try to get tool input preview
                $inputPreview = ""
                $toolInput = $null
                if ($event.content_block -and $event.content_block.input) { $toolInput = $event.content_block.input }
                if (-not $toolInput -and $event.input) { $toolInput = $event.input }
                if (-not $toolInput -and $event.tool_use -and $event.tool_use.input) { $toolInput = $event.tool_use.input }
                if (-not $toolInput -and $event.tool -and $event.tool.input) { $toolInput = $event.tool.input }

                if ($toolInput) {
                    $preview = $toolInput.pattern
                    if (-not $preview) { $preview = $toolInput.command }
                    if (-not $preview) { $preview = $toolInput.file_path }
                    if (-not $preview) { $preview = $toolInput.path }
                    if (-not $preview) { $preview = $toolInput.description }
                    if (-not $preview) { $preview = $toolInput.query }
                    if ($preview) {
                        if ($preview.Length -gt 50) { $preview = $preview.Substring(0, 47) + "..." }
                        $inputPreview = ": $preview"
                    }
                }

                $toolDisplay = "$emoji $detectedTool$inputPreview"
                $toolsUsed += $toolDisplay
                Write-Host "[ClaudePlus] Stream: Tool #$toolCount -> $detectedTool$inputPreview" -ForegroundColor Magenta

                # Throttled Telegram progress update
                $now = Get-Date
                $elapsed = [int]($now - $startTime).TotalSeconds
                if ($TelegramToken -and $TelegramChatId) {
                    $secsSinceUpdate = ($now - $lastTelegramUpdate).TotalSeconds
                    if ($secsSinceUpdate -ge $telegramUpdateInterval -or $toolCount -eq 1) {
                        $progressMsg = "$([char]0x23F3) Claude travaille... (${elapsed}s)`n"
                        foreach ($t in $toolsUsed) { $progressMsg += "  $t`n" }
                        Send-TelegramMessage -Message $progressMsg.TrimEnd() -Token $TelegramToken -ChatId $TelegramChatId
                        $lastTelegramUpdate = $now
                    }
                }
            }

            # --- DETECT TEXT CONTENT (any format) ---
            # Format A: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Hello"}}
            if ($event.delta -and $event.delta.type -eq "text_delta" -and $event.delta.text) {
                [void]$textBuilder.Append($event.delta.text)
            }
            # Format B: {"type":"text","text":"Hello"}
            elseif ($event.type -eq "text" -and $event.text) {
                [void]$textBuilder.Append($event.text)
            }
            # Format C: Wrapped {"event":{"delta":{"type":"text_delta","text":"Hello"}}}
            elseif ($event.event -and $event.event.delta -and $event.event.delta.text) {
                [void]$textBuilder.Append($event.event.delta.text)
            }
            # Format D: {"type":"assistant","message":{"content":[{"type":"text","text":"Hello"}]}}
            elseif ($event.message -and $event.message.content) {
                foreach ($block in $event.message.content) {
                    if ($block.type -eq "text" -and $block.text) {
                        [void]$textBuilder.Append($block.text)
                    }
                }
            }
            # Format E: {"type":"result","result":"Hello full response"}
            elseif ($event.type -eq "result" -and $event.result) {
                [void]$textBuilder.Append($event.result)
            }
            # Format F: {"content":"Hello"} or {"response":"Hello"}
            elseif ($event.content -and $event.content -is [string] -and $event.content.Length -gt 2) {
                [void]$textBuilder.Append($event.content)
            }
            elseif ($event.response -and $event.response -is [string]) {
                [void]$textBuilder.Append($event.response)
            }
        }

        # Wait for process to finish
        $exited = $proc.WaitForExit(30000)
        if (-not $exited) { try { $proc.Kill() } catch { } }

        $stderr = $stderrTask.Result
        $exitCode = if ($proc.HasExited) { $proc.ExitCode } else { -1 }
        $elapsed = [int]((Get-Date) - $startTime).TotalSeconds

        Write-Host "[ClaudePlus] Stream: exit=$exitCode, ${elapsed}s, $lineCount lines, $toolCount tools, $($textBuilder.Length) chars text" -ForegroundColor DarkGray

        if ($stderr -and $stderr.Trim().Length -gt 0) {
            $stderrPreview = $stderr.Substring(0, [Math]::Min(300, $stderr.Length))
            Write-Host "[ClaudePlus] Stream stderr: $stderrPreview" -ForegroundColor DarkGray
        }

        # Send final progress summary if tools were used
        if ($TelegramToken -and $TelegramChatId -and $toolCount -gt 0) {
            $summaryMsg = "$([char]0x2705) Termine (${elapsed}s, $toolCount outils)`n"
            foreach ($t in $toolsUsed) { $summaryMsg += "  $t`n" }
            Send-TelegramMessage -Message $summaryMsg.TrimEnd() -Token $TelegramToken -ChatId $TelegramChatId
        }

        $result = $textBuilder.ToString().Trim()
        if ($result.Length -gt 0) {
            Write-Host "[ClaudePlus] Stream: reponse OK ($($result.Length) chars)" -ForegroundColor Green
            return $result
        } else {
            Write-Host "[ClaudePlus] Stream: reponse vide, $lineCount lignes lues. Voir debug: $debugLogFile" -ForegroundColor Yellow
            Write-Host "[ClaudePlus] Stream: fallback vers pipe mode classique..." -ForegroundColor Yellow
            return Invoke-ClaudePipePlain -Message $Message -ClaudePath $ClaudePath -WorkDir $WorkDir -Continue:$Continue -DangerouslySkipPermissions:$DangerouslySkipPermissions
        }
    } catch {
        Write-Host "[ClaudePlus] Stream erreur: $_ — fallback pipe plain" -ForegroundColor Red
        return Invoke-ClaudePipePlain -Message $Message -ClaudePath $ClaudePath -WorkDir $WorkDir -Continue:$Continue -DangerouslySkipPermissions:$DangerouslySkipPermissions
    }
}

# Fallback plain pipe mode (no streaming, no Telegram updates)
function Invoke-ClaudePipePlain {
    param(
        [string]$Message,
        [string]$ClaudePath,
        [string]$WorkDir,
        [switch]$Continue,
        [switch]$DangerouslySkipPermissions
    )

    $escapedMsg = $Message -replace '"', '\"'
    $argList = @("-p", "`"$escapedMsg`"")
    if ($Continue) { $argList += "--continue" }
    if ($DangerouslySkipPermissions) { $argList += "--dangerously-skip-permissions" }
    $argStr = $argList -join " "

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $ClaudePath
        $psi.Arguments = $argStr
        $psi.WorkingDirectory = $WorkDir
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8

        $proc = [System.Diagnostics.Process]::Start($psi)
        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()
        $exited = $proc.WaitForExit(180000)
        if (-not $exited) { try { $proc.Kill() } catch { }; return $null }

        $result = $stdoutTask.Result.Trim()
        if ($result.Length -gt 0) { return $result }
        $stderr = $stderrTask.Result.Trim()
        if ($stderr.Length -gt 10) { return $stderr }
        return $null
    } catch {
        Write-Host "[ClaudePlus] PipePlain erreur: $_" -ForegroundColor Red
        return $null
    }
}

# ============================================================================
# SEND TEXT TO CLAUDE WINDOW (legacy - used for TUI terminal mode)
# Uses: MainWindowHandle + SetForegroundWindow + WshShell.SendKeys
# ============================================================================

function Send-TextToClaude {
    param([string]$Text)

    # Find window handle if needed
    if (-not $script:ClaudeWindowHandle -or $script:ClaudeWindowHandle -eq [IntPtr]::Zero) {
        if ($script:ClaudeCmdPid -gt 0) {
            $script:ClaudeWindowHandle = Get-WindowHandleForPid -Pid1 $script:ClaudeCmdPid
        }
    }

    $hwnd = $script:ClaudeWindowHandle
    if (-not $hwnd -or $hwnd -eq [IntPtr]::Zero) {
        Write-Host "[ClaudePlus] ERREUR: Handle fenetre introuvable" -ForegroundColor Red
        return $false
    }

    Write-Host "[ClaudePlus] Envoi via PostMessage WM_CHAR vers hwnd=$hwnd..." -ForegroundColor DarkGray

    try {
        # PostMessage WM_CHAR approach -- no AttachConsole needed, works with conhost.exe windows
        # Builds a C# helper that posts chars directly to the window message queue
        $csFile = "$env:TEMP\claudeplus_poster.cs"
        $csCode = @"
using System;
using System.Runtime.InteropServices;
using System.Threading;
public class ConPoster {
    const uint WM_CHAR    = 0x0102;
    const uint WM_KEYDOWN = 0x0100;
    const uint WM_KEYUP   = 0x0101;
    const int  VK_RETURN  = 0x0D;
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr wp, IntPtr lp);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hwnd);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, IntPtr pid);
    [DllImport("user32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint from, uint to, bool attach);
    public static string Send(long hwnd, string text) {
        var h = new IntPtr(hwnd);
        // Bring window to foreground using AttachThreadInput trick
        try {
            ShowWindow(h, 9); // SW_RESTORE
            BringWindowToTop(h);
            uint wndThread = GetWindowThreadProcessId(h, IntPtr.Zero);
            uint curThread = GetCurrentThreadId();
            AttachThreadInput(curThread, wndThread, true);
            SetForegroundWindow(h);
            AttachThreadInput(curThread, wndThread, false);
        } catch {}
        Thread.Sleep(400);
        // Send each character via WM_CHAR
        int n = 0;
        foreach (char c in text) {
            PostMessage(h, WM_CHAR, (IntPtr)c, (IntPtr)1);
            Thread.Sleep(30);
            n++;
        }
        Thread.Sleep(150);
        // Send Enter
        PostMessage(h, WM_KEYDOWN, (IntPtr)VK_RETURN, (IntPtr)1);
        Thread.Sleep(30);
        PostMessage(h, WM_KEYUP,   (IntPtr)VK_RETURN, (IntPtr)1);
        return "OK:sent=" + n;
    }
}
"@
        $textFile = "$env:TEMP\claudeplus_sendtext.txt"
        $logFile  = "$textFile.log"
        [System.IO.File]::WriteAllText($csFile, $csCode, [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText($textFile, $Text, [System.Text.Encoding]::UTF8)
        if (Test-Path $logFile) { Remove-Item $logFile -ErrorAction SilentlyContinue }

        $hwndLong = [long]$hwnd
        $helperScript = "$env:TEMP\claudeplus_poster.ps1"
        $helperLines = @(
            'param([long]$Hwnd, [string]$TextFile, [string]$CsFile, [string]$LogFile)',
            'try {',
            '    $text = [System.IO.File]::ReadAllText($TextFile, [System.Text.Encoding]::UTF8)',
            '    Add-Type -Path $CsFile -ErrorAction Stop',
            '    $result = [ConPoster]::Send($Hwnd, $text)',
            '    [System.IO.File]::WriteAllText($LogFile, "Result=$result")',
            '} catch {',
            '    [System.IO.File]::WriteAllText($LogFile, "EXCEPTION: $_")',
            '}'
        )
        [System.IO.File]::WriteAllText($helperScript, ($helperLines -join "`r`n"), [System.Text.Encoding]::UTF8)

        $argStr = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$helperScript`" -Hwnd $hwndLong -TextFile `"$textFile`" -CsFile `"$csFile`" -LogFile `"$logFile`""
        $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $argStr -PassThru -WindowStyle Hidden
        $proc.WaitForExit(15000) | Out-Null
        if (-not $proc.HasExited) { $proc.Kill() }

        if (Test-Path $logFile) {
            $logContent = [System.IO.File]::ReadAllText($logFile)
            Write-Host "[ClaudePlus] Send result: $logContent" -ForegroundColor DarkGray
            Remove-Item $logFile -ErrorAction SilentlyContinue
            if ($logContent -match "^Result=OK") { return $true }
        } else {
            Write-Host "[ClaudePlus] Pas de log (process crash?)" -ForegroundColor Red
        }
    } catch {
        Write-Host "[ClaudePlus] Erreur send: $_" -ForegroundColor Red
    }

    return $false
}

# ============================================================================
# MULTI-SESSION MANAGEMENT (@prefix routing)
# Each terminal registers itself with a name. Telegram messages prefixed
# with @name are routed to the matching terminal. No prefix = default session.
# Session registry: %LOCALAPPDATA%\ClaudePlus\sessions\<name>.json
# ============================================================================

function Register-Session {
    param([string]$Name)
    if (-not $Name) { return }
    if (-not (Test-Path $script:SessionRegistryDir)) {
        New-Item -ItemType Directory -Path $script:SessionRegistryDir -Force | Out-Null
    }
    $sessionFile = Join-Path $script:SessionRegistryDir "$Name.json"
    $sessionData = @{
        Name = $Name
        Pid = $PID
        StartTime = (Get-Date).ToString("o")
        WorkDir = (Get-Location).Path
    } | ConvertTo-Json
    [System.IO.File]::WriteAllText($sessionFile, $sessionData, [System.Text.Encoding]::UTF8)
    $script:SessionName = $Name
    Write-Host "[ClaudePlus] Session '$Name' enregistree (PID=$PID)" -ForegroundColor Green
}

function Unregister-Session {
    param([string]$Name)
    if (-not $Name) { return }
    $sessionFile = Join-Path $script:SessionRegistryDir "$Name.json"
    if (Test-Path $sessionFile) {
        Remove-Item $sessionFile -ErrorAction SilentlyContinue
    }
}

function Get-ActiveSessions {
    if (-not (Test-Path $script:SessionRegistryDir)) { return @() }
    $sessions = @()
    Get-ChildItem -Path $script:SessionRegistryDir -Filter "*.json" -ErrorAction SilentlyContinue | ForEach-Object {
        try {
            $data = Get-Content $_.FullName -Raw | ConvertFrom-Json
            # Check if process is still alive
            $proc = Get-Process -Id $data.Pid -ErrorAction SilentlyContinue
            if ($proc) {
                $sessions += $data
            } else {
                # Dead session, clean up
                Remove-Item $_.FullName -ErrorAction SilentlyContinue
            }
        } catch {
            Remove-Item $_.FullName -ErrorAction SilentlyContinue
        }
    }
    return $sessions
}

# Parse @prefix from message. Returns @{Target="name"; Text="actual message"} or @{Target=$null; Text="original"}
function Parse-MessageTarget {
    param([string]$Message)
    if ($Message -match '^@(\S+)\s+(.+)$') {
        return @{ Target = $Matches[1].ToLower(); Text = $Matches[2] }
    }
    if ($Message -match '^@(\S+)$') {
        # Just "@name" with no message — ignore
        return @{ Target = $Matches[1].ToLower(); Text = "" }
    }
    return @{ Target = $null; Text = $Message }
}

# Check if this session should handle the message
function Test-MessageForMe {
    param([string]$RawText)

    $parsed = Parse-MessageTarget -Message $RawText

    # Commands like /stop, /list are always handled by everyone
    if ($parsed.Text -match '^/(stop|list|sessions|help)') {
        return @{ ShouldHandle = $true; Text = $parsed.Text; IsCommand = $true }
    }

    # If message has @prefix
    if ($parsed.Target) {
        $myName = $script:SessionName
        if (-not $myName) { return @{ ShouldHandle = $false; Text = ""; IsCommand = $false } }

        # Check if target matches my name (case-insensitive)
        if ($parsed.Target -eq $myName.ToLower()) {
            return @{ ShouldHandle = $true; Text = $parsed.Text; IsCommand = $false }
        }
        # Partial match (e.g., @fis matches "fiscal")
        if ($myName.ToLower().StartsWith($parsed.Target)) {
            return @{ ShouldHandle = $true; Text = $parsed.Text; IsCommand = $false }
        }
        return @{ ShouldHandle = $false; Text = ""; IsCommand = $false }
    }

    # No @prefix: only handle if I'm the only session, or if I'm the default (first registered)
    $activeSessions = Get-ActiveSessions
    if ($activeSessions.Count -le 1) {
        # I'm the only one — handle it
        return @{ ShouldHandle = $true; Text = $parsed.Text; IsCommand = $false }
    }

    # Multiple sessions active: only the "default" (lowest PID = oldest) handles unprefixed messages
    $minPid = ($activeSessions | Sort-Object { [int]$_.Pid } | Select-Object -First 1).Pid
    if ([int]$minPid -eq $PID) {
        return @{ ShouldHandle = $true; Text = $parsed.Text; IsCommand = $false }
    }

    # Not for me
    return @{ ShouldHandle = $false; Text = ""; IsCommand = $false }
}

# ============================================================================
# MAIN COMMAND: claudeplus
# ============================================================================

function Invoke-ClaudePlus {
    [CmdletBinding()]
    param(
        [string]$Name,
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

    # Auto-generate session name from current folder if not provided
    if (-not $Name) {
        $Name = (Split-Path -Leaf (Get-Location).Path).ToLower() -replace '[^a-z0-9]', ''
        if (-not $Name -or $Name.Length -lt 2) { $Name = "claude" }
    }
    $Name = $Name.ToLower()

    Write-Host ""
    Write-Host "  +============================================+" -ForegroundColor DarkCyan
    Write-Host "  |         ClaudePlus - FiscalIQ               |" -ForegroundColor DarkCyan
    Write-Host "  |   Claude Code + Telegram Mirror             |" -ForegroundColor DarkCyan
    Write-Host "  |   Session: @$Name                           " -ForegroundColor DarkCyan
    Write-Host "  +============================================+" -ForegroundColor DarkCyan
    Write-Host ""

    $useTelegram = (-not $NoTelegram -and $config.AutoTelegram -and $config.TelegramBotToken -and $config.TelegramChatId)

    if ($useTelegram) {
        # Register this session
        Register-Session -Name $Name
        $activeSessions = Get-ActiveSessions
        $sessionCount = $activeSessions.Count

        Write-Host "[ClaudePlus] Mode Mirror Telegram actif. Session '@$Name' ($sessionCount active(s))" -ForegroundColor Green
        if ($sessionCount -gt 1) {
            Write-Host "[ClaudePlus] Sessions actives:" -ForegroundColor Cyan
            foreach ($s in $activeSessions) {
                $marker = if ($s.Pid -eq $PID) { " (moi)" } else { "" }
                Write-Host "  @$($s.Name)$marker — $($s.WorkDir)" -ForegroundColor White
            }
            Write-Host "[ClaudePlus] Utilisez @$Name <message> pour cibler cette session" -ForegroundColor Yellow
        }

        # RESET all state from previous session
        $script:ClaudeProcess = $null
        $script:ClaudeCmdPid = 0
        $script:ClaudeWindowHandle = [IntPtr]::Zero
        $script:LastConsoleText = ""
        $script:PipeMessageCount = 0

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

        # Launch via Windows Terminal (wt.exe) if available, fallback to conhost.exe
        $wtPath = (Get-Command wt.exe -ErrorAction SilentlyContinue).Source
        if ($wtPath) {
            Write-Host "[ClaudePlus] Lancement via Windows Terminal (wt.exe)..." -ForegroundColor Cyan
            $script:ClaudeProcess = Start-Process -FilePath $wtPath -ArgumentList "--title `"ClaudePlus @$Name`" cmd.exe /c `"$batPath`"" -PassThru
            $script:UsesWindowsTerminal = $true
        } else {
            $conhost = "$env:SystemRoot\System32\conhost.exe"
            if (Test-Path $conhost) {
                Write-Host "[ClaudePlus] Lancement via conhost.exe..." -ForegroundColor Cyan
                $script:ClaudeProcess = Start-Process -FilePath $conhost -ArgumentList "cmd.exe /c `"$batPath`"" -PassThru
            } else {
                $script:ClaudeProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$batPath`"" -PassThru
            }
            $script:UsesWindowsTerminal = $false
        }

        Write-Host "[ClaudePlus] PID: $($script:ClaudeProcess.Id) ($(if($script:UsesWindowsTerminal){'Windows Terminal'}else{'conhost'}))" -ForegroundColor DarkGray

        # Wait and find the terminal window handle
        Write-Host "[ClaudePlus] Recherche de la fenetre Claude..." -ForegroundColor DarkGray
        $found = $false

        if ($script:UsesWindowsTerminal) {
            # Windows Terminal: wt.exe spawns WindowsTerminal.exe which owns the window
            # The window title contains our custom title "ClaudePlus @name"
            for ($i = 0; $i -lt 30; $i++) {
                Start-Sleep -Milliseconds 500
                # Find WindowsTerminal process with our title
                Get-Process -Name "WindowsTerminal" -ErrorAction SilentlyContinue | ForEach-Object {
                    if ($_.MainWindowHandle -ne [IntPtr]::Zero -and -not $found) {
                        if ($_.MainWindowTitle -match "ClaudePlus" -or $_.MainWindowTitle -match $Name) {
                            $script:ClaudeWindowHandle = $_.MainWindowHandle
                            $found = $true
                            Write-Host "[ClaudePlus] FENETRE Windows Terminal TROUVEE! Handle=$($_.MainWindowHandle) Title='$($_.MainWindowTitle)'" -ForegroundColor Green
                        }
                    }
                }
                if ($found) { break }

                # Also try: WT may reuse existing window, find cmd.exe child
                if ($script:ClaudeCmdPid -eq 0) {
                    # WT process tree: wt.exe -> OpenConsole.exe -> cmd.exe -> claude
                    try {
                        $wtChildren = Get-CimInstance Win32_Process -Filter "ParentProcessId=$($script:ClaudeProcess.Id)" -ErrorAction SilentlyContinue
                        foreach ($wc in $wtChildren) {
                            $wcChildren = Get-CimInstance Win32_Process -Filter "ParentProcessId=$($wc.ProcessId)" -ErrorAction SilentlyContinue
                            foreach ($wcc in $wcChildren) {
                                if ($wcc.Name -eq "cmd.exe") { $script:ClaudeCmdPid = [int]$wcc.ProcessId }
                            }
                            if ($wc.Name -eq "cmd.exe") { $script:ClaudeCmdPid = [int]$wc.ProcessId }
                        }
                    } catch { }
                }
            }

            # If we didn't find by title, try any WT window that appeared after our launch
            if (-not $found) {
                Get-Process -Name "WindowsTerminal" -ErrorAction SilentlyContinue | ForEach-Object {
                    if ($_.MainWindowHandle -ne [IntPtr]::Zero -and -not $found) {
                        $script:ClaudeWindowHandle = $_.MainWindowHandle
                        $found = $true
                        Write-Host "[ClaudePlus] FENETRE WT trouvee (fallback)! Handle=$($_.MainWindowHandle)" -ForegroundColor Yellow
                    }
                }
            }
        } else {
            # conhost.exe: classic approach
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Milliseconds 500

                if ($script:ClaudeCmdPid -eq 0) {
                    $script:ClaudeCmdPid = Find-ChildCmdPid -ParentPid $script:ClaudeProcess.Id
                }

                if ($script:ClaudeCmdPid -gt 0) {
                    $hwnd = Get-WindowHandleForPid -Pid1 $script:ClaudeCmdPid
                    if ($hwnd -ne [IntPtr]::Zero) {
                        $found = $true
                        $script:ClaudeWindowHandle = $hwnd
                        Write-Host "[ClaudePlus] FENETRE conhost TROUVEE! cmd PID=$($script:ClaudeCmdPid), Handle=$hwnd" -ForegroundColor Green
                        break
                    }
                }
            }
        }

        # Fallback scan for both modes
        if (-not $found) {
            Write-Host "[ClaudePlus] Fenetre non trouvee, tentative scan..." -ForegroundColor Yellow
            Get-Process -Name "cmd" -ErrorAction SilentlyContinue | ForEach-Object {
                if ($_.MainWindowHandle -ne [IntPtr]::Zero -and -not $found) {
                    try {
                        $wmi = Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" -ErrorAction SilentlyContinue
                        if ($wmi -and [int]$wmi.ParentProcessId -eq $script:ClaudeProcess.Id) {
                            $script:ClaudeCmdPid = $_.Id
                            $script:ClaudeWindowHandle = $_.MainWindowHandle
                            $found = $true
                            Write-Host "[ClaudePlus] FENETRE TROUVEE via scan! PID=$($_.Id) Handle=$($_.MainWindowHandle)" -ForegroundColor Green
                        }
                    } catch { }
                }
            }
        }

        if (-not $found) {
            Write-Host "[ClaudePlus] ATTENTION: Pas de handle fenetre. Le mirror TUI peut echouer (pipe mode OK)." -ForegroundColor Red
        }

        # Initialize voice transcription (Python + faster-whisper + ffmpeg)
        $voiceOk = Initialize-Transcription

        # Build startup message with session info
        $startupLines = @()
        $startupLines += "$([char]0x1F680) @$Name demarre!"
        $startupLines += "Dossier: $(Split-Path -Leaf $workDir)"
        if ($sessionCount -gt 1) {
            $otherNames = ($activeSessions | Where-Object { $_.Pid -ne $PID } | ForEach-Object { "@$($_.Name)" }) -join ", "
            $startupLines += "Autres sessions: $otherNames"
            $startupLines += ""
            $startupLines += "Prefixez avec @$Name pour cibler cette session"
        }
        if ($voiceOk) {
            $startupLines += "Texte et vocal OK"
        } else {
            $startupLines += "Texte OK (vocal desactive)"
        }
        $startupLines += "/stop pour arreter | /list pour voir les sessions"
        Send-TelegramMessage -Message ($startupLines -join "`n") -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

        # Wait for Claude TUI to start in the visible terminal
        Write-Host "[ClaudePlus] Attente demarrage Claude TUI (3s)..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 3

        # Telegram polling loop -- uses pipe mode (claude -p) for clean responses
        Write-Host "[ClaudePlus] Mode PIPE actif: les messages Telegram passent par 'claude -p' (pas par le terminal)" -ForegroundColor Cyan
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
        $pollCount = 0

        try {
            while (-not $stopRequested) {
                $pollCount++
                # Check if conhost or cmd.exe has exited
                if ($script:ClaudeProcess.HasExited) { Write-Host "[ClaudePlus] conhost a quitte." -ForegroundColor Yellow; break }
                if ($script:ClaudeCmdPid -gt 0) {
                    $cmdProc = Get-Process -Id $script:ClaudeCmdPid -ErrorAction SilentlyContinue
                    if (-not $cmdProc) { Write-Host "[ClaudePlus] cmd.exe termine, arret." -ForegroundColor Yellow; break }
                }

                # --- CHECK TELEGRAM ---
                try {
                    $url = "https://api.telegram.org/bot$($config.TelegramBotToken)/getUpdates?limit=10&timeout=2"
                    if ($lastUpdateId -gt 0) { $url += "&offset=$($lastUpdateId + 1)" }
                    if ($pollCount % 10 -eq 1) {
                        Write-Host "[ClaudePlus] Telegram poll #$pollCount (offset=$lastUpdateId, hwnd=$($script:ClaudeWindowHandle))" -ForegroundColor DarkGray
                    }
                    $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 5 -ErrorAction Stop

                    if ($response.ok -and $response.result) {
                        foreach ($update in $response.result) {
                            if ($update.update_id -gt $lastUpdateId) { $lastUpdateId = $update.update_id }
                            $msg = $update.message
                            if (-not $msg) { continue }
                            if ($msg.chat.id -ne [long]$config.TelegramChatId) { continue }

                            # Determine message type: text or voice
                            $text = $null
                            $isVoice = $false

                            if ($msg.voice -or $msg.audio) {
                                # Voice message or audio file
                                $isVoice = $true
                                $fileId = if ($msg.voice) { $msg.voice.file_id } else { $msg.audio.file_id }
                                $duration = if ($msg.voice) { $msg.voice.duration } else { $msg.audio.duration }

                                if (-not $script:TranscriptionReady) {
                                    Send-TelegramMessage -Message "[Vocal non supporte - Python/faster-whisper manquant]" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                    continue
                                }

                                Write-Host ""
                                Write-Host "  >> [Vocal ${duration}s] Transcription..." -ForegroundColor DarkCyan
                                Send-TelegramMessage -Message "[Transcription en cours...]" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

                                $transcription = Transcribe-TelegramAudio -FileId $fileId -Token $config.TelegramBotToken
                                if (-not $transcription -or [string]::IsNullOrEmpty($transcription.text)) {
                                    Send-TelegramMessage -Message "[Erreur transcription - audio non reconnu]" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                    continue
                                }

                                $text = $transcription.text.Trim()
                                $lang = $transcription.language
                                $conf = [int]($transcription.probability * 100)
                                Write-Host "  >> [Vocal -> $lang ${conf}%] $text" -ForegroundColor Cyan
                                Send-TelegramMessage -Message "[Transcription ($lang ${conf}%)]: $text" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

                            } elseif (-not [string]::IsNullOrEmpty($msg.text)) {
                                $text = $msg.text.Trim()
                            } else {
                                continue
                            }

                            if (-not $text -or $text.Length -eq 0) { continue }

                            # --- MULTI-SESSION ROUTING ---
                            $routing = Test-MessageForMe -RawText $text

                            # Handle /list command (show active sessions)
                            if ($routing.IsCommand -and $routing.Text -match '^/list|^/sessions') {
                                $sessions = Get-ActiveSessions
                                if ($sessions.Count -eq 0) {
                                    $listMsg = "Aucune session active"
                                } else {
                                    $listMsg = "$([char]0x1F4CB) Sessions actives ($($sessions.Count)):`n"
                                    foreach ($s in $sessions) {
                                        $me = if ($s.Pid -eq $PID) { " $([char]0x2190) ici" } else { "" }
                                        $listMsg += "  @$($s.Name) — $(Split-Path -Leaf $s.WorkDir)$me`n"
                                    }
                                    $listMsg += "`nPrefixez: @nom message"
                                }
                                Send-TelegramMessage -Message $listMsg.TrimEnd() -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }

                            if ($routing.IsCommand -and $routing.Text -match '^/stop') {
                                Send-TelegramMessage -Message "@$Name arrete." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                $stopRequested = $true
                                break
                            }

                            if (-not $routing.ShouldHandle) {
                                # Message is for another session, skip silently
                                continue
                            }

                            # Use the cleaned text (without @prefix)
                            $text = $routing.Text
                            if (-not $text -or $text.Length -eq 0) { continue }

                            # Skip if already waiting for a response
                            if ($waitingForResponse) {
                                Send-TelegramMessage -Message "[@$Name] ATTENTE — Claude est en train de repondre, patientez..." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }

                            if (-not $isVoice) {
                                Write-Host ""
                                Write-Host "  >> [$Name] $text" -ForegroundColor Cyan
                            }

                            $waitingForResponse = $true

                            # HYBRID: also send to TUI terminal for visual display
                            Send-TextToClaude -Text $text | Out-Null

                            # Use streaming pipe mode for clean response + real-time Telegram updates
                            $useSkip = ($config.DangerouslySkipPermissions -eq $true)
                            $responseText = Invoke-ClaudePipe `
                                -Message $text `
                                -ClaudePath $claudePath `
                                -WorkDir $workDir `
                                -Continue:($script:PipeMessageCount -gt 0) `
                                -DangerouslySkipPermissions:$useSkip `
                                -TelegramToken $config.TelegramBotToken `
                                -TelegramChatId $config.TelegramChatId

                            $script:PipeMessageCount++
                            $waitingForResponse = $false

                            if ($responseText -and $responseText.Length -gt 3) {
                                # Show response in PS console
                                $previewLen = [Math]::Min(200, $responseText.Length)
                                Write-Host "  << $($responseText.Substring(0, $previewLen))$(if($responseText.Length -gt 200){'...'})" -ForegroundColor Green
                                Write-Host ""

                                # Truncate if too long for Telegram (4096 char limit)
                                if ($responseText.Length -gt 3900) {
                                    $responseText = $responseText.Substring(0, 3900) + "`n[... tronque]"
                                }
                                # Send final response with separator and session tag
                                $finalMsg = "$([char]0x2500)$([char]0x2500) @$Name $([char]0x2500)$([char]0x2500)`n$responseText"
                                Send-TelegramMessage -Message $finalMsg -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            } else {
                                Write-Host "  << [pas de reponse]" -ForegroundColor Yellow
                                Write-Host ""
                                Send-TelegramMessage -Message "[@$Name] Pas de reponse - verifiez le terminal" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            }
                        }
                    }
                }
                catch { }

                Start-Sleep -Milliseconds 500
            }
        }
        finally {
            Unregister-Session -Name $Name
            Send-TelegramMessage -Message "@$Name terminee." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
            if (Test-Path $batPath) { Remove-Item $batPath -ErrorAction SilentlyContinue }
            Write-Host ""
            Write-Host "[ClaudePlus] Session '@$Name' terminee." -ForegroundColor DarkCyan
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
