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
$script:PendingMessage = $null  # Stores message awaiting session choice: @{Text; Type; ImagePaths; FilePaths; Timestamp}
$script:WaitingForSessionChoice = $false  # True when waiting for user to pick session number
$script:SessionChoiceMap = @{}  # Maps number -> session name for current choice

# No direct P/Invoke for console write - PowerShell shares its own console
# All WriteConsoleInput calls go through a helper process (see Send-TextToClaude)

$script:LastConsoleText = ""
$script:ReaderExe = $null
$script:PendingImagePaths = @()
$script:PendingFilePaths = @()
$script:LastPipeTools = @()
$script:LastPipeToolCount = 0
$script:LastPipeElapsed = 0
$script:TelegramVerbose = $true  # true = outils+progression+resultat, false = resultat seulement

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
    if ([string]::IsNullOrEmpty($Message) -or [string]::IsNullOrEmpty($Token) -or [string]::IsNullOrEmpty($ChatId)) {
        Write-Host "[ClaudePlus] TG SKIP: msg=$([string]::IsNullOrEmpty($Message)) tok=$([string]::IsNullOrEmpty($Token)) chat=$([string]::IsNullOrEmpty($ChatId))" -ForegroundColor Red
        return
    }
    if ($Message.Length -gt 4000) { $Message = $Message.Substring(0, 4000) + "`n[...]" }
    try {
        $body = @{ chat_id = $ChatId; text = $Message; disable_web_page_preview = "true" }
        $result = Invoke-RestMethod -Uri "https://api.telegram.org/bot$Token/sendMessage" -Method Post -Body $body -TimeoutSec 10 -ErrorAction Stop
        if (-not $result.ok) {
            Write-Host "[ClaudePlus] TG ERREUR: $($result | ConvertTo-Json -Compress)" -ForegroundColor Red
        }
    } catch {
        Write-Host "[ClaudePlus] TG EXCEPTION: $_" -ForegroundColor Red
    }
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
        return $false
    }
    Write-Host "[ClaudePlus] Python OK: $py" -ForegroundColor DarkGray

    # Check/install faster-whisper
    $checkWhisper = (& $py -c "import faster_whisper; print('ok')" 2>$null) | Select-Object -Last 1
    $checkWhisper = "$checkWhisper".Trim()
    Write-Host "[ClaudePlus] Import check faster-whisper: '$checkWhisper'" -ForegroundColor DarkGray
    if ($checkWhisper -ne "ok") {
        Write-Host "[ClaudePlus] Installation de faster-whisper (premiere fois)..." -ForegroundColor Yellow
        $pipResult = & $py -m pip install faster-whisper --break-system-packages 2>&1
        if ($LASTEXITCODE -ne 0) {
            $pipResult = & $py -m pip install faster-whisper 2>&1
        }
        if ($LASTEXITCODE -ne 0) {
            Write-Host "[ClaudePlus] WARN: faster-whisper echoue, essai openai-whisper..." -ForegroundColor Yellow
            & $py -m pip install openai-whisper --break-system-packages 2>&1 | Out-Null
        }
    }

    # Check/install PyAV (decode OGG/Opus nativement, pas besoin de ffmpeg)
    $checkAv = (& $py -c "import av; print('ok')" 2>$null) | Select-Object -Last 1
    $checkAv = "$checkAv".Trim()
    if ($checkAv -ne "ok") {
        Write-Host "[ClaudePlus] Installation de PyAV (conversion audio OGG/Opus)..." -ForegroundColor Yellow
        & $py -m pip install av --break-system-packages 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            & $py -m pip install av 2>&1 | Out-Null
        }
        $checkAv = (& $py -c "import av; print('ok')" 2>$null) | Select-Object -Last 1
        $checkAv = "$checkAv".Trim()
    }
    if ($checkAv -eq "ok") {
        Write-Host "[ClaudePlus] PyAV OK (conversion OGG/Opus native)" -ForegroundColor DarkGray
    } else {
        Write-Host "[ClaudePlus] WARN: PyAV non installe — installation ffmpeg comme fallback..." -ForegroundColor Yellow

        # Check/install ffmpeg as fallback for OGG/Opus decoding
        $ffmpegOk = $false
        try {
            $ffVer = & ffmpeg -version 2>&1
            if ($ffVer -match "ffmpeg") { $ffmpegOk = $true }
        } catch { }

        if (-not $ffmpegOk) {
            Write-Host "[ClaudePlus] Installation de ffmpeg via winget..." -ForegroundColor Yellow
            try {
                $wingetResult = & winget install Gyan.FFmpeg --accept-source-agreements --accept-package-agreements -q 2>&1
                # Refresh PATH after install
                $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
                try {
                    $ffVer = & ffmpeg -version 2>&1
                    if ($ffVer -match "ffmpeg") { $ffmpegOk = $true }
                } catch { }
            } catch { }

            if (-not $ffmpegOk) {
                # Try chocolatey
                try {
                    $chocoPath = (Get-Command choco -ErrorAction SilentlyContinue).Source
                    if ($chocoPath) {
                        Write-Host "[ClaudePlus] Essai via chocolatey..." -ForegroundColor Yellow
                        & choco install ffmpeg -y 2>&1 | Out-Null
                        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
                        try {
                            $ffVer = & ffmpeg -version 2>&1
                            if ($ffVer -match "ffmpeg") { $ffmpegOk = $true }
                        } catch { }
                    }
                } catch { }
            }
        }

        if ($ffmpegOk) {
            Write-Host "[ClaudePlus] ffmpeg OK (fallback PyAV)" -ForegroundColor DarkGray
        } else {
            Write-Host "[ClaudePlus] WARN: ni PyAV ni ffmpeg disponible. Audio OGG/Opus peut echouer." -ForegroundColor Yellow
            Write-Host "[ClaudePlus] Installez manuellement: winget install Gyan.FFmpeg" -ForegroundColor Yellow
        }
    }

    # Verify whisper availability
    $whisperOk = $false
    $checkFw = (& $py -c "import faster_whisper; print('ok')" 2>$null) | Select-Object -Last 1
    if ("$checkFw".Trim() -eq "ok") { $whisperOk = $true; Write-Host "[ClaudePlus] faster-whisper OK" -ForegroundColor DarkGray }
    else {
        $checkOw = (& $py -c "import whisper; print('ok')" 2>$null) | Select-Object -Last 1
        if ("$checkOw".Trim() -eq "ok") { $whisperOk = $true; Write-Host "[ClaudePlus] openai-whisper OK (fallback)" -ForegroundColor DarkGray }
    }
    if (-not $whisperOk) {
        Write-Host "[ClaudePlus] ERREUR: Aucun moteur Whisper installe. Vocal desactive." -ForegroundColor Red
        return $false
    }

    $script:PythonCmd = $py
    $script:TranscriptionReady = $true

    # Use the FiscalIQ transcription script (PyAV + VAD + GPU auto-detect + retry logic)
    $fiscalIqScript = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) "FiscalIQ\wwwroot\scripts\transcribe_audio.py"
    if (Test-Path $fiscalIqScript) {
        $script:TranscribeScript = $fiscalIqScript
        Write-Host "[ClaudePlus] Transcription: script FiscalIQ (PyAV + VAD + GPU)" -ForegroundColor Green
    } else {
        # Fallback: embedded script with same features
        $script:TranscribeScript = "$env:TEMP\claudeplus_transcribe.py"
        $pyCode = @'
import sys, json, os, wave, shutil, subprocess

def find_ffmpeg():
    found = shutil.which('ffmpeg')
    if found: return found
    for c in [r'C:\ffmpeg-bin\ffmpeg.exe', r'C:\ffmpeg\bin\ffmpeg.exe', r'C:\Program Files\ffmpeg\bin\ffmpeg.exe', r'C:\ProgramData\chocolatey\bin\ffmpeg.exe']:
        if os.path.isfile(c): return c
    return None

def convert_to_wav(input_path):
    ext = os.path.splitext(input_path)[1].lower()
    if ext in ('.wav', '.wave'): return input_path
    try:
        import av
        wav_path = input_path + ".converted.wav"
        container = av.open(input_path)
        resampler = av.AudioResampler(format='s16', layout='mono', rate=16000)
        raw = bytearray()
        for frame in container.decode(audio=0):
            for f in resampler.resample(frame):
                raw.extend(f.to_ndarray().tobytes())
        container.close()
        with wave.open(wav_path, 'w') as wf:
            wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(16000); wf.writeframes(bytes(raw))
        return wav_path
    except ImportError: pass
    except Exception as e: sys.stderr.write(f"PyAV: {e}\n")
    ffmpeg = find_ffmpeg()
    if ffmpeg:
        wav_path = input_path + ".ffmpeg.wav"
        try:
            r = subprocess.run([ffmpeg, '-y', '-i', input_path, '-ar', '16000', '-ac', '1', '-acodec', 'pcm_s16le', wav_path, '-loglevel', 'error'], capture_output=True, timeout=30)
            if r.returncode == 0 and os.path.exists(wav_path): return wav_path
        except: pass
    return input_path

audio_path = sys.argv[1]
language = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] != "auto" else None
wav_path = convert_to_wav(audio_path)
converted = wav_path != audio_path

try:
    try:
        from faster_whisper import WhisperModel
        device, compute = "cpu", "int8"
        try:
            import torch
            if torch.cuda.is_available(): device, compute = "cuda", "float16"
        except: pass
        model = WhisperModel("base", device=device, compute_type=compute)
        segments, info = model.transcribe(wav_path, language=language, beam_size=5, vad_filter=True)
        text = " ".join(s.text for s in segments).strip()
        if not text:
            segments, info = model.transcribe(wav_path, language=language, beam_size=5, vad_filter=False)
            text = " ".join(s.text for s in segments).strip()
        print(json.dumps({"text": text, "language": info.language, "provider": "faster-whisper"}, ensure_ascii=False))
        sys.exit(0)
    except ImportError: pass
    try:
        import whisper
        model = whisper.load_model("base")
        opts = {"language": language} if language else {}
        result = model.transcribe(wav_path, **opts)
        print(json.dumps({"text": result["text"].strip(), "language": result.get("language",""), "provider": "openai-whisper"}, ensure_ascii=False))
        sys.exit(0)
    except ImportError: pass
    print(json.dumps({"error": "Aucun moteur Whisper disponible"}))
finally:
    if converted and os.path.exists(wav_path):
        try: os.unlink(wav_path)
        except: pass
'@
        [System.IO.File]::WriteAllText($script:TranscribeScript, $pyCode, [System.Text.Encoding]::UTF8)
        Write-Host "[ClaudePlus] Transcription: script embarque (PyAV + VAD + GPU)" -ForegroundColor Green
    }

    Write-Host "[ClaudePlus] Transcription vocale prete (auto-detection langue, 99+ langues)" -ForegroundColor Green
    return $true
}

function Download-TelegramFile {
    param(
        [string]$FileId,
        [string]$Token,
        [string]$OutputPath
    )
    try {
        $fileInfo = Invoke-RestMethod -Uri "https://api.telegram.org/bot$Token/getFile?file_id=$FileId" -Method Get -TimeoutSec 10
        if (-not $fileInfo.ok) {
            Write-Host "[ClaudePlus] Erreur getFile Telegram" -ForegroundColor Red
            return $null
        }
        $filePath = $fileInfo.result.file_path
        $downloadUrl = "https://api.telegram.org/file/bot$Token/$filePath"

        # Determine output path from Telegram filename if not specified
        if (-not $OutputPath) {
            $ext = [System.IO.Path]::GetExtension($filePath)
            if (-not $ext) { $ext = ".bin" }
            $OutputPath = "$env:TEMP\claudeplus_file_$(Get-Random)$ext"
        }

        Write-Host "[ClaudePlus] Telechargement: $filePath" -ForegroundColor DarkGray
        Invoke-WebRequest -Uri $downloadUrl -OutFile $OutputPath -TimeoutSec 30 -ErrorAction Stop

        if (Test-Path $OutputPath) {
            $fileSize = [Math]::Round((Get-Item $OutputPath).Length / 1024, 1)
            Write-Host "[ClaudePlus] Fichier telecharge: $OutputPath ($fileSize KB)" -ForegroundColor DarkGray
            return $OutputPath
        }
        return $null
    } catch {
        Write-Host "[ClaudePlus] Erreur telechargement: $_" -ForegroundColor Red
        return $null
    }
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

        # Step 3: Transcribe with FiscalIQ-style script (PyAV + VAD + GPU + retry)
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $script:PythonCmd
        $psi.Arguments = "`"$($script:TranscribeScript)`" `"$audioFile`""
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

        # Parse JSON result (FiscalIQ format: text, language, provider)
        $result = $stdout.Trim() | ConvertFrom-Json
        if ($result.error) {
            Write-Host "[ClaudePlus] Transcription erreur: $($result.error)" -ForegroundColor Red
            return $null
        }
        # Add probability field for compatibility (FiscalIQ script doesn't include it)
        if (-not $result.probability) { $result | Add-Member -NotePropertyName "probability" -NotePropertyValue 0.95 -ErrorAction SilentlyContinue }
        $provider = if ($result.provider) { $result.provider } else { "whisper" }
        Write-Host "[ClaudePlus] Transcription OK ($provider): [$($result.language)] $($result.text)" -ForegroundColor Green
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
        [string]$TelegramChatId,
        [string]$SessionName,
        [string[]]$ImagePaths
    )

    # If images are attached, copy them to workdir and add paths to the message
    if ($ImagePaths -and $ImagePaths.Count -gt 0) {
        $fileRefs = @()
        foreach ($imgPath in $ImagePaths) {
            if (Test-Path $imgPath) {
                $destName = "telegram_$(Split-Path -Leaf $imgPath)"
                $destPath = Join-Path $WorkDir $destName
                Copy-Item -Path $imgPath -Destination $destPath -Force -ErrorAction SilentlyContinue
                if (Test-Path $destPath) {
                    $fileRefs += $destPath
                    Write-Host "[ClaudePlus] Image copiee: $destPath" -ForegroundColor DarkGray
                }
            }
        }
        if ($fileRefs.Count -gt 0) {
            $fileList = ($fileRefs | ForEach-Object { "`"$_`"" }) -join ", "
            $Message = "$Message`n`n[Image(s) jointe(s) — utilise l'outil Read pour les voir: $fileList]"
        }
    }

    $escapedMsg = $Message -replace '"', '\"'
    $argList = @("-p", "`"$escapedMsg`"", "--output-format", "stream-json", "--verbose")
    if ($Continue) { $argList += "--continue" }
    if ($DangerouslySkipPermissions) { $argList += "--dangerously-skip-permissions" }
    $argStr = $argList -join " "

    Write-Host "[ClaudePlus] Stream: claude -p --output-format stream-json --verbose $(if($Continue){'--continue '})" -ForegroundColor DarkCyan

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

        # Symbol map for tool types (BMP chars only — PS 5.x [char] is 16-bit, no emoji > 0xFFFF)
        $toolEmojis = @{
            "Read" = [char]0x25B6; "Write" = [char]0x270F; "Edit" = [char]0x2702
            "Bash" = [char]0x2699; "Grep" = [char]0x25C6; "Glob" = [char]0x25A0
            "Search" = [char]0x25C6; "Agent" = [char]0x25B7; "Explore" = [char]0x25CB
            "TodoWrite" = [char]0x25AA; "WebSearch" = [char]0x25C8; "WebFetch" = [char]0x25C8
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
            # Format G: Claude Code CLI stream-json {"type":"assistant","message":{"content":[{"type":"tool_use","name":"Bash","input":{...}}]}}
            if (-not $detectedTool -and $event.message -and $event.message.content) {
                foreach ($block in $event.message.content) {
                    if ($block.type -eq "tool_use" -and $block.name) {
                        $detectedTool = $block.name
                        # Also extract input from this block directly
                        if (-not $toolInput -and $block.input) { $toolInput = $block.input }
                        break
                    }
                }
            }

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
                # Format G: input from message.content[].input
                if (-not $toolInput -and $event.message -and $event.message.content) {
                    foreach ($block in $event.message.content) {
                        if ($block.type -eq "tool_use" -and $block.input) { $toolInput = $block.input; break }
                    }
                }

                if ($toolInput) {
                    $preview = $toolInput.command
                    if (-not $preview) { $preview = $toolInput.pattern }
                    if (-not $preview) { $preview = $toolInput.file_path }
                    if (-not $preview) { $preview = $toolInput.path }
                    if (-not $preview) { $preview = $toolInput.description }
                    if (-not $preview) { $preview = $toolInput.query }
                    if ($preview) {
                        if ($preview.Length -gt 120) { $preview = $preview.Substring(0, 117) + "..." }
                        $inputPreview = ": $preview"
                    }
                }

                $toolDisplay = "$emoji $detectedTool$inputPreview"
                $toolsUsed += $toolDisplay
                Write-Host "[ClaudePlus] Stream: Tool #$toolCount -> $detectedTool$inputPreview" -ForegroundColor Magenta

                # Throttled Telegram progress update (verbose mode only)
                $now = Get-Date
                $elapsed = [int]($now - $startTime).TotalSeconds
                if ($script:TelegramVerbose -and $TelegramToken -and $TelegramChatId) {
                    $secsSinceUpdate = ($now - $lastTelegramUpdate).TotalSeconds
                    if ($secsSinceUpdate -ge $telegramUpdateInterval -or $toolCount -eq 1) {
                        $tag = if ($SessionName) { "@${SessionName} : " } else { "" }
                        $progressMsg = "${tag}$([char]0x23F3) Claude travaille... (${elapsed}s)`n"
                        foreach ($t in $toolsUsed) { $progressMsg += "  $t`n" }
                        Send-TelegramMessage -Message $progressMsg.TrimEnd() -Token $TelegramToken -ChatId $TelegramChatId
                        $lastTelegramUpdate = $now
                    }
                }
            }

            # --- DETECT TOOL RESULT (capture output for Telegram) ---
            if ($event.type -eq "tool_result" -or $event.subtype -eq "tool_result") {
                $resultContent = $null
                if ($event.content -and $event.content -is [string]) { $resultContent = $event.content }
                elseif ($event.output -and $event.output -is [string]) { $resultContent = $event.output }
                elseif ($event.result -and $event.result -is [string] -and $event.type -ne "result") { $resultContent = $event.result }
                if ($resultContent -and $resultContent.Length -gt 0) {
                    $resultPreview = if ($resultContent.Length -gt 200) { $resultContent.Substring(0, 197) + "..." } else { $resultContent }
                    # Clean newlines for Telegram
                    $resultPreview = $resultPreview -replace "`r`n", "`n"
                    $toolsUsed += "  $([char]0x2192) $resultPreview"
                    Write-Host "[ClaudePlus] Stream: ToolResult ($($resultContent.Length) chars)" -ForegroundColor DarkMagenta
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
            # ONLY use result if we haven't captured text from assistant/message (avoids duplication)
            elseif ($event.type -eq "result" -and $event.result -and $textBuilder.Length -eq 0) {
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

        # Send final progress summary if tools were used (verbose mode only)
        if ($script:TelegramVerbose -and $TelegramToken -and $TelegramChatId -and $toolCount -gt 0) {
            $tag = if ($SessionName) { "@${SessionName} : " } else { "" }
            $summaryMsg = "${tag}$([char]0x2705) Termine (${elapsed}s, $toolCount outils)`n"
            foreach ($t in $toolsUsed) { $summaryMsg += "  $t`n" }
            Send-TelegramMessage -Message $summaryMsg.TrimEnd() -Token $TelegramToken -ChatId $TelegramChatId
        }

        # Save tools info for the caller to include in Telegram
        $script:LastPipeTools = $toolsUsed
        $script:LastPipeToolCount = $toolCount
        $script:LastPipeElapsed = $elapsed

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

# Plain pipe mode with detailed logging (no streaming, no Telegram updates)
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

    Write-Host "[ClaudePlus] PipePlain: $ClaudePath $argStr" -ForegroundColor DarkGray

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

        $startTime = Get-Date
        $proc = [System.Diagnostics.Process]::Start($psi)
        Write-Host "[ClaudePlus] PipePlain: process PID=$($proc.Id) lance" -ForegroundColor DarkGray

        $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
        $stderrTask = $proc.StandardError.ReadToEndAsync()

        # 5 minute timeout (claude can take long on complex tasks)
        $exited = $proc.WaitForExit(300000)
        $elapsed = [int]((Get-Date) - $startTime).TotalSeconds

        if (-not $exited) {
            Write-Host "[ClaudePlus] PipePlain: TIMEOUT 5min, kill" -ForegroundColor Red
            try { $proc.Kill() } catch { }
            return $null
        }

        $exitCode = $proc.ExitCode
        $stdout = $stdoutTask.Result
        $stderr = $stderrTask.Result

        Write-Host "[ClaudePlus] PipePlain: exit=$exitCode, ${elapsed}s, stdout=$($stdout.Length), stderr=$($stderr.Length)" -ForegroundColor DarkGray

        if ($stderr -and $stderr.Trim().Length -gt 0) {
            $stderrPreview = $stderr.Substring(0, [Math]::Min(200, $stderr.Length))
            Write-Host "[ClaudePlus] PipePlain stderr: $stderrPreview" -ForegroundColor DarkGray
        }

        $result = $stdout.Trim()
        if ($result.Length -gt 0) {
            Write-Host "[ClaudePlus] PipePlain: OK ($($result.Length) chars)" -ForegroundColor Green
            return $result
        }

        # stdout empty — check stderr for useful content
        $stderrClean = $stderr.Trim()
        if ($stderrClean.Length -gt 10 -and $stderrClean -notmatch "^Usage:|^Error:|^warn") {
            Write-Host "[ClaudePlus] PipePlain: utilise stderr comme reponse" -ForegroundColor Yellow
            return $stderrClean
        }

        Write-Host "[ClaudePlus] PipePlain: reponse VIDE (exit=$exitCode)" -ForegroundColor Red
        return $null
    } catch {
        Write-Host "[ClaudePlus] PipePlain erreur: $_" -ForegroundColor Red
        return $null
    }
}

# ============================================================================
# SEND TEXT TO CLAUDE WINDOW
# Windows Terminal: uses SendInput (simulated keyboard, works with GPU renderer)
# conhost.exe: uses PostMessage WM_CHAR (direct buffer write, no focus needed)
# ============================================================================

function Send-TextToClaude {
    param([string]$Text)

    Write-Host "[ClaudePlus] Envoi texte au terminal..." -ForegroundColor DarkGray

    try {
        # Escape SendKeys special characters: + ^ % ~ { } [ ] ( )
        $escaped = $Text -replace '([+^%~\{\}\[\]\(\)])', '{$1}'

        $helperScript = "$env:TEMP\claudeplus_sendkeys.ps1"
        $textFile = "$env:TEMP\claudeplus_sendtext.txt"
        $logFile = "$env:TEMP\claudeplus_sendkeys.log"
        $csFile = "$env:TEMP\claudeplus_winfocus.cs"
        [System.IO.File]::WriteAllText($textFile, $escaped, [System.Text.Encoding]::UTF8)
        if (Test-Path $logFile) { Remove-Item $logFile -ErrorAction SilentlyContinue }

        # Write C# code to a separate file (no here-string nesting, no C# 7 syntax)
        $csCode = @'
using System;
using System.Runtime.InteropServices;
public class WinFocus2 {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int cmd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    public static void ForceForeground(IntPtr targetHwnd) {
        IntPtr foreHwnd = GetForegroundWindow();
        uint dummy;
        uint foreThread = GetWindowThreadProcessId(foreHwnd, out dummy);
        uint curThread = GetCurrentThreadId();
        if (foreThread != curThread) {
            AttachThreadInput(curThread, foreThread, true);
            SetForegroundWindow(targetHwnd);
            BringWindowToTop(targetHwnd);
            AttachThreadInput(curThread, foreThread, false);
        } else {
            SetForegroundWindow(targetHwnd);
            BringWindowToTop(targetHwnd);
        }
    }
}
'@
        [System.IO.File]::WriteAllText($csFile, $csCode, [System.Text.Encoding]::UTF8)

        # Write the helper PowerShell script (no nested here-strings)
        $helperCode = @'
param([string]$TextFile, [string]$LogFile, [long]$Hwnd, [string]$CsFile)
try {
    Add-Type -AssemblyName System.Windows.Forms
    $csCode = [System.IO.File]::ReadAllText($CsFile, [System.Text.Encoding]::UTF8)
    Add-Type -TypeDefinition $csCode
    $text = [System.IO.File]::ReadAllText($TextFile, [System.Text.Encoding]::UTF8)
    $h = [IntPtr]::new($Hwnd)
    [WinFocus2]::ShowWindow($h, 9) | Out-Null
    Start-Sleep -Milliseconds 150
    [WinFocus2]::ForceForeground($h)
    Start-Sleep -Milliseconds 500
    $fgNow = [WinFocus2]::GetForegroundWindow()
    $gotFocus = ($fgNow -eq $h)
    [System.Windows.Forms.SendKeys]::SendWait($text)
    Start-Sleep -Milliseconds 300
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    [System.IO.File]::WriteAllText($LogFile, "OK:SendKeys:len=$($text.Length):hwnd=$Hwnd:focus=$gotFocus:fg=$fgNow")
} catch {
    [System.IO.File]::WriteAllText($LogFile, "EXCEPTION: $_")
}
'@
        [System.IO.File]::WriteAllText($helperScript, $helperCode, [System.Text.Encoding]::UTF8)

        $hwndLong = [long]$script:ClaudeWindowHandle
        Write-Host "[ClaudePlus] SendKeys vers hwnd=$hwndLong (AttachThread)..." -ForegroundColor DarkGray

        # Run helper minimized — Hidden blocks SetForegroundWindow, Normal flashes a black window
        # Minimized = window exists (so Win32 focus APIs work) but no visible flash
        $argStr = "-NoProfile -STA -ExecutionPolicy Bypass -File `"$helperScript`" -TextFile `"$textFile`" -LogFile `"$logFile`" -Hwnd $hwndLong -CsFile `"$csFile`""
        $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $argStr -PassThru -WindowStyle Minimized
        # Short timeout — don't block Telegram response
        $proc.WaitForExit(8000) | Out-Null
        if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }

        if (Test-Path $logFile) {
            $logContent = [System.IO.File]::ReadAllText($logFile)
            Write-Host "[ClaudePlus] Send result: $logContent" -ForegroundColor DarkGray
            Remove-Item $logFile -ErrorAction SilentlyContinue
            if ($logContent -match "^OK") { return $true }
        } else {
            Write-Host "[ClaudePlus] Send: pas de log (timeout?)" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "[ClaudePlus] Send erreur: $_" -ForegroundColor Yellow
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
    # Clean up dispatch file for this session
    $dispatchFile = Join-Path $script:SessionRegistryDir "dispatch_$($Name.ToLower()).json"
    if (Test-Path $dispatchFile) { Remove-Item $dispatchFile -ErrorAction SilentlyContinue }
    # If no more active sessions, clean up shared pending files
    $remaining = Get-ActiveSessions
    if ($remaining.Count -eq 0) {
        $pendingMsg = Join-Path $script:SessionRegistryDir "pending_message.json"
        $pendingChoice = Join-Path $script:SessionRegistryDir "pending_choice.json"
        if (Test-Path $pendingMsg) { Remove-Item $pendingMsg -ErrorAction SilentlyContinue }
        if (Test-Path $pendingChoice) { Remove-Item $pendingChoice -ErrorAction SilentlyContinue }
    }
}

function Get-ActiveSessions {
    if (-not (Test-Path $script:SessionRegistryDir)) { return @() }
    $sessions = @()
    Get-ChildItem -Path $script:SessionRegistryDir -Filter "*.json" -ErrorAction SilentlyContinue | ForEach-Object {
        # Skip non-session files (dispatch, pending, choice files)
        if ($_.Name -match '^(dispatch_|pending_)') { return }
        try {
            $data = Get-Content $_.FullName -Raw | ConvertFrom-Json
            # Check if process is still alive
            if (-not $data.Pid) { return }  # Not a session file
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
# Supports: @name : message, @name: message, @name:message, @name message
function Parse-MessageTarget {
    param([string]$Message)
    # Support: @nom : msg, @nom1,nom2 : msg, @all : msg, @nom msg, @nom
    # With commas: target may contain commas for multi-session (colon required)
    # Format: @target : message  (colon separator, target may contain commas/spaces)
    if ($Message -match '^@([a-zA-Z0-9_,\s-]+?)\s*:\s*(.+)$') {
        $rawTarget = $Matches[1].Trim().ToLower()
        return @{ Target = $rawTarget; Text = $Matches[2].Trim() }
    }
    # Format: @name message (space separator, no colon — single name only)
    if ($Message -match '^@([a-zA-Z0-9_-]+)\s+(.+)$') {
        return @{ Target = $Matches[1].ToLower(); Text = $Matches[2] }
    }
    # Format: @target (just the target, no message — may contain commas)
    if ($Message -match '^@([a-zA-Z0-9_,\s-]+?)\s*:?\s*$') {
        $rawTarget = $Matches[1].Trim().ToLower()
        return @{ Target = $rawTarget; Text = "" }
    }
    return @{ Target = $null; Text = $Message }
}

# Parse command parameters: /cmd name1, name2  OR  /cmd name  OR  /cmd (no param = all)
# Returns: @{ Command = "/stop"; Targets = @("fiscal","option1") }  or Targets = @() for all
function Parse-CommandTargets {
    param([string]$CommandText)
    # Match: /command  optionalparams
    if ($CommandText -match '^(/\w+)\s*(.*)$') {
        $cmd = $Matches[1].ToLower()
        $paramStr = $Matches[2].Trim()
        if (-not $paramStr) {
            return @{ Command = $cmd; Targets = @() }  # No param = apply to all
        }
        # Split by comma, trim spaces, lowercase
        $targets = @($paramStr -split ',' | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_.Length -gt 0 })
        return @{ Command = $cmd; Targets = $targets }
    }
    return @{ Command = $CommandText.ToLower(); Targets = @() }
}

# Check if this session is targeted by a command's parameters
# Supports exact match and partial match (e.g., "fis" matches "fiscaliq")
function Test-CommandTargetsMe {
    param([string[]]$Targets)
    $myName = $script:SessionName
    if (-not $myName) { return $false }
    $myNameLower = $myName.ToLower()

    # No targets specified = command applies to ALL sessions
    if (-not $Targets -or $Targets.Count -eq 0) { return $true }

    foreach ($t in $Targets) {
        # Exact match
        if ($t -eq $myNameLower) { return $true }
        # Partial match: target is prefix of my name (e.g., "fis" matches "fiscaliq")
        if ($myNameLower.StartsWith($t)) { return $true }
        # Partial match: my name is prefix of target (e.g., "fiscaliq" matches "fiscaliqpro")
        if ($t.StartsWith($myNameLower)) { return $true }
    }
    return $false
}

# Send numbered session choice list to Telegram
function Send-SessionChoiceList {
    param(
        [string]$Token,
        [string]$ChatId,
        [string]$MessagePreview
    )
    $sessions = Get-ActiveSessions | Sort-Object { $_.Name }
    $script:SessionChoiceMap = @{}

    $msg = "$([char]0x2753) Plusieurs sessions actives.`n"
    if ($MessagePreview) {
        $preview = if ($MessagePreview.Length -gt 50) { $MessagePreview.Substring(0, 47) + "..." } else { $MessagePreview }
        $msg += "Message : $preview`n"
    }
    $msg += "`nChoisissez la destination :`n"
    $i = 1
    foreach ($s in $sessions) {
        $dir = Split-Path -Leaf $s.WorkDir
        $msg += "`n  $i. @$($s.Name) ($dir)"
        $script:SessionChoiceMap[$i] = $s.Name
        $i++
    }
    $msg += "`n  0. Toutes les sessions"
    $msg += "`n`nTapez le numero (ex: 1 ou 1,2,3 ou 0)"
    $msg += "`nAnnulation auto dans 60s."

    Send-TelegramMessage -Message $msg -Token $Token -ChatId $ChatId

    # Write pending_choice.json with the session map so ANY session can resolve the choice
    $pendingFile = Join-Path $script:SessionRegistryDir "pending_choice.json"
    $choiceMapForJson = @{}
    foreach ($key in $script:SessionChoiceMap.Keys) {
        $choiceMapForJson["$key"] = $script:SessionChoiceMap[$key]
    }
    @{
        Timestamp = (Get-Date).ToString("o")
        SenderPid = $PID
        SessionChoiceMap = $choiceMapForJson
    } | ConvertTo-Json | Set-Content -Path $pendingFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

# Resolve user's number input to session names
# Input: "1" or "1,2,3" or "1, 2, 3" or "1 2 3" or "0"
# Returns: array of session names, or $null if invalid
function Resolve-SessionChoice {
    param([string]$UserInput)

    $numbers = @($UserInput -split '[,\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
    if ($numbers.Count -eq 0) { return $null }

    # 0 = all sessions
    if (0 -in $numbers) {
        $sessions = Get-ActiveSessions
        return @($sessions | ForEach-Object { $_.Name.ToLower() })
    }

    # Resolve each number
    $names = @()
    foreach ($n in $numbers) {
        if ($script:SessionChoiceMap.ContainsKey($n)) {
            $name = $script:SessionChoiceMap[$n]
            if ($name -and $name -notin $names) { $names += $name.ToLower() }
        } else {
            return $null  # Invalid number found
        }
    }
    if ($names.Count -eq 0) { return $null }
    return $names
}

# Write a dispatch file for another session to pick up
function ConvertTo-SafeJsonString {
    param([string]$Value)
    if (-not $Value) { return "" }
    $s = $Value -replace '\\', '\\' -replace '"', '\"'
    $s = $s -replace "`r`n", '\n' -replace "`n", '\n' -replace "`r", ''
    return $s
}

function ConvertTo-SafeJsonPathArray {
    # Returns a JSON array string like: ["c:\\path1","c:\\path2"] or []
    # Only includes paths that actually exist on disk
    param([object[]]$Paths)
    $valid = [System.Collections.Generic.List[string]]::new()
    if ($Paths) {
        foreach ($p in $Paths) {
            if ($p -eq $null) { continue }
            $ps = "$p"
            if ($ps.Length -gt 2 -and (Test-Path $ps -ErrorAction SilentlyContinue)) {
                $escaped = ConvertTo-SafeJsonString -Value $ps
                $valid.Add('"' + $escaped + '"')
            }
        }
    }
    if ($valid.Count -gt 0) { return '[' + ($valid -join ',') + ']' }
    return '[]'
}

function Write-DispatchFile {
    param(
        [string]$TargetSessionName,
        [string]$Text,
        [string[]]$ImagePaths,
        [string[]]$FilePaths
    )
    $dispatchFile = Join-Path $script:SessionRegistryDir "dispatch_$($TargetSessionName.ToLower()).json"
    $safeText = ConvertTo-SafeJsonString -Value $Text
    $imgArr = ConvertTo-SafeJsonPathArray -Paths $ImagePaths
    $fileArr = ConvertTo-SafeJsonPathArray -Paths $FilePaths
    $ts = (Get-Date).ToString("o")
    $json = '{"Text":"' + $safeText + '","ImagePaths":' + $imgArr + ',"FilePaths":' + $fileArr + ',"Timestamp":"' + $ts + '"}'
    Set-Content -Path $dispatchFile -Value $json -Encoding UTF8 -ErrorAction SilentlyContinue
}

# Check if this session should handle the message
function Test-MessageForMe {
    param([string]$RawText)

    # Commands: /stop, /list, /help, /verbose, /quiet
    if ($RawText -match '^/(stop|list|sessions|help|aide|verbose|quiet)') {
        $cmdParsed = Parse-CommandTargets -CommandText $RawText
        $activeSessions = Get-ActiveSessions

        # /list and /help are always handled by ONE session (lowest PID) to avoid duplicates
        if ($cmdParsed.Command -in @('/list', '/sessions', '/help', '/aide')) {
            $minPid = ($activeSessions | Sort-Object { [int]$_.Pid } | Select-Object -First 1).Pid
            if ($activeSessions.Count -le 1 -or [int]$minPid -eq $PID) {
                return @{ ShouldHandle = $true; Text = $RawText; IsCommand = $true; CommandTargets = $cmdParsed.Targets }
            }
            return @{ ShouldHandle = $false; Text = ""; IsCommand = $true; CommandTargets = @() }
        }

        # /stop, /verbose, /quiet: check if targets include me
        if ($activeSessions.Count -le 1) {
            return @{ ShouldHandle = $true; Text = $RawText; IsCommand = $true; CommandTargets = $cmdParsed.Targets }
        }
        $targetsMe = Test-CommandTargetsMe -Targets $cmdParsed.Targets
        return @{ ShouldHandle = $targetsMe; Text = $RawText; IsCommand = $true; CommandTargets = $cmdParsed.Targets }
    }

    # Regular message (not a command)
    $activeSessions = Get-ActiveSessions
    if ($activeSessions.Count -le 1) {
        # Single session — handle directly
        return @{ ShouldHandle = $true; Text = $RawText; IsCommand = $false; NeedsTarget = $false }
    }

    # Multi-session: needs numbered list selection
    # Any session that receives this message will handle it (no "default" — Telegram offset race)
    return @{ ShouldHandle = $false; Text = $RawText; IsCommand = $false; NeedsTarget = $true }
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
        $script:PendingFilePaths = @()
        $script:PendingImagePaths = @()

        # Clean up stale pending files from previous sessions
        $stalePending = Join-Path $script:SessionRegistryDir "pending_message.json"
        if (Test-Path $stalePending) { Remove-Item $stalePending -ErrorAction SilentlyContinue }
        $stalePendingChoice = Join-Path $script:SessionRegistryDir "pending_choice.json"
        if (Test-Path $stalePendingChoice) { Remove-Item $stalePendingChoice -ErrorAction SilentlyContinue }

        Delete-TelegramWebhook -Token $config.TelegramBotToken

        $workDir = (Get-Location).Path
        $allArgs = $claudeArgsList -join " "
        $sessionId = Get-Random -Minimum 10000 -Maximum 99999
        $batPath = "$env:TEMP\claudeplus_$sessionId.bat"
        $batLines = @(
            "@echo off",
            "title ClaudePlus @$Name",
            "cd /d `"$workDir`"",
            "cls",
            "`"$claudePath`" $allArgs",
            "pause"
        )
        $batLines -join "`r`n" | Set-Content $batPath -Encoding ASCII

        # Launch terminal for Claude TUI — always use Windows Terminal if available
        $wtPath = (Get-Command wt.exe -ErrorAction SilentlyContinue).Source

        if ($wtPath) {
            Write-Host "[ClaudePlus] Lancement via Windows Terminal (wt.exe)..." -ForegroundColor Cyan
            # --window new : force une NOUVELLE fenetre WT (pas un onglet) pour isolation multi-session
            $script:ClaudeProcess = Start-Process -FilePath $wtPath -ArgumentList "--window new --title `"ClaudePlus @$Name`" cmd.exe /c `"$batPath`"" -PassThru
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
            # Windows Terminal: wt.exe -> WindowsTerminal.exe -> OpenConsole.exe -> cmd.exe -> claude
            # STRATEGY: Find cmd.exe PID first (unique per session), then walk UP process tree
            # to find the WindowsTerminal.exe parent. This is 100% reliable even in multi-session
            # because each --window new creates a separate WT process.
            # NOTE: Title-based search does NOT work — Claude Code TUI overwrites the window title.

            # Step 1: Wait for cmd.exe to spawn and find its PID via our unique bat filename
            Write-Host "[ClaudePlus] Recherche cmd.exe via bat file unique..." -ForegroundColor DarkGray
            $batFileName = Split-Path -Leaf $batPath
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Milliseconds 500
                if ($script:ClaudeCmdPid -eq 0) {
                    try {
                        # Search by command line containing our unique bat filename
                        $cmdProcs = Get-CimInstance Win32_Process -Filter "Name='cmd.exe'" -ErrorAction SilentlyContinue
                        foreach ($cp in $cmdProcs) {
                            if ($cp.CommandLine -and $cp.CommandLine -match [regex]::Escape($batFileName)) {
                                $script:ClaudeCmdPid = [int]$cp.ProcessId
                                Write-Host "[ClaudePlus] cmd.exe SESSION TROUVE! PID=$($script:ClaudeCmdPid) bat=$batFileName" -ForegroundColor Green
                                break
                            }
                        }
                    } catch { }
                }
                if ($script:ClaudeCmdPid -gt 0) { break }
            }

            # Step 2: Walk UP the process tree from cmd.exe to find WindowsTerminal.exe
            if ($script:ClaudeCmdPid -gt 0) {
                Write-Host "[ClaudePlus] Remontee arbre process depuis cmd.exe PID=$($script:ClaudeCmdPid)..." -ForegroundColor DarkGray
                $walkPid = $script:ClaudeCmdPid
                for ($walk = 0; $walk -lt 10; $walk++) {
                    try {
                        $walkProc = Get-CimInstance Win32_Process -Filter "ProcessId=$walkPid" -ErrorAction SilentlyContinue
                        if (-not $walkProc) { break }
                        $parentPid = [int]$walkProc.ParentProcessId
                        if ($parentPid -le 0) { break }
                        $parentProc = Get-Process -Id $parentPid -ErrorAction SilentlyContinue
                        if ($parentProc -and $parentProc.ProcessName -eq "WindowsTerminal") {
                            if ($parentProc.MainWindowHandle -ne [IntPtr]::Zero) {
                                $script:ClaudeWindowHandle = $parentProc.MainWindowHandle
                                $found = $true
                                Write-Host "[ClaudePlus] FENETRE WT TROUVEE (arbre process)! Handle=$($parentProc.MainWindowHandle) PID=$parentPid" -ForegroundColor Green
                                break
                            }
                        }
                        $walkPid = $parentPid
                    } catch { break }
                }
            }

            # Step 3: If tree walk failed, try finding WT by our wt.exe launch PID
            if (-not $found -and $script:ClaudeProcess) {
                Write-Host "[ClaudePlus] Arbre process: pas trouve, essai via PID lancement..." -ForegroundColor Yellow
                for ($i = 0; $i -lt 10; $i++) {
                    Start-Sleep -Milliseconds 500
                    try {
                        # wt.exe (our process) may have spawned WindowsTerminal.exe as child
                        $wtChildren = Get-CimInstance Win32_Process -Filter "ParentProcessId=$($script:ClaudeProcess.Id)" -ErrorAction SilentlyContinue
                        foreach ($wc in $wtChildren) {
                            $proc = Get-Process -Id $wc.ProcessId -ErrorAction SilentlyContinue
                            if ($proc -and $proc.ProcessName -eq "WindowsTerminal" -and $proc.MainWindowHandle -ne [IntPtr]::Zero) {
                                $script:ClaudeWindowHandle = $proc.MainWindowHandle
                                $found = $true
                                Write-Host "[ClaudePlus] FENETRE WT TROUVEE (enfant wt.exe)! Handle=$($proc.MainWindowHandle)" -ForegroundColor Green
                                break
                            }
                        }
                    } catch { }
                    if ($found) { break }
                }
            }

            # Step 4: Mono-session fallback — accept unique WT window
            if (-not $found) {
                $activeSessions = Get-ActiveSessions
                $allWt = @(Get-Process -Name "WindowsTerminal" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
                if ($activeSessions.Count -le 1 -and $allWt.Count -eq 1) {
                    $script:ClaudeWindowHandle = $allWt[0].MainWindowHandle
                    $found = $true
                    Write-Host "[ClaudePlus] FENETRE WT trouvee (unique, mono-session)! Handle=$($allWt[0].MainWindowHandle)" -ForegroundColor Yellow
                }
            }

            # If cmd.exe PID still not found (Step 1 failed), try broader search
            if ($script:ClaudeCmdPid -eq 0) {
                Write-Host "[ClaudePlus] Recherche cmd.exe enfant (fallback)..." -ForegroundColor DarkGray
                try {
                    $allCmds = Get-CimInstance Win32_Process -Filter "Name='cmd.exe'" -ErrorAction SilentlyContinue
                    foreach ($cmd in $allCmds) {
                        if ($cmd.CommandLine -and $cmd.CommandLine -match "claude") {
                            $script:ClaudeCmdPid = [int]$cmd.ProcessId
                            Write-Host "[ClaudePlus] cmd.exe claude TROUVE (fallback)! PID=$($script:ClaudeCmdPid)" -ForegroundColor Green
                            break
                        }
                    }
                } catch { }
            }

            # Get console window handle for the cmd.exe (this is the conhost handle that accepts WM_CHAR)
            if ($script:ClaudeCmdPid -gt 0) {
                $cmdHwnd = Get-WindowHandleForPid -Pid1 $script:ClaudeCmdPid
                if ($cmdHwnd -and $cmdHwnd -ne [IntPtr]::Zero) {
                    Write-Host "[ClaudePlus] Console cmd.exe handle TROUVE: $cmdHwnd (pour PostMessage)" -ForegroundColor Green
                    # CRITICAL: If WT window handle wasn't found, use the cmd.exe console handle instead
                    # This is unique per session (each cmd.exe has its own console) — safe for multi-session
                    if ($script:ClaudeWindowHandle -eq [IntPtr]::Zero) {
                        $script:ClaudeWindowHandle = $cmdHwnd
                        Write-Host "[ClaudePlus] Utilisation du handle console cmd.exe comme handle fenetre (fallback)" -ForegroundColor Yellow
                    }
                } else {
                    Write-Host "[ClaudePlus] cmd.exe PID=$($script:ClaudeCmdPid) n'a pas de handle fenetre visible (WT l'heberge)" -ForegroundColor Yellow
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

        if ($script:ClaudeWindowHandle -eq [IntPtr]::Zero) {
            Write-Host "[ClaudePlus] ATTENTION: Pas de handle fenetre. Le mirror TUI peut echouer (pipe mode OK)." -ForegroundColor Red
        } else {
            Write-Host "[ClaudePlus] Handle fenetre final: $($script:ClaudeWindowHandle) (SendKeys actif)" -ForegroundColor Green
        }

        # Initialize voice transcription (Python + faster-whisper + ffmpeg)
        $voiceOk = Initialize-Transcription

        # Build startup message with session info
        $startupLines = @()
        $startupLines += "@${Name} : $([char]0x25BA) Session demarree!"
        $startupLines += "Dossier: $(Split-Path -Leaf $workDir)"
        if ($sessionCount -gt 1) {
            $otherNames = ($activeSessions | Where-Object { $_.Pid -ne $PID } | ForEach-Object { "@$($_.Name)" }) -join ", "
            $startupLines += "Autres sessions: $otherNames"
            $startupLines += ""
            $startupLines += "Repondez avec @${Name} : votre message"
        }
        if ($voiceOk) {
            $startupLines += "Texte et vocal OK"
        } else {
            $startupLines += "Texte OK (vocal desactive)"
        }
        $modeLabel = if ($script:TelegramVerbose) { "detaille" } else { "discret" }
        $startupLines += "Mode: $modeLabel (/verbose ou /quiet)"
        $startupLines += "/help | /list | /verbose | /quiet | /stop"

        # Wait for Claude TUI to start in the visible terminal BEFORE sending Telegram notification
        Write-Host "[ClaudePlus] Attente demarrage Claude TUI (3s)..." -ForegroundColor DarkGray
        Start-Sleep -Seconds 3

        Send-TelegramMessage -Message ($startupLines -join "`n") -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

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
                # Check if Claude is still running
                # With Windows Terminal, wt.exe exits immediately (delegates to existing WT instance)
                # So we check the cmd.exe child or the WT window instead
                if ($script:UsesWindowsTerminal) {
                    # For WT: check if cmd.exe child (running Claude) is still alive
                    # Don't check "any WT" — user may have other WT tabs open
                    $wtAlive = $false
                    if ($script:ClaudeCmdPid -gt 0) {
                        $cmdProc = Get-Process -Id $script:ClaudeCmdPid -ErrorAction SilentlyContinue
                        if ($cmdProc) { $wtAlive = $true }
                    }
                    if (-not $wtAlive) {
                        Write-Host "[ClaudePlus] Terminal Claude ferme (cmd.exe PID=$($script:ClaudeCmdPid) termine), arret." -ForegroundColor Yellow
                        break
                    }
                } else {
                    # For conhost: classic check
                    if ($script:ClaudeProcess.HasExited) { Write-Host "[ClaudePlus] Terminal ferme." -ForegroundColor Yellow; break }
                    if ($script:ClaudeCmdPid -gt 0) {
                        $cmdProc = Get-Process -Id $script:ClaudeCmdPid -ErrorAction SilentlyContinue
                        if (-not $cmdProc) { Write-Host "[ClaudePlus] cmd.exe termine, arret." -ForegroundColor Yellow; break }
                    }
                }

                # --- CHECK DISPATCH FILES (inter-session routing) ---
                $dispatchFile = Join-Path $script:SessionRegistryDir "dispatch_$($Name.ToLower()).json"
                if (Test-Path $dispatchFile) {
                    try {
                        # CRITICAL: Reset ALL file state before dispatch processing
                        $script:PendingFilePaths = @()
                        $script:PendingImagePaths = @()

                        $dispatchData = Get-Content $dispatchFile -Raw | ConvertFrom-Json
                        Remove-Item $dispatchFile -ErrorAction SilentlyContinue
                        Write-Host "[ClaudePlus] Dispatch recu: '$($dispatchData.Text)'" -ForegroundColor Green

                        # Extract ONLY the text — NO file handling from dispatch
                        # PS 5 ConvertFrom-Json turns [] into $null or phantom objects
                        # Solution: extract file paths with EXPLICIT string validation
                        $dispatchText = if ($dispatchData.Text) { [string]$dispatchData.Text } else { "" }
                        $safeImagePaths = [System.Collections.Generic.List[string]]::new()
                        $safeFilePaths = [System.Collections.Generic.List[string]]::new()

                        # Bulletproof extraction: iterate and validate each path is a real non-empty string
                        if ($dispatchData.ImagePaths -ne $null) {
                            foreach ($p in $dispatchData.ImagePaths) {
                                $ps = [string]$p
                                if ($ps -and $ps.Length -gt 2 -and (Test-Path $ps -ErrorAction SilentlyContinue)) {
                                    $safeImagePaths.Add($ps)
                                }
                            }
                        }
                        if ($dispatchData.FilePaths -ne $null) {
                            foreach ($p in $dispatchData.FilePaths) {
                                $ps = [string]$p
                                if ($ps -and $ps.Length -gt 2 -and (Test-Path $ps -ErrorAction SilentlyContinue)) {
                                    $safeFilePaths.Add($ps)
                                }
                            }
                        }

                        # Only append "Fichier joint:" if the file ACTUALLY EXISTS on disk
                        if ($safeFilePaths.Count -gt 0) {
                            $script:PendingFilePaths = @($safeFilePaths)
                            foreach ($f in $safeFilePaths) {
                                $dispatchText += "`nFichier joint: $f"
                            }
                        }
                        if ($safeImagePaths.Count -gt 0) {
                            $script:PendingImagePaths = @($safeImagePaths)
                        }

                        Write-Host "[ClaudePlus] [DEBUG-DISPATCH] dispatchText='$dispatchText' safeImages=$($safeImagePaths.Count) safeFiles=$($safeFilePaths.Count)" -ForegroundColor Magenta

                        if (-not $waitingForResponse -and $dispatchText) {
                            Write-Host "`n  >> [$Name] (dispatch) $dispatchText" -ForegroundColor Cyan
                            $waitingForResponse = $true

                            # Send text to TUI terminal for visual display
                            if ($script:ClaudeWindowHandle -ne [IntPtr]::Zero) {
                                $tText = $dispatchText
                                if ($tText.Length -gt 200) { $tText = $tText.Substring(0, 197) + "..." }
                                $tText = ($tText -split "`n" | Where-Object { $_ -notmatch '\\claudeplus_|Fichier joint:|\[Image' }) -join " "
                                if ($tText.Length -gt 0) { Send-TextToClaude -Text $tText | Out-Null }
                            }

                            $useSkip = ($config.DangerouslySkipPermissions -eq $true)
                            # FORCE fresh session for dispatch — never use --continue
                            $useContinue = $false
                            $script:PipeMessageCount = 0
                            $responseText = $null
                            $imgArgs = @()
                            if ($safeImagePaths.Count -gt 0) {
                                $imgArgs = @($safeImagePaths)
                            }

                            Write-Host "[ClaudePlus] Tentative 1/3: stream-json..." -ForegroundColor DarkGray
                            $responseText = Invoke-ClaudePipe -Message $dispatchText -ClaudePath $claudePath -WorkDir $workDir -Continue:$false -DangerouslySkipPermissions:$useSkip -TelegramToken $config.TelegramBotToken -TelegramChatId $config.TelegramChatId -SessionName $Name -ImagePaths $imgArgs

                            if (-not $responseText -or $responseText.Length -le 3) {
                                Write-Host "[ClaudePlus] Tentative 2/3: pipe plain..." -ForegroundColor Yellow
                                Start-Sleep -Seconds 2
                                $responseText = Invoke-ClaudePipePlain -Message $dispatchText -ClaudePath $claudePath -WorkDir $workDir -DangerouslySkipPermissions:$useSkip
                            }
                            if (-not $responseText -or $responseText.Length -le 3) {
                                Write-Host "[ClaudePlus] Tentative 3/3: pipe plain (session fraiche)..." -ForegroundColor Yellow
                                Start-Sleep -Seconds 3
                                $responseText = Invoke-ClaudePipePlain -Message $dispatchText -ClaudePath $claudePath -WorkDir $workDir -DangerouslySkipPermissions:$useSkip
                                if ($responseText -and $responseText.Length -gt 3) { $script:PipeMessageCount = 0 }
                            }

                            $script:PipeMessageCount++
                            $waitingForResponse = $false

                            if ($responseText -and $responseText.Length -gt 3) {
                                $previewLen = [Math]::Min(200, $responseText.Length)
                                Write-Host "  << $($responseText.Substring(0, $previewLen))$(if($responseText.Length -gt 200){'...'})" -ForegroundColor Green
                                $toolSection = ""
                                if ($script:TelegramVerbose -and $script:LastPipeTools -and $script:LastPipeToolCount -gt 0) {
                                    $toolSection = "$([char]0x2699) Outils ($($script:LastPipeToolCount)):`n"
                                    foreach ($t in $script:LastPipeTools) { $toolSection += "$t`n" }
                                    $toolSection += "`n"
                                }
                                $maxTextLen = 3900 - $toolSection.Length
                                if ($maxTextLen -lt 500) { $maxTextLen = 500; $toolSection = "" }
                                if ($responseText.Length -gt $maxTextLen) { $responseText = $responseText.Substring(0, $maxTextLen) + "`n[... tronque]" }
                                Send-TelegramMessage -Message "@${Name} : ${toolSection}${responseText}" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            } else {
                                Send-TelegramMessage -Message "@${Name} : Echec apres 3 tentatives." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            }

                            # Cleanup ONLY real files
                            foreach ($tmpFile in @($safeImagePaths) + @($safeFilePaths)) {
                                if ($tmpFile -and $tmpFile.Length -gt 2 -and (Test-Path $tmpFile)) { Remove-Item $tmpFile -ErrorAction SilentlyContinue }
                            }
                            Get-ChildItem -Path $workDir -Filter "telegram_claudeplus_*" -ErrorAction SilentlyContinue | Remove-Item -ErrorAction SilentlyContinue
                            $script:PendingFilePaths = @()
                            $script:PendingImagePaths = @()
                        }
                    } catch {
                        Write-Host "[ClaudePlus] Erreur dispatch: $_" -ForegroundColor Red
                    }
                }

                # --- CHECK PENDING CHOICE TIMEOUT (shared file) ---
                $pendingMsgFile = Join-Path $script:SessionRegistryDir "pending_message.json"
                if (Test-Path $pendingMsgFile) {
                    try {
                        $pendingCheck = Get-Content $pendingMsgFile -Raw | ConvertFrom-Json
                        $elapsed = (Get-Date) - [datetime]$pendingCheck.Timestamp
                        if ($elapsed.TotalSeconds -gt 60) {
                            Remove-Item $pendingMsgFile -ErrorAction SilentlyContinue
                            $pendingChoiceCleanup = Join-Path $script:SessionRegistryDir "pending_choice.json"
                            Remove-Item $pendingChoiceCleanup -ErrorAction SilentlyContinue
                            $script:WaitingForSessionChoice = $false
                            $script:SessionChoiceMap = @{}
                            # Only one session sends the cancellation message (the one that stored it)
                            if ($pendingCheck.SenderPid -eq $PID) {
                                Send-TelegramMessage -Message "$([char]0x2716) Demande annulee (delai de 60s depasse)." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            }
                            Write-Host "[ClaudePlus] Timeout choix session — demande annulee." -ForegroundColor Yellow
                        }
                    } catch {
                        # Corrupted file, clean up
                        Remove-Item $pendingMsgFile -ErrorAction SilentlyContinue
                    }
                }

                # --- CHECK TELEGRAM ---
                try {
                    # Use short polling (timeout=0) to avoid 409 conflict when multiple sessions share the same bot
                    $activeSessions = Get-ActiveSessions
                    $multiSession = ($activeSessions.Count -gt 1)
                    $pollTimeout = if ($multiSession) { 0 } else { 2 }
                    $url = "https://api.telegram.org/bot$($config.TelegramBotToken)/getUpdates?limit=10&timeout=$pollTimeout"
                    if ($lastUpdateId -gt 0) { $url += "&offset=$($lastUpdateId + 1)" }
                    if ($pollCount -le 5 -or $pollCount % 10 -eq 1) {
                        Write-Host "[ClaudePlus] Telegram poll #$pollCount (offset=$lastUpdateId, hwnd=$($script:ClaudeWindowHandle)$(if($multiSession){', multi'}))" -ForegroundColor DarkGray
                    }
                    $response = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 5 -ErrorAction Stop

                    if ($pollCount -le 5) {
                        $msgCount = if ($response.result) { $response.result.Count } else { 0 }
                        Write-Host "[ClaudePlus] Poll #${pollCount} ok=$($response.ok) messages=$msgCount" -ForegroundColor DarkGray
                    }

                    if ($response.ok -and $response.result) {
                        foreach ($update in $response.result) {
                            if ($update.update_id -gt $lastUpdateId) { $lastUpdateId = $update.update_id }
                            $msg = $update.message
                            if (-not $msg) { continue }
                            if ($msg.chat.id -ne [long]$config.TelegramChatId) { continue }

                            # Determine message type: text or voice
                            $text = $null
                            $isVoice = $false
                            # CRITICAL: Reset file/image paths at each iteration to prevent leakage from previous messages
                            $script:PendingFilePaths = @()
                            $script:PendingImagePaths = @()

                            if ($msg.voice -or $msg.audio) {
                                # Voice message or audio file
                                # In multi-session, use caption for routing: send vocal with caption "@fiscaliq"
                                # Caption is checked BEFORE transcription to avoid transcribing for wrong session
                                $isVoice = $true
                                $voiceCaption = if (-not [string]::IsNullOrEmpty($msg.caption)) { $msg.caption.Trim() } else { $null }

                                # Multi-session: only default session handles voice
                                $activeSessions = Get-ActiveSessions
                                if ($activeSessions.Count -gt 1) {
                                    $minPid = ($activeSessions | Sort-Object { [int]$_.Pid } | Select-Object -First 1).Pid
                                    if ([int]$minPid -ne $PID) {
                                        Write-Host "[ClaudePlus] Vocal en multi-session — pas le default, skip." -ForegroundColor DarkGray
                                        continue
                                    }
                                }

                                $fileId = if ($msg.voice) { $msg.voice.file_id } else { $msg.audio.file_id }
                                $duration = if ($msg.voice) { $msg.voice.duration } else { $msg.audio.duration }

                                if (-not $script:TranscriptionReady) {
                                    Send-TelegramMessage -Message "@${Name} : Vocal non supporte (Python/faster-whisper manquant)" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                    continue
                                }

                                Write-Host ""
                                Write-Host "  >> [Vocal ${duration}s] Transcription..." -ForegroundColor DarkCyan
                                Send-TelegramMessage -Message "@${Name} : Transcription en cours..." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

                                $transcription = Transcribe-TelegramAudio -FileId $fileId -Token $config.TelegramBotToken
                                if (-not $transcription -or [string]::IsNullOrEmpty($transcription.text)) {
                                    Send-TelegramMessage -Message "@${Name} : Erreur transcription (audio non reconnu)" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                    continue
                                }

                                # Build final text: transcription + optional extra instruction from caption
                                $transcribedText = $transcription.text.Trim()
                                $lang = $transcription.language
                                $conf = [int]($transcription.probability * 100)
                                Write-Host "  >> [Vocal -> $lang ${conf}%] $transcribedText" -ForegroundColor Cyan
                                Send-TelegramMessage -Message "@${Name} : Transcription ($lang ${conf}%): $transcribedText" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

                                # If caption had extra text after @name (e.g. "@fiscal : en anglais"), prepend it
                                if ($voiceCaption) {
                                    $captionParsed = Parse-MessageTarget -Message $voiceCaption
                                    $extraInstruction = $captionParsed.Text
                                    if ($extraInstruction -and $extraInstruction.Length -gt 0) {
                                        $text = "$extraInstruction : $transcribedText"
                                    } else {
                                        $text = $transcribedText
                                    }
                                } else {
                                    $text = $transcribedText
                                }

                                # Multi-session: store transcription as pending in shared file, ask user to choose session
                                if ($activeSessions.Count -gt 1) {
                                    $voicePendingFile = Join-Path $script:SessionRegistryDir "pending_message.json"
                                    $safeVT = ConvertTo-SafeJsonString -Value $text
                                    $ts = (Get-Date).ToString("o")
                                    $voiceJson = '{"Text":"' + $safeVT + '","Type":"text","ImagePaths":[],"FilePaths":[],"Timestamp":"' + $ts + '","SenderPid":' + $PID + '}'
                                    Set-Content -Path $voicePendingFile -Value $voiceJson -Encoding UTF8 -ErrorAction SilentlyContinue
                                    $script:WaitingForSessionChoice = $true
                                    Send-SessionChoiceList -Token $config.TelegramBotToken -ChatId $config.TelegramChatId -MessagePreview $text
                                    continue
                                }

                            } elseif ($msg.photo -or $msg.document) {
                                # Photo or document/file from Telegram
                                # Multi-session: only default session handles files
                                $activeSessions = Get-ActiveSessions
                                if ($activeSessions.Count -gt 1) {
                                    $minPid = ($activeSessions | Sort-Object { [int]$_.Pid } | Select-Object -First 1).Pid
                                    if ([int]$minPid -ne $PID) {
                                        Write-Host "[ClaudePlus] Fichier en multi-session — pas le default, skip." -ForegroundColor DarkGray
                                        continue
                                    }
                                }

                                $attachedFiles = @()
                                $isImage = $false

                                if ($msg.photo) {
                                    # Photo: array of sizes, take the largest (last)
                                    $photoSizes = @($msg.photo)
                                    $bestPhoto = $photoSizes[-1]
                                    $fileId = $bestPhoto.file_id
                                    $isImage = $true
                                    Write-Host "[ClaudePlus] Photo recue ($($bestPhoto.width)x$($bestPhoto.height))" -ForegroundColor DarkCyan
                                    $localPath = Download-TelegramFile -FileId $fileId -Token $config.TelegramBotToken -OutputPath "$env:TEMP\claudeplus_photo_$(Get-Random).jpg"
                                    if ($localPath) { $attachedFiles += $localPath }
                                }
                                elseif ($msg.document) {
                                    $fileId = $msg.document.file_id
                                    $fileName = $msg.document.file_name
                                    $mimeType = $msg.document.mime_type
                                    $fileSize = $msg.document.file_size
                                    Write-Host "[ClaudePlus] Document recu: $fileName ($mimeType, $([Math]::Round($fileSize/1024,1)) KB)" -ForegroundColor DarkCyan

                                    # Check file size limit (Telegram API: 20MB max download)
                                    if ($fileSize -gt 20 * 1024 * 1024) {
                                        Send-TelegramMessage -Message "@${Name} : Fichier trop volumineux (max 20MB)" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                        continue
                                    }

                                    $ext = [System.IO.Path]::GetExtension($fileName)
                                    $localPath = Download-TelegramFile -FileId $fileId -Token $config.TelegramBotToken -OutputPath "$env:TEMP\claudeplus_doc_$(Get-Random)$ext"
                                    if ($localPath) {
                                        $attachedFiles += $localPath
                                        # Check if this is an image
                                        if ($mimeType -match "^image/") { $isImage = $true }
                                    }
                                }

                                if ($attachedFiles.Count -eq 0) {
                                    Send-TelegramMessage -Message "@${Name} : Erreur telechargement du fichier" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                    continue
                                }

                                # Caption = text accompanying the file, or default prompt
                                $text = if (-not [string]::IsNullOrEmpty($msg.caption)) { $msg.caption.Trim() } else { "Analyse ce fichier" }

                                # Determine which files are images (for --image flag) vs other files (mention in prompt)
                                $imageExts = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg", ".tiff", ".tif")
                                $script:PendingImagePaths = @()
                                $script:PendingFilePaths = @()
                                foreach ($f in $attachedFiles) {
                                    $fExt = [System.IO.Path]::GetExtension($f).ToLower()
                                    if ($isImage -or $fExt -in $imageExts) {
                                        $script:PendingImagePaths += $f
                                    } else {
                                        $script:PendingFilePaths += $f
                                        # For non-image files, add path to the prompt so Claude can read them
                                        $text += "`nFichier joint: $f"
                                    }
                                }

                                $fileNames = ($attachedFiles | ForEach-Object { Split-Path -Leaf $_ }) -join ", "
                                Write-Host "  >> [$Name] $text [fichier: $fileNames]" -ForegroundColor Cyan

                                # Multi-session: store file info as pending in shared file, ask user to choose session
                                if ($activeSessions.Count -gt 1) {
                                    $filePendingFile = Join-Path $script:SessionRegistryDir "pending_message.json"
                                    $safeFT = ConvertTo-SafeJsonString -Value $text
                                    $imgArr = ConvertTo-SafeJsonPathArray -Paths $script:PendingImagePaths
                                    $fileArr = ConvertTo-SafeJsonPathArray -Paths $script:PendingFilePaths
                                    $ts = (Get-Date).ToString("o")
                                    $fileJson = '{"Text":"' + $safeFT + '","Type":"file","ImagePaths":' + $imgArr + ',"FilePaths":' + $fileArr + ',"Timestamp":"' + $ts + '","SenderPid":' + $PID + '}'
                                    Set-Content -Path $filePendingFile -Value $fileJson -Encoding UTF8 -ErrorAction SilentlyContinue
                                    $script:WaitingForSessionChoice = $true
                                    Send-SessionChoiceList -Token $config.TelegramBotToken -ChatId $config.TelegramChatId -MessagePreview $text
                                    continue
                                }

                            } elseif (-not [string]::IsNullOrEmpty($msg.text)) {
                                $text = $msg.text.Trim()
                            } else {
                                # Unknown message type, log and skip
                                Write-Host "[ClaudePlus] Message ignore (type non supporte)" -ForegroundColor DarkGray
                                continue
                            }

                            if (-not $text -or $text.Length -eq 0) { continue }

                            Write-Host "[ClaudePlus] Message recu: '$text' (chat=$($msg.chat.id))" -ForegroundColor Cyan

                            # --- SESSION CHOICE INTERCEPTION ---
                            # Check if a pending message exists (shared file) and user sent a number
                            # ANY session can handle this — no "default" concept (Telegram offset race)
                            $choiceResolved = $false
                            $pendingMsgFile = Join-Path $script:SessionRegistryDir "pending_message.json"
                            $pendingChoiceFile = Join-Path $script:SessionRegistryDir "pending_choice.json"
                            if ($text -match '^\s*[\d,\s]+\s*$' -and (Test-Path $pendingMsgFile)) {
                                try {
                                    # ATOMIC CLAIM: rename file to claim ownership
                                    # First session to rename wins, others find file gone
                                    $claimedFile = Join-Path $script:SessionRegistryDir "pending_message_claimed_$PID.json"
                                    try {
                                        [System.IO.File]::Move(
                                            (Join-Path $script:SessionRegistryDir "pending_message.json"),
                                            $claimedFile
                                        )
                                    } catch {
                                        # Another session already claimed it — skip
                                        Write-Host "[ClaudePlus] Pending deja reclame par une autre session, skip." -ForegroundColor DarkGray
                                        continue
                                    }
                                    # Also claim choice file
                                    $claimedChoiceFile = Join-Path $script:SessionRegistryDir "pending_choice_claimed_$PID.json"
                                    if (Test-Path $pendingChoiceFile) {
                                        try { [System.IO.File]::Move($pendingChoiceFile, $claimedChoiceFile) } catch {}
                                    }

                                    # Load the claimed pending message and choice map
                                    $pending = Get-Content $claimedFile -Raw | ConvertFrom-Json
                                    $choiceData = $null
                                    if (Test-Path $claimedChoiceFile) {
                                        $choiceData = Get-Content $claimedChoiceFile -Raw | ConvertFrom-Json
                                    }
                                    # Clean up claimed files immediately
                                    Remove-Item $claimedFile -ErrorAction SilentlyContinue
                                    Remove-Item $claimedChoiceFile -ErrorAction SilentlyContinue

                                    # Rebuild SessionChoiceMap from shared file
                                    if ($choiceData -and $choiceData.SessionChoiceMap) {
                                        $script:SessionChoiceMap = @{}
                                        $choiceData.SessionChoiceMap.PSObject.Properties | ForEach-Object {
                                            $script:SessionChoiceMap[[int]$_.Name] = $_.Value
                                        }
                                    }

                                    $chosenNames = Resolve-SessionChoice -UserInput $text
                                    if ($chosenNames) {
                                        $script:WaitingForSessionChoice = $false
                                        $script:SessionChoiceMap = @{}

                                        $myNameLower = $Name.ToLower()
                                        $targetLabel = "@" + ($chosenNames -join ", @")
                                        Send-TelegramMessage -Message "$([char]0x2192) Message envoye a $targetLabel" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId

                                        # Bulletproof extraction of file paths from pending JSON
                                        # PS 5 ConvertFrom-Json turns [] into phantom objects — validate each path
                                        $pendingSafeImages = [System.Collections.Generic.List[string]]::new()
                                        $pendingSafeFiles = [System.Collections.Generic.List[string]]::new()
                                        if ($pending.ImagePaths -ne $null) {
                                            foreach ($p in $pending.ImagePaths) {
                                                $ps = [string]$p
                                                if ($ps -and $ps.Length -gt 2 -and (Test-Path $ps -ErrorAction SilentlyContinue)) {
                                                    $pendingSafeImages.Add($ps)
                                                }
                                            }
                                        }
                                        if ($pending.FilePaths -ne $null) {
                                            foreach ($p in $pending.FilePaths) {
                                                $ps = [string]$p
                                                if ($ps -and $ps.Length -gt 2 -and (Test-Path $ps -ErrorAction SilentlyContinue)) {
                                                    $pendingSafeFiles.Add($ps)
                                                }
                                            }
                                        }

                                        # Dispatch to target sessions via files
                                        $iAmTarget = $false
                                        foreach ($targetName in $chosenNames) {
                                            if ($targetName -eq $myNameLower) {
                                                $iAmTarget = $true
                                            } else {
                                                Write-DispatchFile -TargetSessionName $targetName -Text $pending.Text -ImagePaths @($pendingSafeImages) -FilePaths @($pendingSafeFiles)
                                                Write-Host "[ClaudePlus] Dispatch ecrit pour @$targetName" -ForegroundColor DarkCyan
                                            }
                                        }

                                        if ($iAmTarget) {
                                            # I'm one of the targets — process the pending message directly
                                            $text = $pending.Text
                                            # Use pre-validated safe arrays — NEVER trust raw ConvertFrom-Json
                                            $script:PendingImagePaths = @($pendingSafeImages)
                                            $script:PendingFilePaths = @($pendingSafeFiles)
                                            $choiceResolved = $true
                                            # CRITICAL: Delete any dispatch file for myself that another session may have written
                                            # (prevents double-processing: once via iAmTarget, once via dispatch)
                                            $myDispatchFile = Join-Path $script:SessionRegistryDir "dispatch_$myNameLower.json"
                                            if (Test-Path $myDispatchFile) {
                                                Remove-Item $myDispatchFile -ErrorAction SilentlyContinue
                                                Write-Host "[ClaudePlus] Dispatch file pour moi supprime (evite double traitement)" -ForegroundColor DarkYellow
                                            }
                                            Write-Host "[ClaudePlus] Je suis la cible — text='$text', images=$($pendingSafeImages.Count), files=$($pendingSafeFiles.Count)" -ForegroundColor Green
                                            # Fall through to send-to-Claude below (skip routing)
                                        } else {
                                            # Not my target, dispatched to others only
                                            continue
                                        }
                                    } else {
                                        # Invalid number — show error
                                        $maxNum = 0
                                        if ($script:SessionChoiceMap.Count -gt 0) {
                                            $maxNum = ($script:SessionChoiceMap.Keys | Measure-Object -Maximum).Maximum
                                        }
                                        Send-TelegramMessage -Message "$([char]0x2757) Numero invalide. Tapez un numero entre 0 et $maxNum (ou plusieurs: 1,2,3)." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                        continue
                                    }
                                } catch {
                                    Write-Host "[ClaudePlus] Erreur lecture pending: $_" -ForegroundColor Red
                                    # Clean up all possible files (original + claimed)
                                    Remove-Item $pendingMsgFile -ErrorAction SilentlyContinue
                                    Remove-Item $pendingChoiceFile -ErrorAction SilentlyContinue
                                    $clF = Join-Path $script:SessionRegistryDir "pending_message_claimed_$PID.json"
                                    $clC = Join-Path $script:SessionRegistryDir "pending_choice_claimed_$PID.json"
                                    Remove-Item $clF -ErrorAction SilentlyContinue
                                    Remove-Item $clC -ErrorAction SilentlyContinue
                                    $script:WaitingForSessionChoice = $false
                                }
                            }

                            # --- MULTI-SESSION ROUTING (skip if choice was just resolved) ---
                            if ($choiceResolved) {
                                Write-Host "[ClaudePlus] Choix resolu, skip routing direct vers Claude." -ForegroundColor Green
                                # FORCE fresh pipe session — never use --continue for resolved choices
                                # This prevents stale context (e.g. previous "fichier joint" being remembered by Claude)
                                $script:PipeMessageCount = 0
                            } else {
                            $routing = Test-MessageForMe -RawText $text
                            Write-Host "[ClaudePlus] Routing: ShouldHandle=$($routing.ShouldHandle), IsCmd=$($routing.IsCommand), NeedsTarget=$($routing.NeedsTarget)" -ForegroundColor DarkGray

                            # If command but not targeted at me, skip silently
                            if ($routing.IsCommand -and -not $routing.ShouldHandle) {
                                Write-Host "[ClaudePlus] Commande non ciblee pour moi, skip." -ForegroundColor DarkGray
                                continue
                            }

                            # Handle /list command (show active sessions)
                            if ($routing.IsCommand -and $routing.Text -match '^/list|^/sessions') {
                                $sessions = Get-ActiveSessions
                                if ($sessions.Count -eq 0) {
                                    $listMsg = "@${Name} : Aucune session active"
                                } else {
                                    $listMsg = "@${Name} : $([char]0x25A0) Sessions actives ($($sessions.Count)):`n"
                                    foreach ($s in $sessions) {
                                        $me = if ($s.Pid -eq $PID) { " $([char]0x2190) ici" } else { "" }
                                        $listMsg += "  @$($s.Name) — $(Split-Path -Leaf $s.WorkDir)$me`n"
                                    }
                                }
                                Send-TelegramMessage -Message $listMsg.TrimEnd() -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }

                            if ($routing.IsCommand -and $routing.Text -match '^/stop') {
                                $cmdInfo = Parse-CommandTargets -CommandText $routing.Text
                                $targetLabel = if ($cmdInfo.Targets.Count -gt 0) { " (cibles: $($cmdInfo.Targets -join ', '))" } else { "" }
                                Send-TelegramMessage -Message "@${Name} : Session arretee.$targetLabel" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                $stopRequested = $true
                                break
                            }

                            # Help command — detailed, well-formatted help
                            if ($routing.IsCommand -and $routing.Text -match '^/help|^/aide') {
                                $activeSessions = Get-ActiveSessions
                                $isMulti = ($activeSessions.Count -gt 1)
                                $sep = "$([char]0x2500)" * 30
                                $bullet = "$([char]0x2022)"
                                $arrow = "$([char]0x2192)"
                                $gear = "$([char]0x2699)"
                                $info = "$([char]0x2139)"
                                $excl = "$([char]0x2757)"
                                $quest = "$([char]0x2753)"
                                $modeIcon = if ($script:TelegramVerbose) { "$([char]0x266A)" } else { "$([char]0x266B)" }
                                $modeNow = if ($script:TelegramVerbose) { "$modeIcon Detaille" } else { "$modeIcon Discret" }

                                $h = ""
                                $h += "$gear ClaudePlus $([char]0x2014) Guide complet`n$sep"

                                # --- SECTION: Envoyer des messages ---
                                $h += "`n`n$([char]0x270F) ENVOYER DES MESSAGES`n$sep"
                                if ($isMulti) {
                                    $h += "`n"
                                    $h += "`n$bullet Ecrivez votre message directement."
                                    $h += "`n$bullet Le bot affiche la liste des sessions."
                                    $h += "`n$bullet Tapez le numero pour choisir :"
                                    $h += "`n   1 $arrow une session"
                                    $h += "`n   1,2,3 $arrow plusieurs sessions"
                                    $h += "`n   0 $arrow toutes les sessions"
                                } else {
                                    $h += "`n"
                                    $h += "`n$bullet Ecrivez votre message directement,"
                                    $h += "`n   Claude le recoit et repond."
                                }

                                # --- SECTION: Vocaux ---
                                $h += "`n`n$([char]0x266A) MESSAGES VOCAUX`n$sep"
                                $h += "`n"
                                $h += "`n$bullet Envoyez un message vocal."
                                $h += "`n$bullet Il sera transcrit automatiquement"
                                $h += "`n   puis envoye a Claude."
                                if ($isMulti) {
                                    $h += "`n"
                                    $h += "`n$bullet Apres transcription, le bot vous"
                                    $h += "`n   demande de choisir la session."
                                }

                                # --- SECTION: Images and Fichiers ---
                                $h += "`n`n$([char]0x25A0) IMAGES ET FICHIERS`n$sep"
                                $h += "`n"
                                $h += "`n$bullet Images : jpg, png, gif, webp..."
                                $h += "`n$bullet Fichiers : pdf, docx, xlsx, code..."
                                $h += "`n$bullet Claude analyse le contenu et repond."
                                if ($isMulti) {
                                    $h += "`n"
                                    $h += "`n$bullet Apres reception, le bot vous"
                                    $h += "`n   demande de choisir la session."
                                }

                                # --- SECTION: Commandes ---
                                $h += "`n`n$gear COMMANDES`n$sep"
                                $h += "`n"
                                $h += "`n$quest /help ou /aide"
                                $h += "`n   Affiche ce guide complet."
                                $h += "`n"
                                $h += "`n$([char]0x25A0) /list"
                                $h += "`n   Affiche toutes les sessions actives"
                                $h += "`n   avec leur nom et dossier de travail."
                                $h += "`n"
                                $h += "`n$([char]0x25B6) /verbose"
                                $h += "`n   Active le mode detaille :"
                                $h += "`n   vous recevez les outils utilises par"
                                $h += "`n   Claude (Bash, Read, Grep...), la"
                                $h += "`n   progression en temps reel, puis le"
                                $h += "`n   resultat final."
                                if ($isMulti) {
                                    $h += "`n   $bullet /verbose $arrow toutes les sessions"
                                    $h += "`n   $bullet /verbose nom $arrow une session"
                                    $h += "`n   $bullet /verbose nom1, nom2 $arrow plusieurs"
                                }
                                $h += "`n"
                                $h += "`n$([char]0x25C0) /quiet"
                                $h += "`n   Active le mode discret :"
                                $h += "`n   vous recevez uniquement le resultat"
                                $h += "`n   final de Claude, sans le detail des"
                                $h += "`n   outils et commandes executees."
                                if ($isMulti) {
                                    $h += "`n   $bullet /quiet $arrow toutes les sessions"
                                    $h += "`n   $bullet /quiet nom $arrow une session"
                                    $h += "`n   $bullet /quiet nom1, nom2 $arrow plusieurs"
                                }
                                $h += "`n"
                                $h += "`n$([char]0x2716) /stop"
                                if ($isMulti) {
                                    $h += "`n   Arrete les sessions Claude et ferme"
                                    $h += "`n   les terminaux associes."
                                    $h += "`n   $bullet /stop $arrow toutes les sessions"
                                    $h += "`n   $bullet /stop nom $arrow une session"
                                    $h += "`n   $bullet /stop nom1, nom2 $arrow plusieurs"
                                } else {
                                    $h += "`n   Arrete la session Claude et ferme"
                                    $h += "`n   le terminal associe."
                                }

                                # --- SECTION: Infos session ---
                                $h += "`n`n$info INFORMATIONS`n$sep"
                                $h += "`n"
                                $h += "`n$bullet Mode actuel : $modeNow"
                                $h += "`n$bullet Session : @${Name}"
                                if ($isMulti) {
                                    $h += "`n$bullet Sessions actives ($($activeSessions.Count)) :"
                                    foreach ($s in $activeSessions) {
                                        $me = if ($s.Pid -eq $PID) { " $([char]0x2190) ici" } else { "" }
                                        $h += "`n   $([char]0x2502) @$($s.Name) $([char]0x2014) $(Split-Path -Leaf $s.WorkDir)$me"
                                    }
                                }

                                # --- SECTION: Astuces ---
                                $h += "`n`n$([char]0x2605) ASTUCES`n$sep"
                                $h += "`n"
                                $h += "`n$bullet Types acceptes : texte, vocal,"
                                $h += "`n   images, PDF, Word, Excel, code..."
                                $h += "`n$bullet Claude peut lire, modifier et"
                                $h += "`n   creer des fichiers dans le projet."

                                Send-TelegramMessage -Message $h -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }

                            # Toggle verbose/quiet mode (with session targeting in multi-session)
                            if ($routing.IsCommand -and $routing.Text -match '^/verbose') {
                                $script:TelegramVerbose = $true
                                Send-TelegramMessage -Message "@${Name} : $([char]0x2699) Mode detaille ON" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }
                            if ($routing.IsCommand -and $routing.Text -match '^/quiet') {
                                $script:TelegramVerbose = $false
                                Send-TelegramMessage -Message "@${Name} : $([char]0x2709) Mode discret ON" -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }

                            # Multi-session without target: store as pending and show choice list
                            # Any session that receives this handles it (no "default" — Telegram offset race)
                            if (-not $routing.ShouldHandle -and $routing.NeedsTarget) {
                                # Check if another session already stored a pending for the same message (avoid duplicates)
                                $existingPending = Join-Path $script:SessionRegistryDir "pending_message.json"
                                if (-not (Test-Path $existingPending)) {
                                    # Store pending message in SHARED file — use manual JSON to avoid PS 5 empty array issues
                                    $safePT = ConvertTo-SafeJsonString -Value $text
                                    $ts = (Get-Date).ToString("o")
                                    $pendingJson = '{"Text":"' + $safePT + '","Type":"text","ImagePaths":[],"FilePaths":[],"Timestamp":"' + $ts + '","SenderPid":' + $PID + '}'
                                    Set-Content -Path $existingPending -Value $pendingJson -Encoding UTF8 -ErrorAction SilentlyContinue
                                    $script:WaitingForSessionChoice = $true
                                    Send-SessionChoiceList -Token $config.TelegramBotToken -ChatId $config.TelegramChatId -MessagePreview $text
                                    Write-Host "[ClaudePlus] Liste de choix envoyee (pending stocke)." -ForegroundColor Green
                                } else {
                                    Write-Host "[ClaudePlus] Pending existe deja (autre session a gere)." -ForegroundColor DarkGray
                                }
                                continue
                            }

                            if (-not $routing.ShouldHandle) { continue }

                            } # end of: } else { (not choiceResolved — routing section)

                            # Skip if already waiting for a response
                            if ($waitingForResponse) {
                                Send-TelegramMessage -Message "@${Name} : ATTENTE — Claude est en train de repondre, patientez..." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                                continue
                            }

                            if (-not $isVoice) {
                                Write-Host ""
                                Write-Host "  >> [$Name] $text" -ForegroundColor Cyan
                            }

                            $waitingForResponse = $true

                            # HYBRID: send to TUI terminal for visual display
                            # Each session has its own WT window with unique title "ClaudePlus @Name"
                            # so the window handle is correct even in multi-session
                            if ($script:ClaudeWindowHandle -ne [IntPtr]::Zero) {
                                $terminalText = $text
                                if ($terminalText.Length -gt 200) {
                                    $terminalText = $terminalText.Substring(0, 197) + "..."
                                }
                                $terminalText = ($terminalText -split "`n" | Where-Object { $_ -notmatch '\\claudeplus_|Fichier joint:|\[Image' }) -join " "
                                if ($terminalText.Length -gt 0) {
                                    Send-TextToClaude -Text $terminalText | Out-Null
                                }
                            }

                            # PIPE WITH RETRY ESCALATION:
                            # Attempt 1: stream-json (real-time updates)
                            # Attempt 2: plain pipe with --continue
                            # Attempt 3: plain pipe WITHOUT --continue (fresh session)
                            $useSkip = ($config.DangerouslySkipPermissions -eq $true)
                            $responseText = $null
                            $useContinue = ($script:PipeMessageCount -gt 0)

                            # Attempt 1: stream-json
                            Write-Host "[ClaudePlus] [DEBUG-MAIN] Texte exact envoye au pipe: '$text'" -ForegroundColor Magenta
                            Write-Host "[ClaudePlus] [DEBUG-MAIN] choiceResolved=$choiceResolved, useContinue=$useContinue, PipeMessageCount=$($script:PipeMessageCount), PendingImages=$($script:PendingImagePaths.Count), PendingFiles=$($script:PendingFilePaths.Count)" -ForegroundColor Magenta
                            Write-Host "[ClaudePlus] Tentative 1/3: stream-json..." -ForegroundColor DarkGray
                            # Gather attached files (images + documents)
                            $imgArgs = @()
                            if ($script:PendingImagePaths -and $script:PendingImagePaths.Count -gt 0) {
                                $imgArgs = $script:PendingImagePaths
                                $script:PendingImagePaths = @()
                            }

                            $responseText = Invoke-ClaudePipe `
                                -Message $text `
                                -ClaudePath $claudePath `
                                -WorkDir $workDir `
                                -Continue:$useContinue `
                                -DangerouslySkipPermissions:$useSkip `
                                -TelegramToken $config.TelegramBotToken `
                                -TelegramChatId $config.TelegramChatId `
                                -SessionName $Name `
                                -ImagePaths $imgArgs

                            # Attempt 2: plain pipe with --continue
                            if (-not $responseText -or $responseText.Length -le 3) {
                                Write-Host "[ClaudePlus] Tentative 2/3: pipe plain + continue..." -ForegroundColor Yellow
                                Start-Sleep -Seconds 2
                                $responseText = Invoke-ClaudePipePlain `
                                    -Message $text `
                                    -ClaudePath $claudePath `
                                    -WorkDir $workDir `
                                    -Continue:$useContinue `
                                    -DangerouslySkipPermissions:$useSkip
                            }

                            # Attempt 3: plain pipe WITHOUT --continue (fresh session)
                            if (-not $responseText -or $responseText.Length -le 3) {
                                Write-Host "[ClaudePlus] Tentative 3/3: pipe plain SANS continue (session fraiche)..." -ForegroundColor Yellow
                                Start-Sleep -Seconds 3
                                $responseText = Invoke-ClaudePipePlain `
                                    -Message $text `
                                    -ClaudePath $claudePath `
                                    -WorkDir $workDir `
                                    -DangerouslySkipPermissions:$useSkip
                                # Reset continue counter since we started fresh
                                if ($responseText -and $responseText.Length -gt 3) {
                                    $script:PipeMessageCount = 0
                                }
                            }

                            $script:PipeMessageCount++
                            $waitingForResponse = $false

                            if ($responseText -and $responseText.Length -gt 3) {
                                # Show response in PS console
                                $previewLen = [Math]::Min(200, $responseText.Length)
                                Write-Host "  << $($responseText.Substring(0, $previewLen))$(if($responseText.Length -gt 200){'...'})" -ForegroundColor Green
                                Write-Host ""

                                # Build Telegram message — with or without tool details
                                $toolSection = ""
                                if ($script:TelegramVerbose -and $script:LastPipeTools -and $script:LastPipeToolCount -gt 0) {
                                    $toolSection = "$([char]0x2699) Outils ($($script:LastPipeToolCount)):`n"
                                    foreach ($t in $script:LastPipeTools) { $toolSection += "$t`n" }
                                    $toolSection += "`n"
                                }

                                # Truncate text if too long (4096 char limit minus tool section)
                                $maxTextLen = 3900 - $toolSection.Length
                                if ($maxTextLen -lt 500) { $maxTextLen = 500; $toolSection = "" }
                                if ($responseText.Length -gt $maxTextLen) {
                                    $responseText = $responseText.Substring(0, $maxTextLen) + "`n[... tronque]"
                                }
                                # Send final response with @name : prefix + tools (if verbose)
                                $finalMsg = "@${Name} : ${toolSection}${responseText}"
                                Send-TelegramMessage -Message $finalMsg -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            } else {
                                Write-Host "  << [ECHEC 3 tentatives]" -ForegroundColor Red
                                Write-Host ""
                                Send-TelegramMessage -Message "@${Name} : Echec apres 3 tentatives. Verifiez le terminal ou renvoyez le message." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
                            }

                            # Cleanup temp files (images + documents from TEMP)
                            foreach ($tmpFile in ($imgArgs + $script:PendingFilePaths)) {
                                if ($tmpFile -and (Test-Path $tmpFile)) {
                                    Remove-Item $tmpFile -ErrorAction SilentlyContinue
                                }
                            }
                            # Cleanup images copied to workdir
                            Get-ChildItem -Path $workDir -Filter "telegram_claudeplus_*" -ErrorAction SilentlyContinue | Remove-Item -ErrorAction SilentlyContinue
                            $script:PendingFilePaths = @()
                        }
                    }
                }
                catch {
                    if ($pollCount -le 5) {
                        Write-Host "[ClaudePlus] Poll #$pollCount ERREUR: $_" -ForegroundColor Red
                    }
                }

                # Jitter sleep in multi-session to desync polls and reduce 409 conflicts
                if ($multiSession) {
                    $jitter = Get-Random -Minimum 300 -Maximum 1200
                    Start-Sleep -Milliseconds $jitter
                } else {
                    Start-Sleep -Milliseconds 500
                }
            }
        }
        finally {
            Unregister-Session -Name $Name
            Send-TelegramMessage -Message "@${Name} : Session terminee." -Token $config.TelegramBotToken -ChatId $config.TelegramChatId
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
