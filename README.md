# ClaudePlus

<div align="center">

![ClaudePlus Logo](logo.svg)

**Manage multiple Claude Code terminals from Telegram — code anywhere, anytime**

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)](https://microsoft.com/powershell)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows%2010%2F11-blue?logo=windows)](https://microsoft.com/windows)
[![Version](https://img.shields.io/badge/Version-3.0-orange)](https://github.com/hayefmajid/ClaudePlus)

</div>

---

## Overview

**ClaudePlus** is a PowerShell module that turns Claude Code into a multi-session, Telegram-controlled AI assistant. Launch multiple Claude Code instances in separate Windows Terminal windows, each working on a different project, and control them all from your phone via a single Telegram bot.

Send text, voice, images, or files from Telegram. ClaudePlus routes your message to the right session, pipes it to Claude Code, and streams the response back to Telegram in real-time.

---

## Key Features

### Multi-Session Management
- **Multiple Claude Code terminals** running simultaneously, each in its own Windows Terminal window
- **Numbered session picker** — when multiple sessions are active, a numbered list appears:
  ```
  ? Plusieurs sessions actives.
  Message: Salut

  Choisissez la destination:
    1. @fiscaliq (FiscalIQ)
    2. @myproject (MyProject)
    0. Toutes les sessions

  Tapez le numero (ex: 1 ou 1,2,3 ou 0)
  ```
- **Single, multiple, or broadcast**: type `1`, `1,2,3`, `1 2 3`, or `0` (all sessions)
- **60-second auto-cancellation** for pending choices
- **Atomic routing** — only one session processes each message (no duplicates)

### Telegram Bidirectional Mirror
- **Text messages** — send from Telegram, see in terminal, get response back
- **Voice messages** — auto-transcribed via faster-whisper (99+ languages, GPU-accelerated)
- **Images & files** — send photos, PDFs, Word docs, Excel files; Claude processes them
- **Real-time streaming** — tool progress updates appear on Telegram as Claude works
- **Verbose mode** — `/verbose` shows which tools Claude used (Read, Edit, Bash, etc.)

### Hybrid Architecture
Each message follows two paths simultaneously:
1. **TUI Path** — Text appears in the visible Claude Code terminal via Win32 `SendKeys`
2. **Pipe Path** — `claude -p` returns clean UTF-8 response for Telegram

The pipe response goes to Telegram. The TUI display is visual feedback only.

### Smart Retry Logic
3-attempt escalation for maximum reliability:
1. **stream-json** — real-time progress with tool updates
2. **plain pipe + continue** — fallback with conversation context
3. **plain pipe fresh** — clean session if all else fails

---

## Architecture

```
                    ClaudePlus v3.0 — Multi-Session Architecture

  Telegram App                    ClaudePlus Module (PowerShell)
  (Phone/Web)                     One bot, multiple sessions
       |                                    |
       |  "Salut"                    +--------------+
       |  =========================> | Telegram Poll |
       |                             |  (getUpdates) |
       |                             +--------------+
       |                                    |
       |                          Multiple sessions active?
       |                           /                \
       |                         NO                  YES
       |                          |                   |
       |                     Direct to            Numbered picker
       |                     only session         sent to Telegram
       |                          |                   |
       |     "1"                  |              User picks "1"
       |  ========================|==================>|
       |                          |                   |
       |                     +----v-------------------v----+
       |                     |     Atomic Claim (File.Move) |
       |                     |  First session to claim wins |
       |                     +----+------------------------+
       |                          |
       |              +-----------v-----------+
       |              |   Dispatch to target  |
       |              |   session via file    |
       |              +-----------+-----------+
       |                          |
       |                +---------v----------+
       |                |  Target Session    |
       |                |                    |
       |                |  1. SendKeys->TUI  |
       |                |  2. claude -p      |
       |                |     (stream-json)  |
       |                +---------+----------+
       |                          |
       |     @fiscaliq: Response  |
       |  <=======================/
       |

  Windows Terminal           Windows Terminal
  ┌──────────────┐          ┌──────────────┐
  | Claude Code  |          | Claude Code  |
  | @fiscaliq    |          | @myproject   |
  | FiscalIQ/    |          | MyProject/   |
  | Handle: A    |          | Handle: B    |
  └──────────────┘          └──────────────┘
     Unique window             Unique window
     handle per session        handle per session
```

### Process-Tree Window Detection

Each session finds its own unique Windows Terminal window handle:

1. **Find cmd.exe** by unique bat filename (`claudeplus_XXXXX.bat`)
2. **Walk UP** the process tree: `cmd.exe` -> `OpenConsole` -> `WindowsTerminal.exe`
3. Each `--window new` creates a separate WT process with its own `MainWindowHandle`
4. **Result**: SendKeys always targets the correct terminal, even with 5+ sessions

### Inter-Session Communication

Sessions communicate via JSON files in `%LOCALAPPDATA%\ClaudePlus\sessions\`:

| File | Purpose |
|------|---------|
| `<name>.json` | Session registration (PID, workdir, name) |
| `pending_message.json` | Message awaiting session choice |
| `pending_choice.json` | Session choice map (number -> name) |
| `dispatch_<name>.json` | Message routed to a specific session |

**Atomic claim**: `File.Move` ensures only one session processes each choice (no race conditions).

---

## Prerequisites

### Required
- **Windows 10/11** (64-bit)
- **PowerShell 5.1+** (built into Windows)
- **Claude Code** installed (`claude.exe` in PATH)
- **Windows Terminal** (recommended for multi-session, auto-detected)
- **Telegram Bot Token** (from @BotFather)
- **Telegram Chat ID** (from @userinfobot)

### Optional (Auto-Installed)
- **Python 3.x** — for voice transcription
- **faster-whisper** — speech-to-text engine (auto-installed via pip)
- **PyAV** — native OGG/Opus audio conversion (no ffmpeg needed)
- **ffmpeg** — fallback audio codec (auto-installed via winget)

---

## Installation

### 1. Create a Telegram Bot

1. Open Telegram, search **@BotFather**, send `/newbot`
2. Choose a name and username (must end with `bot`)
3. Copy the **Bot Token**

### 2. Get Your Chat ID

1. Search **@userinfobot** on Telegram, send any message
2. Copy the **Chat ID** number

### 3. Install ClaudePlus

```powershell
# Quick install
.\Install-ClaudePlus.ps1

# Or manual
Import-Module .\ClaudePlus.psm1
claudeplus-config -TelegramBotToken "YOUR_TOKEN" -TelegramChatId "YOUR_ID"
```

### 4. Launch

```powershell
# Single session
Import-Module .\ClaudePlus.psm1
claudeplus

# Named session (for multi-session)
claudeplus -Name "fiscaliq"
```

### 5. Multi-Session

Open multiple PowerShell terminals, each in a different project folder:

```powershell
# Terminal 1: FiscalIQ project
cd "C:\Projects\FiscalIQ"
Import-Module C:\path\to\ClaudePlus.psm1
claudeplus -Name "fiscaliq"

# Terminal 2: MyWebApp project
cd "C:\Projects\MyWebApp"
Import-Module C:\path\to\ClaudePlus.psm1
claudeplus -Name "webapp"
```

All sessions share the same Telegram bot. Messages are routed via the numbered picker.

---

## Usage

### Text Messages
Send a text message to your Telegram bot. In single-session mode, it goes directly. In multi-session, you pick which session receives it.

### Voice Messages
Record a voice message in Telegram. ClaudePlus transcribes it with faster-whisper (auto language detection, 99+ languages, GPU-accelerated if available).

### Images & Files
Send photos, PDFs, Word docs, Excel files. Claude Code receives the file path and can read/process it.

### Telegram Commands

All commands support **session targeting** in multi-session mode. Without a target, the command applies to **all sessions**.

| Command | Description |
|---------|-------------|
| `/help` or `/aide` | Show complete interactive guide with all commands, session info, and usage tips |
| `/list` or `/sessions` | Show all active sessions with name, working directory, and PID |
| `/verbose` | Enable detailed mode — shows tool usage (Bash, Read, Grep, Edit...) and real-time progress |
| `/verbose name` | Enable verbose for a specific session (partial match supported: `fis` → `fiscaliq`) |
| `/verbose n1, n2` | Enable verbose for multiple sessions |
| `/quiet` | Enable quiet mode — shows only the final response, no tool details |
| `/quiet name` | Enable quiet for a specific session |
| `/quiet n1, n2` | Enable quiet for multiple sessions |
| `/stop` | Stop all sessions and close their terminals |
| `/stop name` | Stop a specific session |
| `/stop n1, n2` | Stop multiple sessions |

**Targeting examples:**
```
/stop fiscaliq          → stops the fiscaliq session
/stop fis               → partial match, also stops fiscaliq
/verbose fis, web       → enables verbose on fiscaliq and webapp
/quiet                  → enables quiet on ALL sessions
```

### PowerShell Commands

```powershell
# Launch single session (default name = folder name)
claudeplus

# Launch named session (for multi-session)
claudeplus -Name "fiscaliq"

# Launch without Telegram mirror
claudeplus -NoTelegram

# View current config
claudeplus-config

# Set Telegram credentials
claudeplus-config -TelegramBotToken "YOUR_TOKEN" -TelegramChatId "YOUR_ID"

# Toggle dangerous permissions (skip Claude Code confirmations)
claudeplus-config -DangerouslySkipPermissions $true

# Force reimport after code changes
Import-Module .\ClaudePlus.psm1 -Force
```

---

## Technical Details

### Stream-JSON Mode

ClaudePlus uses Claude Code's `--output-format stream-json` for real-time progress:

```
claude -p "message" --output-format stream-json --verbose
```

Each line is a JSON event: `init`, `assistant` (text chunks), `tool_use`, `tool_result`, `result`. Tool progress is sent to Telegram in real-time so you can see Claude reading files, running commands, etc.

### Pipe Retry Escalation

| Attempt | Method | Details |
|---------|--------|---------|
| 1 | `stream-json` | Real-time with `--verbose`, tool progress updates |
| 2 | `plain pipe` | Simple stdout capture, with `--continue` |
| 3 | `plain pipe fresh` | No `--continue`, clean session |

### Voice Transcription Pipeline

```
  Telegram Voice Message
        |
        v
  Download OGG file
        |
        v
  PyAV decode (native OGG/Opus → WAV 16kHz mono)
  [fallback: ffmpeg if PyAV unavailable]
        |
        v
  faster-whisper (CTranslate2)
  [fallback: openai-whisper]
        |
        v
  VAD filter (Voice Activity Detection)
  Removes silence, improves accuracy
  [retry without VAD if empty result]
        |
        v
  JSON output: {text, language, provider}
        |
        v
  Confirmation message on Telegram
  "[lang] Transcription: text..."
        |
        v
  Process as normal text message
```

**Transcription Engine** — 3-tier fallback:

| Priority | Engine | Details |
|----------|--------|---------|
| 1 | FiscalIQ script | Shared script at `wwwroot/scripts/transcribe_audio.py` (PyAV + VAD + GPU) |
| 2 | Embedded script | Standalone Python script with same features, auto-generated at `%TEMP%` |
| 3 | openai-whisper | Fallback if faster-whisper unavailable |

**Audio Conversion** — 2-tier fallback:

| Priority | Method | Details |
|----------|--------|---------|
| 1 | **PyAV** (native) | Pure Python, no external dependencies, OGG/Opus → WAV 16kHz mono |
| 2 | **ffmpeg** (fallback) | Auto-installed via winget if PyAV unavailable |

**Key Features:**
- **Model**: `base` (~150MB, auto-downloaded on first use)
- **GPU**: CUDA auto-detection — uses `float16` on GPU, `int8` on CPU
- **Languages**: 99+ with automatic detection (no configuration needed)
- **VAD**: Voice Activity Detection filters silence for better accuracy; automatic retry without VAD if result is empty
- **Auto-install**: faster-whisper, PyAV, and ffmpeg are installed automatically on first voice message
- **Confirmation**: After transcription, the recognized text and detected language are sent back to Telegram before processing

### Win32 Window Management

- **Process-tree detection**: `cmd.exe PID` → walk parents → `WindowsTerminal.exe`
- **SendKeys**: Compiled C# helper with `SetForegroundWindow` + `AttachThreadInput`
- **PostMessage WM_CHAR**: Direct character posting without focus stealing
- **Per-session isolation**: Each session has a unique bat file, cmd.exe PID, and WT window handle

### PowerShell 5 Compatibility

Special handling for PS 5.1 quirks:
- `ConvertFrom-Json` turns `[]` into phantom objects — all JSON uses manual serialization
- `@($null)` creates a 1-element array — explicit `Test-Path` validation on all file paths
- Emoji fallback for PS versions that crash on Unicode emojis

---

## Configuration Sources

ClaudePlus looks for Telegram credentials in this order:

1. **ClaudePlus config** — `%LOCALAPPDATA%\ClaudePlus\config.json`
2. **Project config** — `telegram-config.json` in working directory
3. **VS Extension config** — `%LOCALAPPDATA%\ClaudeCodeExtension\claudecode-settings.json`

Credentials found in source 2 or 3 are auto-imported into source 1.

---

## Troubleshooting

### Window handle not found
If ClaudePlus reports "Pas de handle fenetre", the process-tree walk couldn't find a `WindowsTerminal.exe` parent. The console cmd.exe handle is used as fallback. Pipe mode works regardless.

### Double responses in multi-session
This was fixed in v3.0 with atomic `File.Move` claim. If you see duplicates, ensure all sessions are running the latest code (`Import-Module -Force`).

### Voice transcription fails
- Ensure Python 3.x is in PATH
- Run `python -m pip install faster-whisper PyAV`
- PyAV handles OGG natively; ffmpeg is only needed as fallback

### Module not reloading
Always use `-Force`:
```powershell
Import-Module .\ClaudePlus.psm1 -Force
```

---

## Version History

### v3.0 — Multi-Session Numbered Picker
- Replace `@name` prefix routing with numbered session picker
- Atomic `File.Move` claim prevents duplicate processing
- Process-tree window detection (no more title-based search)
- Fix PS 5 `ConvertFrom-Json` ghost "Fichier joint" bug
- Manual JSON serialization helpers (`ConvertTo-SafeJsonString`, `ConvertTo-SafeJsonPathArray`)
- Fresh pipe sessions for dispatched messages (no stale `--continue` context)
- Debug logging for pipe text tracing

### v2.7 — Stream-JSON Fix
- Fix `--verbose` flag for stream-json mode
- Fix Windows Terminal text input via SendInput

### v2.6 — PS 5.x Stability
- Fix emoji crash on PowerShell 5.x
- Fix Windows Terminal exit detection

### v2.5 — Retry Logic
- 3-attempt escalation: stream-json -> plain+continue -> plain fresh
- Timeout handling per attempt

### v2.4 — Windows Terminal
- Switch from conhost to Windows Terminal (`wt.exe --window new`)
- Separate window per session for multi-session isolation

### v2.3 — Multi-Session
- Session registration and discovery via JSON files
- `@prefix` routing for targeting specific sessions
- `/list` command to show active sessions

### v2.2 — Stream-JSON Parser
- Real-time tool progress updates on Telegram
- Multi-format stream-json parser with debug logging

### v2.1 — Real-Time Streaming
- Tool progress (Read, Edit, Bash) shown on Telegram as Claude works

### v2.0 — Pipe Mode
- Switch from console buffer reading to `claude -p` pipe mode
- Hybrid architecture (TUI display + pipe response)
- Voice transcription via faster-whisper

### v1.0 — Initial Release
- Console buffer reading with `AttachConsole` + `ReadConsoleOutputCharacter`
- Basic Telegram mirror

---

## Project Structure

```
ClaudePlus/
├── ClaudePlus.psm1          # Main module (~2900 lines)
├── Install-ClaudePlus.ps1   # Quick install script
├── README.md                # This file
├── LICENSE                  # MIT License
├── logo.svg                 # Project logo
```

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Author

**Majid Hayef** — [GitHub](https://github.com/hayefmajid)

Built as part of the [FiscalIQ](https://github.com/hayefmajid) project ecosystem.
