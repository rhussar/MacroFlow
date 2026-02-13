const { spawn } = require('node:child_process');

class ExcelForegroundHook {
  constructor(onStateChange) {
    this.onStateChange = onStateChange;
    this.process = null;
    this.stdoutBuffer = '';
    this.lastState = null;
    this.stopped = false;
    this.restartTimer = null;
  }

  start() {
    if (process.platform !== 'win32') {
      return;
    }
    if (this.process) {
      return;
    }

    this.stopped = false;
    const child = spawn(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-EncodedCommand',
        Buffer.from(this.buildScript(), 'utf16le').toString('base64')
      ],
      { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] }
    );

    this.process = child;

    child.stdout.setEncoding('utf8');
    child.stdout.on('data', (chunk) => this.handleStdout(chunk));

    child.stderr.setEncoding('utf8');
    child.stderr.on('data', () => {
      // Ignore PowerShell host noise (CLIXML) on stderr.
    });

    child.on('exit', (code) => {
      this.process = null;
      if (this.stopped) {
        return;
      }
      console.error(`[ExcelHook] Hook process exited with code ${code}. Restarting...`);
      this.scheduleRestart();
    });
  }

  stop() {
    this.stopped = true;
    this.clearRestart();

    if (this.process) {
      try {
        this.process.kill();
      } catch (error) {
        // Ignore termination errors.
      } finally {
        this.process = null;
      }
    }
  }

  scheduleRestart() {
    this.clearRestart();
    this.restartTimer = setTimeout(() => {
      if (!this.stopped) {
        this.start();
      }
    }, 500);
    if (typeof this.restartTimer.unref === 'function') {
      this.restartTimer.unref();
    }
  }

  clearRestart() {
    if (this.restartTimer) {
      clearTimeout(this.restartTimer);
      this.restartTimer = null;
    }
  }

  handleStdout(chunk) {
    this.stdoutBuffer += String(chunk || '');
    const lines = this.stdoutBuffer.split(/\r?\n/);
    this.stdoutBuffer = lines.pop() || '';

    lines.forEach((lineRaw) => {
      const line = lineRaw.trim();
      if (!line) {
        return;
      }
      if (line === 'HOOK_FAILED') {
        console.error('[ExcelHook] WinEvent hook failed to initialize.');
        return;
      }
      if (line.startsWith('EXCEL_ACTIVE')) {
        const [, hwndRaw] = line.split('|');
        let hwnd = null;
        if (hwndRaw && /^-?\d+$/.test(hwndRaw)) {
          try {
            hwnd = BigInt(hwndRaw);
          } catch (error) {
            hwnd = null;
          }
        }
        this.emitState({ isExcelActive: true, hwnd });
        return;
      }
      if (line === 'EXCEL_INACTIVE') {
        this.emitState({ isExcelActive: false, hwnd: null });
      }
    });
  }

  emitState(state) {
    const key =
      state && state.isExcelActive
        ? `active:${state.hwnd ? state.hwnd.toString() : 'none'}`
        : 'inactive';
    if (this.lastState === key) {
      return;
    }
    this.lastState = key;
    if (typeof this.onStateChange === 'function') {
      this.onStateChange(state);
    }
  }

  buildScript() {
    return [
      '$ErrorActionPreference = "SilentlyContinue";',
      'Add-Type -TypeDefinition @\'',
      'using System;',
      'using System.Diagnostics;',
      'using System.Runtime.InteropServices;',
      'public static class WinEventBridge {',
      '  public delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint idEventThread, uint eventTime);',
      '  [DllImport("user32.dll")] public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);',
      '  [DllImport("user32.dll")] public static extern bool UnhookWinEvent(IntPtr hWinEventHook);',
      '  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();',
      '  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);',
      '}',
      '\'@ | Out-Null;',
      'Add-Type -AssemblyName System.Windows.Forms;',
      '$EVENT_SYSTEM_FOREGROUND = 0x0003;',
      '$WINEVENT_OUTOFCONTEXT = 0x0000;',
      '$WINEVENT_SKIPOWNPROCESS = 0x0002;',
      '$script:lastState = "";',
      'function Publish-ExcelState([IntPtr]$hwnd) {',
      '  $targetProcessId = 0;',
      '  if ($hwnd -ne [IntPtr]::Zero) {',
      '    [WinEventBridge]::GetWindowThreadProcessId($hwnd, [ref]$targetProcessId) | Out-Null;',
      '  }',
      '  $processName = "";',
      '  if ($targetProcessId -ne 0) {',
      '    try { $processName = ([Diagnostics.Process]::GetProcessById([int]$targetProcessId)).ProcessName } catch { $processName = "" }',
      '  }',
      '  $state = if ($processName -ieq "EXCEL" -or $processName -ieq "EXCEL.EXE") { "EXCEL_ACTIVE|" + [string]$hwnd.ToInt64() } else { "EXCEL_INACTIVE" };',
      '  if ($state -ne $script:lastState) {',
      '    $script:lastState = $state;',
      '    [Console]::Out.WriteLine($state);',
      '    [Console]::Out.Flush();',
      '  }',
      '}',
      '$callback = [WinEventBridge+WinEventDelegate]{',
      '  param([IntPtr]$hWinEventHook, [uint32]$eventType, [IntPtr]$hwnd, [int]$idObject, [int]$idChild, [uint32]$idEventThread, [uint32]$eventTime)',
      '  if ($idObject -ne 0 -or $idChild -ne 0) { return }',
      '  Publish-ExcelState $hwnd;',
      '};',
      '$script:callbackRef = $callback;',
      '$hook = [WinEventBridge]::SetWinEventHook($EVENT_SYSTEM_FOREGROUND, $EVENT_SYSTEM_FOREGROUND, [IntPtr]::Zero, $callback, 0, 0, ($WINEVENT_OUTOFCONTEXT -bor $WINEVENT_SKIPOWNPROCESS));',
      'if ($hook -eq [IntPtr]::Zero) {',
      '  [Console]::Out.WriteLine("HOOK_FAILED");',
      '  [Console]::Out.Flush();',
      '  exit 1;',
      '}',
      'Publish-ExcelState ([WinEventBridge]::GetForegroundWindow());',
      'Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action { [WinEventBridge]::UnhookWinEvent($hook) | Out-Null } | Out-Null;',
      '$ctx = New-Object System.Windows.Forms.ApplicationContext;',
      '[System.Windows.Forms.Application]::Run($ctx);'
    ].join('\n');
  }
}

module.exports = ExcelForegroundHook;
