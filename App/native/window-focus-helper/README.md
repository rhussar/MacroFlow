# WindowFocusHelper

Native Windows helper process for MacroFlow window z-order behavior.

## Behavior
- Keeps MacroFlow above Excel when Excel is active.
- Returns MacroFlow to normal z-order when Excel is not active.
- Debounces foreground-change events (100ms) to reduce churn during rapid focus changes.
- Validates HWND values with `IsWindow` before Win32 operations.
- Emits Win32 error details if `SetWindowLongPtr` / `SetWindowPos` fail.

## Build (from `App/`)

```powershell
dotnet publish native/window-focus-helper/WindowFocusHelper.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:EnableCompressionInSingleFile=true -o native/window-focus-helper/bin-helper
```

Output executable:
- `native/window-focus-helper/bin-helper/WindowFocusHelper.exe`

## Runtime Protocol (stdin/stdout JSON lines)
Commands sent by Electron:
- `{"type":"setTarget","hwnd":"<window-handle>"}`
- `{"type":"ping"}`
- `{"type":"shutdown"}`

Events emitted by helper:
- `{"type":"state","excelActive":true|false,"excelRect":{...}|null,...}`
- `{"type":"pong"}`
- `{"type":"error",...}`

## Manual QA Scenarios
1. `Excel -> MacroFlow -> Excel -> Chrome`: verify MacroFlow stays above Excel only while Excel is foreground.
2. Kill `WindowFocusHelper.exe` while app runs: verify Electron restarts helper (max 3 attempts) or enters fallback always-on-top mode.
3. Close Excel while helper is active: verify owner/topmost cleanup and MacroFlow returns to normal z-order.
4. Multiple Excel windows: verify the active Excel window is used as owner.
5. Freeze/Not-Responding Excel: verify helper heartbeat timeout triggers recovery.
6. Start app without helper binary: verify fallback always-on-top mode is enabled and logged.

