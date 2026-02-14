using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;

namespace WindowFocusHelper;

internal static class Program
{
  private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
  private const uint WINEVENT_OUTOFCONTEXT = 0x0000;
  private const uint WINEVENT_SKIPOWNPROCESS = 0x0002;

  private const int GWL_HWNDPARENT = -8;
  private const uint SWP_NOSIZE = 0x0001;
  private const uint SWP_NOMOVE = 0x0002;
  private const uint SWP_NOACTIVATE = 0x0010;
  private const uint SWP_SHOWWINDOW = 0x0040;
  private const int OBJID_WINDOW = 0;

  private const int FOREGROUND_DEBOUNCE_MS = 100;

  private static readonly IntPtr HWND_TOPMOST = new(-1);
  private static readonly IntPtr HWND_NOTOPMOST = new(-2);
  private static readonly int[] REASSERT_DELAYS_MS = [35, 90, 170, 280, 420];

  private static readonly object Sync = new();
  private static readonly object DebounceSync = new();

  private static readonly JsonSerializerOptions JsonOptions = new()
  {
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    WriteIndented = false,
    PropertyNameCaseInsensitive = true
  };

  private static WinEventDelegate? _hookCallback;
  private static IntPtr _hookHandle = IntPtr.Zero;

  private static IntPtr _targetWindow = IntPtr.Zero;
  private static IntPtr _currentOwner = IntPtr.Zero;
  private static bool _excelActive;
  private static int _reassertToken;
  private static string _lastStateKey = string.Empty;

  private static System.Threading.Timer? _foregroundDebounceTimer;
  private static IntPtr _pendingForeground = IntPtr.Zero;

  [STAThread]
  private static int Main()
  {
    AppDomain.CurrentDomain.ProcessExit += (_, _) => Cleanup();

    _hookCallback = OnWinEvent;
    _hookHandle = SetWinEventHook(
      EVENT_SYSTEM_FOREGROUND,
      EVENT_SYSTEM_FOREGROUND,
      IntPtr.Zero,
      _hookCallback,
      0,
      0,
      WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS
    );

    if (_hookHandle == IntPtr.Zero)
    {
      Emit(new
      {
        type = "error",
        message = "Failed to install foreground window hook.",
        win32Error = Marshal.GetLastWin32Error()
      });
      return 1;
    }

    _ = Task.Run(ReadCommandsAsync);

    PublishForegroundState(GetForegroundWindow());

    Application.Run();

    Cleanup();
    return 0;
  }

  private static async Task ReadCommandsAsync()
  {
    try
    {
      while (true)
      {
        var line = await Console.In.ReadLineAsync().ConfigureAwait(false);
        if (line is null)
        {
          break;
        }

        var trimmed = line.Trim();
        if (trimmed.Length == 0)
        {
          continue;
        }

        HelperCommand? command;
        try
        {
          command = JsonSerializer.Deserialize<HelperCommand>(trimmed, JsonOptions);
        }
        catch
        {
          Emit(new { type = "error", message = "Invalid command JSON." });
          continue;
        }

        if (command is null || string.IsNullOrWhiteSpace(command.Type))
        {
          continue;
        }

        switch (command.Type)
        {
          case "setTarget":
            HandleSetTarget(command.Hwnd);
            break;
          case "shutdown":
            Cleanup();
            Application.Exit();
            return;
          case "ping":
            Emit(new { type = "pong" });
            break;
        }
      }
    }
    catch (Exception ex)
    {
      Emit(new { type = "error", message = ex.Message });
    }

    Cleanup();
    Application.Exit();
  }

  private static void HandleSetTarget(string? hwndRaw)
  {
    if (!TryParseHwnd(hwndRaw, out var hwnd))
    {
      Emit(new { type = "error", message = "Invalid target hwnd." });
      return;
    }

    if (!IsWindow(hwnd))
    {
      Emit(new { type = "error", message = "Target hwnd is not a valid window." });
      return;
    }

    lock (Sync)
    {
      _targetWindow = hwnd;
      _reassertToken++;
      _currentOwner = IntPtr.Zero;
    }

    ScheduleForegroundPublish(GetForegroundWindow());
  }

  private static void OnWinEvent(
    IntPtr hWinEventHook,
    uint eventType,
    IntPtr hwnd,
    int idObject,
    int idChild,
    uint idEventThread,
    uint eventTime
  )
  {
    if (idObject != OBJID_WINDOW || idChild != 0)
    {
      return;
    }

    ScheduleForegroundPublish(hwnd);
  }

  private static void ScheduleForegroundPublish(IntPtr hwnd)
  {
    lock (DebounceSync)
    {
      _pendingForeground = hwnd;

      if (_foregroundDebounceTimer is null)
      {
        _foregroundDebounceTimer = new System.Threading.Timer(
          _ => FlushDebouncedForeground(),
          null,
          FOREGROUND_DEBOUNCE_MS,
          Timeout.Infinite
        );
      }
      else
      {
        _foregroundDebounceTimer.Change(FOREGROUND_DEBOUNCE_MS, Timeout.Infinite);
      }
    }
  }

  private static void FlushDebouncedForeground()
  {
    IntPtr hwnd;
    lock (DebounceSync)
    {
      hwnd = _pendingForeground;
    }

    PublishForegroundState(hwnd);
  }

  private static void PublishForegroundState(IntPtr foregroundHwnd)
  {
    if (foregroundHwnd == IntPtr.Zero || !IsWindow(foregroundHwnd))
    {
      foregroundHwnd = IntPtr.Zero;
    }

    var processName = GetProcessNameForWindow(foregroundHwnd);
    var isExcel = string.Equals(processName, "EXCEL", StringComparison.OrdinalIgnoreCase);

    IntPtr target;
    IntPtr excelHwnd = isExcel ? foregroundHwnd : IntPtr.Zero;

    lock (Sync)
    {
      target = _targetWindow;

      if (target != IntPtr.Zero && !IsWindow(target))
      {
        Emit(new { type = "error", message = "Target window handle became invalid. Clearing target." });
        _targetWindow = IntPtr.Zero;
        _currentOwner = IntPtr.Zero;
        target = IntPtr.Zero;
      }

      if (target != IntPtr.Zero)
      {
        if (isExcel && excelHwnd != IntPtr.Zero)
        {
          ApplyExcelActiveLocked(target, excelHwnd);
        }
        else
        {
          ApplyExcelInactiveLocked(target);
        }
      }

      _excelActive = isExcel;
    }

    EmitState(isExcel, excelHwnd, processName, target);
  }

  private static void ApplyExcelActiveLocked(IntPtr target, IntPtr excelHwnd)
  {
    if (_currentOwner != excelHwnd)
    {
      if (TrySetWindowOwner(target, excelHwnd, "excel-active"))
      {
        _currentOwner = excelHwnd;
      }
    }

    var token = ++_reassertToken;
    EnforceTopmost(target);

    foreach (var delay in REASSERT_DELAYS_MS)
    {
      _ = Task.Run(async () =>
      {
        await Task.Delay(delay).ConfigureAwait(false);
        lock (Sync)
        {
          if (!_excelActive || token != _reassertToken || _targetWindow == IntPtr.Zero)
          {
            return;
          }

          if (!IsWindow(_targetWindow))
          {
            _targetWindow = IntPtr.Zero;
            _currentOwner = IntPtr.Zero;
            return;
          }

          EnforceTopmost(_targetWindow);
        }
      });
    }
  }

  private static void ApplyExcelInactiveLocked(IntPtr target)
  {
    _reassertToken++;

    if (_currentOwner != IntPtr.Zero)
    {
      if (TrySetWindowOwner(target, IntPtr.Zero, "excel-inactive"))
      {
        _currentOwner = IntPtr.Zero;
      }
    }

    _ = TrySetWindowPos(
      target,
      HWND_NOTOPMOST,
      SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
      "excel-inactive"
    );
  }

  private static void EnforceTopmost(IntPtr target)
  {
    _ = TrySetWindowPos(
      target,
      HWND_TOPMOST,
      SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
      "excel-active"
    );
  }

  private static bool TrySetWindowOwner(IntPtr target, IntPtr owner, string phase)
  {
    if (!IsWindow(target))
    {
      Emit(new { type = "error", message = "Cannot set owner: target hwnd invalid.", phase });
      return false;
    }

    if (owner != IntPtr.Zero && !IsWindow(owner))
    {
      Emit(new { type = "error", message = "Cannot set owner: owner hwnd invalid.", phase });
      return false;
    }

    SetLastError(0);
    _ = SetWindowLongPtrCompat(target, GWL_HWNDPARENT, owner);
    var error = Marshal.GetLastWin32Error();

    if (error != 0)
    {
      EmitWin32Error("SetWindowLongPtr", error, phase, target, owner);
      return false;
    }

    return true;
  }

  private static bool TrySetWindowPos(IntPtr target, IntPtr insertAfter, uint flags, string phase)
  {
    if (!IsWindow(target))
    {
      Emit(new { type = "error", message = "Cannot set window position: target hwnd invalid.", phase });
      return false;
    }

    var ok = SetWindowPos(target, insertAfter, 0, 0, 0, 0, flags);
    if (ok)
    {
      return true;
    }

    EmitWin32Error("SetWindowPos", Marshal.GetLastWin32Error(), phase, target, insertAfter);
    return false;
  }

  private static void EmitWin32Error(string api, int error, string phase, IntPtr target, IntPtr related)
  {
    Emit(new
    {
      type = "error",
      message = $"{api} failed.",
      api,
      phase,
      win32Error = error,
      targetHwnd = target.ToInt64().ToString(),
      relatedHwnd = related.ToInt64().ToString()
    });
  }

  private static void EmitState(bool excelActive, IntPtr excelHwnd, string processName, IntPtr target)
  {
    var key = $"{excelActive}:{excelHwnd.ToInt64()}:{target.ToInt64()}";
    lock (Sync)
    {
      if (_lastStateKey == key)
      {
        return;
      }
      _lastStateKey = key;
    }

    Emit(new
    {
      type = "state",
      excelActive,
      excelHwnd = excelActive ? excelHwnd.ToInt64().ToString() : null,
      process = processName,
      targetHwnd = target != IntPtr.Zero ? target.ToInt64().ToString() : null
    });
  }

  private static string GetProcessNameForWindow(IntPtr hwnd)
  {
    if (hwnd == IntPtr.Zero)
    {
      return string.Empty;
    }

    _ = GetWindowThreadProcessId(hwnd, out var processId);
    if (processId == 0)
    {
      return string.Empty;
    }

    try
    {
      return Process.GetProcessById((int)processId).ProcessName;
    }
    catch (ArgumentException)
    {
      return string.Empty;
    }
    catch (InvalidOperationException)
    {
      return string.Empty;
    }
    catch
    {
      return string.Empty;
    }
  }

  private static bool TryParseHwnd(string? raw, out IntPtr hwnd)
  {
    hwnd = IntPtr.Zero;
    if (string.IsNullOrWhiteSpace(raw))
    {
      return false;
    }

    if (!long.TryParse(raw, out var value))
    {
      return false;
    }

    hwnd = new IntPtr(value);
    return hwnd != IntPtr.Zero;
  }

  private static void Cleanup()
  {
    lock (Sync)
    {
      _reassertToken++;
      if (_targetWindow != IntPtr.Zero && IsWindow(_targetWindow))
      {
        if (_currentOwner != IntPtr.Zero)
        {
          if (TrySetWindowOwner(_targetWindow, IntPtr.Zero, "cleanup"))
          {
            _currentOwner = IntPtr.Zero;
          }
        }

        _ = TrySetWindowPos(
          _targetWindow,
          HWND_NOTOPMOST,
          SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
          "cleanup"
        );
      }

      _targetWindow = IntPtr.Zero;
      _currentOwner = IntPtr.Zero;
    }

    lock (DebounceSync)
    {
      _foregroundDebounceTimer?.Dispose();
      _foregroundDebounceTimer = null;
      _pendingForeground = IntPtr.Zero;
    }

    if (_hookHandle != IntPtr.Zero)
    {
      _ = UnhookWinEvent(_hookHandle);
      _hookHandle = IntPtr.Zero;
    }
  }

  private static void Emit(object payload)
  {
    try
    {
      Console.WriteLine(JsonSerializer.Serialize(payload, JsonOptions));
      Console.Out.Flush();
    }
    catch
    {
      // Ignore broken pipe or serialization failures.
    }
  }

  private sealed class HelperCommand
  {
    public string? Type { get; set; }
    public string? Hwnd { get; set; }
  }

  private delegate void WinEventDelegate(
    IntPtr hWinEventHook,
    uint eventType,
    IntPtr hwnd,
    int idObject,
    int idChild,
    uint idEventThread,
    uint eventTime
  );

  [DllImport("user32.dll")]
  private static extern IntPtr SetWinEventHook(
    uint eventMin,
    uint eventMax,
    IntPtr hmodWinEventProc,
    WinEventDelegate lpfnWinEventProc,
    uint idProcess,
    uint idThread,
    uint dwFlags
  );

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

  [DllImport("user32.dll")]
  private static extern IntPtr GetForegroundWindow();

  [DllImport("user32.dll")]
  private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  private static extern bool IsWindow(IntPtr hWnd);

  [DllImport("kernel32.dll")]
  private static extern void SetLastError(uint dwErrCode);

  [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
  private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

  [DllImport("user32.dll", EntryPoint = "SetWindowLongW", SetLastError = true)]
  private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  private static extern bool SetWindowPos(
    IntPtr hWnd,
    IntPtr hWndInsertAfter,
    int x,
    int y,
    int cx,
    int cy,
    uint uFlags
  );

  private static IntPtr SetWindowLongPtrCompat(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
  {
    if (IntPtr.Size == 8)
    {
      return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
    }

    return new IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32()));
  }
}


