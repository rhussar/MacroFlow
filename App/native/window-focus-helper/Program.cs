using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;

namespace WindowFocusHelper;

internal static class Program
{
  private const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
  private const uint WINEVENT_SKIPOWNPROCESS = 0x0002;

  private const uint SWP_NOSIZE = 0x0001;
  private const uint SWP_NOMOVE = 0x0002;
  private const uint SWP_NOACTIVATE = 0x0010;
  private const uint SWP_SHOWWINDOW = 0x0040;
  private const int OBJID_WINDOW = 0;
  private const uint GA_ROOT = 2;
  private const uint GA_ROOTOWNER = 3;

  private const int FOREGROUND_DEBOUNCE_MS = 40;
  private const int EXCEL_INACTIVE_DEBOUNCE_MS = 140;

  private static readonly IntPtr HWND_TOPMOST = new(-1);
  private static readonly IntPtr HWND_NOTOPMOST = new(-2);
  private static readonly IntPtr HWND_BOTTOM = new(1);
  private static readonly int[] REASSERT_DELAYS_MS = [35, 90, 170, 280, 420];
  private static readonly int[] DEMOTE_REASSERT_DELAYS_MS = [30, 85, 170, 280];

  // Sync: core state (_targetWindow, _excelActive, _reassertToken).
  // DebounceSync / InactiveSync: timer bookkeeping only, so the STA
  // win-event callback never blocks on slow stdout writes under Sync.
  private static readonly object Sync = new();
  private static readonly object DebounceSync = new();
  private static readonly object InactiveSync = new();

  private static readonly JsonSerializerOptions JsonOptions = new()
  {
    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    WriteIndented = false,
    PropertyNameCaseInsensitive = true
  };

  private static WinEventDelegate? _hookCallback;
  private static IntPtr _hookHandle = IntPtr.Zero;

  private static IntPtr _targetWindow = IntPtr.Zero;
  private static bool _excelActive;
  private static int _reassertToken;
  private static string _lastStateKey = string.Empty;

  private static System.Threading.Timer? _foregroundDebounceTimer;
  private static IntPtr _pendingForeground = IntPtr.Zero;
  private static int _foregroundDebounceSequence;
  private static int _scheduledForegroundSequence;

  private static System.Threading.Timer? _excelInactiveDebounceTimer;
  private static IntPtr _pendingInactiveForeground = IntPtr.Zero;
  private static int _inactiveDebounceSequence;
  private static int _scheduledInactiveSequence;

  // -----------------------------------------------------------------------
  // Entry point
  // -----------------------------------------------------------------------

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
      WINEVENT_SKIPOWNPROCESS
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

  // -----------------------------------------------------------------------
  // Stdin command loop
  // -----------------------------------------------------------------------

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
      _excelActive = false;
      _reassertToken++;
      _lastStateKey = string.Empty;
    }

    CancelForegroundDebounce();
    CancelExcelInactiveDebounce();

    // Determine the initial foreground. If it is our own window (common at
    // startup when launched from the Excel ribbon), look for an Excel
    // window instead so we immediately overlay.
    var fg = GetForegroundWindow();
    if (fg == IntPtr.Zero || IsTargetSelf(fg))
    {
      var excelHwnd = FindExcelMainWindow();
      if (excelHwnd != IntPtr.Zero)
      {
        fg = excelHwnd;
      }
    }

    PublishForegroundState(fg);
  }

  private static IntPtr FindExcelMainWindow()
  {
    try
    {
      foreach (var proc in Process.GetProcessesByName("EXCEL"))
      {
        var hwnd = proc.MainWindowHandle;
        if (hwnd != IntPtr.Zero && IsWindow(hwnd))
        {
          return hwnd;
        }
      }
    }
    catch
    {
      // Process enumeration can fail; not critical.
    }

    return IntPtr.Zero;
  }

  // -----------------------------------------------------------------------
  // Foreground event handling & debounce
  // -----------------------------------------------------------------------

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
      var sequence = ++_foregroundDebounceSequence;
      _scheduledForegroundSequence = sequence;

      _foregroundDebounceTimer?.Dispose();
      _foregroundDebounceTimer = new System.Threading.Timer(
        _ => FlushDebouncedForeground(sequence),
        null,
        FOREGROUND_DEBOUNCE_MS,
        Timeout.Infinite
      );
    }
  }

  private static void FlushDebouncedForeground(int sequence)
  {
    IntPtr hwnd;
    lock (DebounceSync)
    {
      if (sequence != _scheduledForegroundSequence)
      {
        return;
      }

      hwnd = _pendingForeground;
      _pendingForeground = IntPtr.Zero;
      _scheduledForegroundSequence = 0;
      _foregroundDebounceTimer?.Dispose();
      _foregroundDebounceTimer = null;
    }

    // Use authoritative current foreground to avoid applying stale hook events.
    var currentForeground = GetForegroundWindow();
    if (currentForeground != IntPtr.Zero && IsWindow(currentForeground))
    {
      hwnd = currentForeground;
    }

    PublishForegroundState(hwnd);
  }

  private static void CancelForegroundDebounce()
  {
    lock (DebounceSync)
    {
      _foregroundDebounceSequence++;
      _scheduledForegroundSequence = 0;
      _pendingForeground = IntPtr.Zero;
      _foregroundDebounceTimer?.Dispose();
      _foregroundDebounceTimer = null;
    }
  }

  // -----------------------------------------------------------------------
  // Core state machine
  // -----------------------------------------------------------------------

  private static void PublishForegroundState(IntPtr foregroundHwnd)
  {
    if (foregroundHwnd == IntPtr.Zero || !IsWindow(foregroundHwnd))
    {
      foregroundHwnd = IntPtr.Zero;
    }

    var processName = GetProcessNameForWindow(foregroundHwnd);
    var isExcel = string.Equals(processName, "EXCEL", StringComparison.OrdinalIgnoreCase);

    if (isExcel)
    {
      CancelExcelInactiveDebounce();

      var excelHwnd = GetExcelOwnerWindow(foregroundHwnd);
      var excelRect = TryGetWindowRectPayload(excelHwnd);
      ApplyForegroundState(true, excelHwnd, excelRect, processName);
      return;
    }

    // If the foreground window belongs to our own target (MacroFlow),
    // keep the current state to avoid flicker when clicking on the overlay.
    if (foregroundHwnd != IntPtr.Zero && IsTargetSelf(foregroundHwnd))
    {
      CancelExcelInactiveDebounce();
      return;
    }

    bool shouldDebounceInactive;
    lock (Sync)
    {
      shouldDebounceInactive = _excelActive;
      // Immediately cancel pending topmost reasserts so they cannot
      // re-promote the window during the inactive debounce window.
      if (shouldDebounceInactive)
      {
        _reassertToken++;
      }
    }

    if (shouldDebounceInactive)
    {
      ScheduleExcelInactiveDebounce(foregroundHwnd);
      return;
    }

    CancelExcelInactiveDebounce();
    ApplyForegroundState(false, IntPtr.Zero, null, processName);
  }

  private static void ScheduleExcelInactiveDebounce(IntPtr foregroundHwnd)
  {
    lock (InactiveSync)
    {
      _pendingInactiveForeground = foregroundHwnd;
      var sequence = ++_inactiveDebounceSequence;
      _scheduledInactiveSequence = sequence;

      _excelInactiveDebounceTimer?.Dispose();
      _excelInactiveDebounceTimer = new System.Threading.Timer(
        _ => FlushExcelInactiveDebounce(sequence),
        null,
        EXCEL_INACTIVE_DEBOUNCE_MS,
        Timeout.Infinite
      );
    }
  }

  private static void FlushExcelInactiveDebounce(int sequence)
  {
    IntPtr hwnd;
    lock (InactiveSync)
    {
      if (sequence != _scheduledInactiveSequence)
      {
        return;
      }

      hwnd = _pendingInactiveForeground;
      _pendingInactiveForeground = IntPtr.Zero;
      _scheduledInactiveSequence = 0;
      _excelInactiveDebounceTimer?.Dispose();
      _excelInactiveDebounceTimer = null;
    }

    // Use current foreground to avoid stale transition events causing false re-activation.
    var currentForeground = GetForegroundWindow();
    if (currentForeground != IntPtr.Zero && IsWindow(currentForeground))
    {
      hwnd = currentForeground;
    }
    else if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
    {
      hwnd = IntPtr.Zero;
    }

    if (hwnd != IntPtr.Zero && IsWindow(hwnd))
    {
      var processName = GetProcessNameForWindow(hwnd);
      if (string.Equals(processName, "EXCEL", StringComparison.OrdinalIgnoreCase))
      {
        PublishForegroundState(hwnd);
        return;
      }

      ApplyForegroundState(false, IntPtr.Zero, null, processName);
      return;
    }

    ApplyForegroundState(false, IntPtr.Zero, null, string.Empty);
  }

  private static void CancelExcelInactiveDebounce()
  {
    lock (InactiveSync)
    {
      _inactiveDebounceSequence++;
      _scheduledInactiveSequence = 0;
      _pendingInactiveForeground = IntPtr.Zero;
      _excelInactiveDebounceTimer?.Dispose();
      _excelInactiveDebounceTimer = null;
    }
  }

  // -----------------------------------------------------------------------
  // Window Z-order management
  // -----------------------------------------------------------------------

  private static void ApplyForegroundState(bool isExcel, IntPtr excelHwnd, WindowRectPayload? excelRect, string processName)
  {
    IntPtr target;

    lock (Sync)
    {
      target = _targetWindow;

      if (target != IntPtr.Zero && !IsWindow(target))
      {
        Emit(new { type = "error", message = "Target window handle became invalid. Clearing target." });
        _targetWindow = IntPtr.Zero;
        target = IntPtr.Zero;
      }

      if (target != IntPtr.Zero)
      {
        if (isExcel && excelHwnd != IntPtr.Zero)
        {
          ApplyExcelActiveLocked(target);
        }
        else if (_excelActive)
        {
          // Only demote on the transition from active to inactive.
          // When already inactive, skip redundant DemoteWindow calls
          // that create unnecessary Win32 SetWindowPos churn.
          ApplyExcelInactiveLocked(target);
        }
      }

      _excelActive = isExcel;
    }

    EmitState(isExcel, excelHwnd, excelRect, processName, target);
  }

  private static void ApplyExcelActiveLocked(IntPtr target)
  {
    var token = ++_reassertToken;
    EnforceTopmost(target);
    ScheduleReassertions(token, REASSERT_DELAYS_MS, () => !_excelActive, EnforceTopmost);
  }

  private static void ApplyExcelInactiveLocked(IntPtr target)
  {
    _reassertToken++;
    DemoteWindow(target, "excel-inactive");
    ScheduleReassertions(
      _reassertToken, DEMOTE_REASSERT_DELAYS_MS,
      () => _excelActive,
      t => DemoteWindow(t, "demote-reassert")
    );
  }

  private static void ScheduleReassertions(int token, int[] delays, Func<bool> shouldCancel, Action<IntPtr> action)
  {
    foreach (var delay in delays)
    {
      _ = Task.Run(async () =>
      {
        await Task.Delay(delay).ConfigureAwait(false);
        lock (Sync)
        {
          if (shouldCancel() || token != _reassertToken || _targetWindow == IntPtr.Zero)
          {
            return;
          }

          if (!IsWindow(_targetWindow))
          {
            _targetWindow = IntPtr.Zero;
            return;
          }

          action(_targetWindow);
        }
      });
    }
  }

  private static void DemoteWindow(IntPtr target, string phase)
  {
    _ = TrySetWindowPos(
      target,
      HWND_NOTOPMOST,
      SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
      phase
    );

    // Place behind the current foreground window.
    // HWND_NOTOPMOST alone leaves the window at the TOP of all non-topmost
    // windows, so it can still visually cover the app the user switched to.
    var fg = GetForegroundWindow();
    if (fg != IntPtr.Zero && fg != target && IsWindow(fg))
    {
      _ = TrySetWindowPos(
        target,
        fg,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        phase + "-zorder"
      );
      return;
    }

    // If no valid foreground exists, force the target to the bottom as a
    // deterministic demotion fallback.
    _ = TrySetWindowPos(
      target,
      HWND_BOTTOM,
      SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
      phase + "-bottom"
    );
  }

  private static void EnforceTopmost(IntPtr target)
  {
    // Clear then re-set TOPMOST to force a full Z-order transition.
    // Without this, Windows can treat the TOPMOST call as a no-op
    // when the flag appears already set from a previous cycle.
    // SWP_SHOWWINDOW is omitted — the window is already visible,
    // and the WM_SHOWWINDOW processing it triggers can fight with
    // concurrent Excel activation z-order changes.
    _ = TrySetWindowPos(
      target,
      HWND_NOTOPMOST,
      SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
      "enforce-clear"
    );
    _ = TrySetWindowPos(
      target,
      HWND_TOPMOST,
      SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
      "enforce-topmost"
    );
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

  // -----------------------------------------------------------------------
  // Window query helpers
  // -----------------------------------------------------------------------

  private static IntPtr GetExcelOwnerWindow(IntPtr hwnd)
  {
    if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
    {
      return IntPtr.Zero;
    }

    var root = GetAncestor(hwnd, GA_ROOT);
    return root != IntPtr.Zero && IsWindow(root) ? root : hwnd;
  }

  private static bool IsTargetSelf(IntPtr foregroundHwnd)
  {
    if (foregroundHwnd == IntPtr.Zero || !IsWindow(foregroundHwnd))
    {
      return false;
    }

    IntPtr target;
    lock (Sync)
    {
      target = _targetWindow;
    }

    if (target == IntPtr.Zero || !IsWindow(target))
    {
      return false;
    }

    if (foregroundHwnd == target)
    {
      return true;
    }

    var root = GetAncestor(foregroundHwnd, GA_ROOT);
    if (root != IntPtr.Zero && root == target)
    {
      return true;
    }

    var rootOwner = GetAncestor(foregroundHwnd, GA_ROOTOWNER);
    if (rootOwner != IntPtr.Zero && rootOwner == target)
    {
      return true;
    }

    // Same-process check covers child/popups in single-process mode.
    _ = GetWindowThreadProcessId(foregroundHwnd, out var fgPid);
    _ = GetWindowThreadProcessId(target, out var targetPid);
    if (fgPid == 0 || targetPid == 0)
    {
      return false;
    }

    if (fgPid == targetPid)
    {
      return true;
    }

    // Multi-process Electron: renderer/process windows can have different PIDs,
    // so compare executable path for robust self-detection.
    try
    {
      using var fgProcess = Process.GetProcessById((int)fgPid);
      using var targetProcess = Process.GetProcessById((int)targetPid);

      var fgPath = fgProcess.MainModule?.FileName;
      var targetPath = targetProcess.MainModule?.FileName;
      if (!string.IsNullOrEmpty(fgPath) && !string.IsNullOrEmpty(targetPath))
      {
        return string.Equals(fgPath, targetPath, StringComparison.OrdinalIgnoreCase);
      }
    }
    catch
    {
      // Ignore process lookup/access errors and fall back to false.
    }

    return false;
  }

  private static WindowRectPayload? TryGetWindowRectPayload(IntPtr hwnd)
  {
    if (hwnd == IntPtr.Zero || !IsWindow(hwnd))
    {
      return null;
    }

    if (!GetWindowRect(hwnd, out var rect))
    {
      EmitWin32Error("GetWindowRect", Marshal.GetLastWin32Error(), "excel-state", hwnd, IntPtr.Zero);
      return null;
    }

    return new WindowRectPayload
    {
      Left = rect.Left,
      Top = rect.Top,
      Right = rect.Right,
      Bottom = rect.Bottom
    };
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
      using var process = Process.GetProcessById((int)processId);
      return process.ProcessName;
    }
    catch
    {
      return string.Empty;
    }
  }

  // -----------------------------------------------------------------------
  // JSON output
  // -----------------------------------------------------------------------

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

  private static void EmitState(bool excelActive, IntPtr excelHwnd, WindowRectPayload? excelRect, string processName, IntPtr target)
  {
    var rectKey = excelRect is null
      ? "none"
      : $"{excelRect.Left}:{excelRect.Top}:{excelRect.Right}:{excelRect.Bottom}";

    var key = $"{excelActive}:{excelHwnd.ToInt64()}:{target.ToInt64()}:{rectKey}:{processName}";
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
      excelRect,
      process = processName,
      targetHwnd = target != IntPtr.Zero ? target.ToInt64().ToString() : null
    });
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

  // -----------------------------------------------------------------------
  // Lifecycle helpers
  // -----------------------------------------------------------------------

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
    CancelExcelInactiveDebounce();

    lock (Sync)
    {
      _reassertToken++;
      if (_targetWindow != IntPtr.Zero && IsWindow(_targetWindow))
      {
        _ = TrySetWindowPos(
          _targetWindow,
          HWND_NOTOPMOST,
          SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW,
          "cleanup"
        );
      }

      _targetWindow = IntPtr.Zero;
    }

    CancelForegroundDebounce();

    if (_hookHandle != IntPtr.Zero)
    {
      _ = UnhookWinEvent(_hookHandle);
      _hookHandle = IntPtr.Zero;
    }
  }

  // -----------------------------------------------------------------------
  // Types
  // -----------------------------------------------------------------------

  private sealed class HelperCommand
  {
    public string? Type { get; set; }
    public string? Hwnd { get; set; }
  }

  private sealed class WindowRectPayload
  {
    public int Left { get; set; }
    public int Top { get; set; }
    public int Right { get; set; }
    public int Bottom { get; set; }
  }

  [StructLayout(LayoutKind.Sequential)]
  private struct RECT
  {
    public int Left;
    public int Top;
    public int Right;
    public int Bottom;
  }

  // -----------------------------------------------------------------------
  // Win32 interop
  // -----------------------------------------------------------------------

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

  [DllImport("user32.dll")]
  private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

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
}
