Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -ReferencedAssemblies @("System.Windows.Forms", "System.Drawing") -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

public class HotKeyWindow : Form {
    public event Action<int> HotKeyPressed;

    [DllImport("user32.dll")]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    protected override void WndProc(ref Message m) {
        const int WM_HOTKEY = 0x0312;
        if (m.Msg == WM_HOTKEY) {
            int id = m.WParam.ToInt32();
            if (HotKeyPressed != null) {
                HotKeyPressed(id);
            }
        }
        base.WndProc(ref m);
    }
}

public static class Win32Drag {
    [DllImport("user32.dll")]
    public static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
}
"@

$AppName = "MyTime"
$AutoSaveSeconds = 60

function Get-AppDataDir {
    $base = $env:APPDATA
    if (-not $base) { throw "APPDATA missing" }
    return Join-Path $base $AppName
}

function Get-DesktopDir {
    $userHome = $env:USERPROFILE
    if (-not $userHome) { throw "USERPROFILE missing" }
    return Join-Path $userHome "Desktop"
}

function Get-AppRoot {
    return Split-Path -Parent $PSCommandPath
}

function Load-AppIcon {
    $root = Get-AppRoot
    $icoPath = Join-Path $root "icon.ico"
    if (Test-Path $icoPath) {
        try {
            return New-Object System.Drawing.Icon($icoPath)
        } catch {
        }
    }
    return [System.Drawing.SystemIcons]::Information
}

function New-DefaultData {
    return [ordered]@{
        timers = @()
        active_timer_id = $null
        active_session_id = $null
    }
}

function New-DefaultSettings {
    return [ordered]@{
        write_daily_desktop_log = $true
        write_per_ticket_desktop_log = $true
        desktop_path_override = ""
        floating_font_scale = 1.0
    }
}

function Write-JsonAtomic {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)]$Value
    )
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $tmp = [IO.Path]::ChangeExtension($Path, ".tmp")
    $json = $Value | ConvertTo-Json -Depth 10
    [IO.File]::WriteAllText($tmp, $json)
    Move-Item -Path $tmp -Destination $Path -Force
}

function Read-JsonOrDefault {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)]$DefaultValue
    )
    if (-not (Test-Path $Path)) {
        return $DefaultValue
    }
    $raw = Get-Content -Path $Path -Raw
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $DefaultValue
    }
    return $raw | ConvertFrom-Json
}

function New-Id {
    return [guid]::NewGuid().ToString()
}

function Utc-Now {
    # Wall-clock UTC for timestamps/logging.
    return [DateTime]::UtcNow
}

function Mono-NowMs {
    # Monotonic-ish milliseconds since system boot. Not subject to wall-clock adjustments.
    # (Used for elapsed time to avoid visual/timekeeping glitches from clock changes or drift.)
    try {
        # Avoid referencing ::TickCount64 directly on older .NET where it doesn't exist.
        $prop = [System.Environment].GetProperty("TickCount64", [System.Reflection.BindingFlags]"Public,Static")
        if ($prop) {
            return [long]$prop.GetValue($null, $null)
        }
    } catch {
    }

    # Fallback: use a process-local stopwatch (works on Windows PowerShell / .NET Framework too).
    if (-not (Get-Variable -Name monoSw -Scope Script -ErrorAction SilentlyContinue) -or -not $script:monoSw) {
        $script:monoSw = [System.Diagnostics.Stopwatch]::StartNew()
    }
    try {
        return [long]$script:monoSw.ElapsedMilliseconds
    } catch {
        return 0
    }
}

function Mono-NowSeconds {
    # Truncated seconds from Mono-NowMs, useful for per-second caching.
    return [long]((Mono-NowMs) / 1000)
}

function To-UtcString {
    param([DateTime]$Dt)
    return $Dt.ToUniversalTime().ToString("o")
}

function From-UtcString {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    return [DateTime]::Parse($Value, $null, [Globalization.DateTimeStyles]::RoundtripKind)
}

function Sanitize-Filename {
    param([string]$Value)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        if ($invalid -contains $ch -or [char]::IsControl($ch)) {
            [void]$sb.Append('_')
        } else {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString()
}

function New-InputDialog {
    param(
        [string]$Title,
        [string]$Prompt,
        [string]$DefaultValue = ""
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = "CenterScreen"
    $form.Width = 440
    $form.Height = 220
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.AutoSize = $true
    $label.Left = 12
    $label.Top = 12
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Left = 12
    $textBox.Top = ($label.Bottom + 10)
    $textBox.Width = 400
    $textBox.Text = $DefaultValue
    $form.Controls.Add($textBox)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "OK"
    $ok.Left = 250
    $ok.Top = ($textBox.Bottom + 15)
    $ok.Width = 75
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.Left = 335
    $cancel.Top = ($textBox.Bottom + 15)
    $cancel.Width = 75
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)

    $form.AcceptButton = $ok
    $form.CancelButton = $cancel
    $form.Add_Shown({
        $form.Activate()
        $form.BringToFront()
    })

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    }
    return $null
}

function New-SelectDialog {
    param(
        [string]$Title,
        [string]$Prompt,
        [string[]]$Items
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = "CenterScreen"
    $form.Width = 420
    $form.Height = 360
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.AutoSize = $true
    $label.Left = 12
    $label.Top = 12
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Left = 12
    $listBox.Top = 35
    $listBox.Width = 380
    $listBox.Height = 230
    $listBox.SelectionMode = "One"
    $listBox.Items.AddRange($Items)
    $form.Controls.Add($listBox)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "OK"
    $ok.Left = 230
    $ok.Top = 275
    $ok.Width = 75
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.Left = 315
    $cancel.Top = 275
    $cancel.Width = 75
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)

    $form.AcceptButton = $ok
    $form.CancelButton = $cancel
    $form.Add_Shown({
        $form.Activate()
        $form.BringToFront()
    })

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $listBox.SelectedItem) {
        return [string]$listBox.SelectedItem
    }
    return $null
}

function New-ConfirmDialog {
    param(
        [string]$Title,
        [string]$Message,
        [string]$YesText = "Yes",
        [string]$NoText = "No"
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = "CenterScreen"
    $form.Width = 420
    $form.Height = 170
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Left = 12
    $label.Top = 12
    $form.Controls.Add($label)

    $yes = New-Object System.Windows.Forms.Button
    $yes.Text = $YesText
    $yes.Left = 210
    $yes.Top = 80
    $yes.Width = 80
    $yes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($yes)

    $no = New-Object System.Windows.Forms.Button
    $no.Text = $NoText
    $no.Left = 300
    $no.Top = 80
    $no.Width = 80
    $no.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($no)

    $form.AcceptButton = $no
    $form.CancelButton = $no
    $form.Add_Shown({
        $form.Activate()
        $form.BringToFront()
    })

    $result = $form.ShowDialog()
    return $result -eq [System.Windows.Forms.DialogResult]::Yes
}

function Get-PropValue {
    param(
        [Parameter(Mandatory=$true)]$Obj,
        [Parameter(Mandatory=$true)][string]$Name
    )
    if ($null -eq $Obj) { return $null }
    if ($Obj -is [System.Collections.IDictionary]) {
        if ($Obj.Contains($Name)) { return $Obj[$Name] }
        return $null
    }
    if ($Obj.PSObject -and $Obj.PSObject.Properties[$Name]) {
        return $Obj.$Name
    }
    return $null
}

function Set-PropValue {
    param(
        [Parameter(Mandatory=$true)]$Obj,
        [Parameter(Mandatory=$true)][string]$Name,
        $Value
    )
    if ($null -eq $Obj) { return }

    if ($Obj -is [System.Collections.IDictionary]) {
        $Obj[$Name] = $Value
        return
    }

    if ($Obj.PSObject -and $Obj.PSObject.Properties[$Name]) {
        $Obj.$Name = $Value
        return
    }

    try {
        $Obj | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    } catch {
    }
}

function New-NoteDialog {
    param([string]$Label, [string]$Duration, [string]$SessionDuration)
    $lines = @("Stopped: $Label", "Total: $Duration")
    if ($SessionDuration) {
        $lines += "Last Session: $SessionDuration"
    }
    $lines += "Optional note:"
    $prompt = $lines -join "`n"
    return New-InputDialog -Title "Stop Timer" -Prompt $prompt -DefaultValue ""
}

$appDir = Get-AppDataDir
$dataPath = Join-Path $appDir "data.json"
$settingsPath = Join-Path $appDir "settings.json"

$State = [ordered]@{
    data = $null
    settings = $null
    runtime = [ordered]@{
        last_autosave_at = $null
        recovery_prompt = $null
    }
}

function Invoke-Safe {
    param([scriptblock]$Action)
    try {
        & $Action
    } catch {
        try {
            $msg = $_.Exception.Message
            $logDir = Get-AppDataDir
            if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
            $logPath = Join-Path $logDir "mytime-error.log"
            $detail = $null
            try { $detail = $_ | Out-String } catch { $detail = $msg }
            $line = "[{0}] {1}`r`n{2}`r`n" -f (Get-Date).ToString("s"), $msg, $detail.TrimEnd()
            Add-Content -Path $logPath -Value $line
        } catch {
        }
    }
}

function Get-PerfLogPaths {
    $paths = @()
    try {
        $logDir = Get-AppDataDir
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        $paths += (Join-Path $logDir "mytime-perf.log")
    } catch {
    }
    try {
        $root = Get-AppRoot
        if ($root -and (Test-Path $root)) {
            $paths += (Join-Path $root "mytime-perf.log")
        }
    } catch {
    }
    return @($paths | Where-Object { $_ } | Select-Object -Unique)
}

function Perf-AppendLine {
    param([string]$Line)
    foreach ($p in (Get-PerfLogPaths)) {
        try {
            [IO.File]::AppendAllText($p, ($Line + "`r`n"))
        } catch {
        }
    }
}

function Perf-Init {
    try {
        $ts = (Get-Date).ToString("s")
        $ver = $PSVersionTable.PSVersion.ToString()
        $pid = $PID
        Perf-AppendLine -Line ("[{0}] perf_init ps={1} pid={2}" -f $ts, $ver, $pid)
    } catch {
    }
}

function Perf-Note {
    param([string]$Line)
    try {
        if (-not (Get-Variable -Name perfRing -Scope Script -ErrorAction SilentlyContinue) -or -not $script:perfRing) {
            $script:perfRing = New-Object System.Collections.Generic.List[string]
        }
        if ($script:perfRing.Count -ge 400) {
            # Keep a small ring buffer; remove the oldest 100 entries.
            $script:perfRing.RemoveRange(0, 100)
        }
        $script:perfRing.Add($Line) | Out-Null

        # Also write-through so the file exists even if the app is closed without using Quit.
        Perf-AppendLine -Line $Line
    } catch {
    }
}

function Perf-Flush {
    try {
        if (-not (Get-Variable -Name perfRing -Scope Script -ErrorAction SilentlyContinue) -or -not $script:perfRing -or $script:perfRing.Count -eq 0) { return }
        $logDir = Get-AppDataDir
        if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
        $logPath = Join-Path $logDir "mytime-perf.log"
        Add-Content -Path $logPath -Value ($script:perfRing.ToArray())
        $script:perfRing.Clear()
    } catch {
    }
}

function Enable-DoubleBuffering {
    param([System.Windows.Forms.Control]$Control)
    if (-not $Control) { return }
    try {
        $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic")
        if ($prop) { $prop.SetValue($Control, $true, $null) }
    } catch {
    }
}

function Set-RoundedCorners {
    param(
        [Parameter(Mandatory=$true)][System.Windows.Forms.Control]$Control,
        [int]$Radius = 14
    )
    if (-not $Control -or $Control.IsDisposed) { return }
    if ($Radius -lt 0) { $Radius = 0 }

    $w = [int]$Control.ClientSize.Width
    $h = [int]$Control.ClientSize.Height
    if ($w -le 0 -or $h -le 0) { return }

    $r = [int][math]::Min([math]::Min($Radius, [int]($w / 2)), [int]($h / 2))
    $d = $r * 2

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    try {
        if ($r -eq 0) {
            $path.AddRectangle((New-Object System.Drawing.Rectangle(0, 0, $w, $h)))
        } else {
            $path.StartFigure() | Out-Null
            $path.AddArc(0, 0, $d, $d, 180, 90)
            $path.AddArc($w - $d, 0, $d, $d, 270, 90)
            $path.AddArc($w - $d, $h - $d, $d, $d, 0, 90)
            $path.AddArc(0, $h - $d, $d, $d, 90, 90)
            $path.CloseFigure()
        }

        $region = New-Object System.Drawing.Region($path)
        $old = $Control.Region
        $Control.Region = $region
        if ($old) { try { $old.Dispose() } catch {} }
    } finally {
        try { $path.Dispose() } catch {}
    }
}

function Save-Data {
    Write-JsonAtomic -Path $dataPath -Value $State.data
}

function Save-Settings {
    Write-JsonAtomic -Path $settingsPath -Value $State.settings
}

function Save-All {
    Save-Data
    Save-Settings
}

function Get-TimerById {
    param([string]$TimerId)
    return $State.data.timers | Where-Object { $_.id -eq $TimerId } | Select-Object -First 1
}

function Get-ActiveTimer {
    if (-not $State.data.active_timer_id) { return $null }
    return Get-TimerById -TimerId $State.data.active_timer_id
}

function Get-ActiveSession {
    $timer = Get-ActiveTimer
    if (-not $timer) { return $null }
    if (-not $State.data.active_session_id) { return $null }
    return $timer.sessions | Where-Object { $_.id -eq $State.data.active_session_id } | Select-Object -First 1
}

function Clear-ActiveTracking {
    $State.data.active_timer_id = $null
    $State.data.active_session_id = $null
}

function Close-ActiveSession {
    param([DateTime]$At)
    $timer = Get-ActiveTimer
    if (-not $timer) {
        if ($State.data.active_session_id) {
            Clear-ActiveTracking
            Clear-ActiveAnchor
            Invalidate-TodayCache
            Save-Data
        }
        return $null
    }

    $session = Get-ActiveSession
    if (-not $session) {
        Clear-ActiveTracking
        Clear-ActiveAnchor
        Invalidate-TodayCache
        Save-Data
        return $null
    }

    $durationSeconds = Get-ActiveSessionElapsedSeconds

    if (-not (Get-PropValue -Obj $session -Name "end")) {
        Set-PropValue -Obj $session -Name "end" -Value (To-UtcString $At)
    }
    Set-PropValue -Obj $session -Name "duration_seconds" -Value ([int]$durationSeconds)
    try { $durationSeconds = [int](Get-PropValue -Obj $session -Name "duration_seconds") } catch { $durationSeconds = 0 }
    if ($durationSeconds -lt 0) {
        $durationSeconds = 0
        Set-PropValue -Obj $session -Name "duration_seconds" -Value 0
    }

    $summary = [ordered]@{
        timer_id = $timer.id
        session_id = $session.id
        label = $timer.label
        duration_seconds = [int]$durationSeconds
        start = $session.start
        end = $session.end
    }

    Clear-ActiveTracking
    Clear-ActiveAnchor
    Invalidate-TodayCache
    Save-Data
    return $summary
}

function Start-TimerInternal {
    param([string]$TimerId)
    $now = Utc-Now
    $null = Close-ActiveSession -At $now

    $timer = Get-TimerById -TimerId $TimerId
    if (-not $timer) { throw "Timer not found" }

    $session = [ordered]@{
        id = New-Id
        start = To-UtcString $now
        end = $null
        duration_seconds = 0
        note = ""
        exported = $false
    }
    $timer.sessions += $session

    $State.data.active_timer_id = [string]$TimerId
    $State.data.active_session_id = $session.id
    Save-Data
    Start-ActiveAnchor -TimerId ([string]$TimerId) -SessionId ([string]$session.id) -BaseSeconds 0
}

function Create-OrStartTimer {
    param([string]$Label)
    $trimmed = $Label.Trim()
    if (-not $trimmed) { return }
    $existing = $State.data.timers | Where-Object { $_.label -and $_.label.ToLower() -eq $trimmed.ToLower() } | Select-Object -First 1
    if ($existing) {
        Start-TimerInternal -TimerId $existing.id
        return
    }
    $timer = [ordered]@{
        id = New-Id
        label = $trimmed
        created_at = To-UtcString (Utc-Now)
        sessions = @()
    }
    $State.data.timers += $timer
    Start-TimerInternal -TimerId $timer.id
}

function Most-RecentTimerId {
    $timers = $State.data.timers
    if (-not $timers -or $timers.Count -eq 0) { return $null }
    $sorted = $timers | Sort-Object -Property @{
        Expression = {
            $t = $_
            $last = $t.sessions | ForEach-Object {
                if ($_.end) { From-UtcString $_.end } else { From-UtcString $_.start }
            } | Sort-Object -Descending | Select-Object -First 1
            if ($last) { $last } else { From-UtcString $t.created_at }
        }
        Descending = $true
    }
    return $sorted[0].id
}

function Active-ElapsedSeconds {
    $timer = Get-ActiveTimer
    if (-not $timer) { return 0 }
    # Show total time for the active timer (ended sessions + running session if any).
    # When paused there is no active_session_id, so session-elapsed would be 0 which looks like a reset.
    return (Timer-TotalSeconds -Timer $timer)
}

function Timer-TotalSeconds {
    param($Timer)
    if (-not $Timer) { return 0 }
    Ensure-RuntimeClock

    $timerId = [string](Get-PropValue -Obj $Timer -Name "id")
    $activeTimerId = [string]$State.data.active_timer_id
    $activeSessionId = [string]$State.data.active_session_id

    $total = 0
    foreach ($session in @($Timer.sessions | Where-Object { $_ })) {
        $endStr = Get-PropValue -Obj $session -Name "end"
        $sid = [string](Get-PropValue -Obj $session -Name "id")

        if (-not $endStr) {
            # Running session: only count if it's the current active session.
            if ($timerId -and $activeTimerId -and $activeSessionId -and $timerId -eq $activeTimerId -and $sid -eq $activeSessionId) {
                $total += [int](Get-ActiveSessionElapsedSeconds)
            }
            continue
        }

        $durVal = Get-PropValue -Obj $session -Name "duration_seconds"
        $dur = 0
        if ($null -ne $durVal) {
            try { $dur = [int]$durVal } catch { $dur = 0 }
        } else {
            # Older data: derive duration from timestamps once, then persist it.
            $startUtc = From-UtcString (Get-PropValue -Obj $session -Name "start")
            $endUtc = From-UtcString $endStr
            if ($startUtc -and $endUtc) {
                $dur = [int][math]::Max(0, ($endUtc - $startUtc).TotalSeconds)
            }
            Set-PropValue -Obj $session -Name "duration_seconds" -Value $dur
        }

        if ($dur -gt 0) { $total += $dur }
    }

    if ($total -lt 0) { $total = 0 }
    return [int]$total
}

function Total-AllTimersSeconds {
    Ensure-RuntimeClock
    $sum = 0
    foreach ($t in @($State.data.timers | Where-Object { $_ })) {
        $sum += [int](Timer-TotalSeconds -Timer $t)
    }
    if ($sum -lt 0) { $sum = 0 }
    return [int]$sum
}

function Ensure-RuntimeClock {
    if (-not $State -or -not $State.runtime) { return }

    if (-not $State.runtime.Contains("mono_epoch_ms") -or $null -eq $State.runtime.mono_epoch_ms) {
        $State.runtime.mono_epoch_ms = Mono-NowMs
    }
    if (-not $State.runtime.Contains("active_anchor")) { $State.runtime.active_anchor = $null }
    if (-not $State.runtime.Contains("last_autosave_mono_s")) { $State.runtime.last_autosave_mono_s = $null }
    if (-not $State.runtime.Contains("last_active_elapsed_s")) { $State.runtime.last_active_elapsed_s = 0 }
    if (-not $State.runtime.Contains("last_active_timer_id")) { $State.runtime.last_active_timer_id = "" }
    if (-not $State.runtime.Contains("last_active_session_id")) { $State.runtime.last_active_session_id = "" }

    # Cache: ended-session totals for "today" (fast UI updates).
    if (-not $State.runtime.Contains("cache_day_local") -or -not $State.runtime.cache_day_local) { $State.runtime.cache_day_local = (Get-Date).Date }
    if (-not $State.runtime.Contains("cache_by_timer") -or -not ($State.runtime.cache_by_timer -is [System.Collections.IDictionary])) { $State.runtime.cache_by_timer = @{} }
    if (-not $State.runtime.Contains("cache_total")) { $State.runtime.cache_total = 0 }
    if (-not $State.runtime.Contains("cache_dirty")) { $State.runtime.cache_dirty = $true }

    # UI: request a full refresh of displayed timer totals (used to avoid per-tick heavy recalculation).
    if (-not $State.runtime.Contains("ui_times_dirty")) { $State.runtime.ui_times_dirty = $true }
}

function Invalidate-TodayCache {
    Ensure-RuntimeClock
    $State.runtime.cache_dirty = $true
    $State.runtime.ui_times_dirty = $true
}

function Ensure-TodayCache {
    Ensure-RuntimeClock
    $todayLocal = (Get-Date).Date

    if (-not $State.runtime.cache_day_local -or $State.runtime.cache_day_local -ne $todayLocal) {
        Rebuild-TodayCache -DayLocal $todayLocal
        return
    }

    $dirty = $false
    try { $dirty = [bool]$State.runtime.cache_dirty } catch { $dirty = $false }
    if ($dirty) {
        Rebuild-TodayCache -DayLocal $todayLocal
    }
}

function Get-ClockTickSeconds {
    Ensure-RuntimeClock
    return [int](Mono-NowSeconds)
}

function Rebuild-TodayCache {
    param([DateTime]$DayLocal)
    Ensure-RuntimeClock
    if (-not $DayLocal) { $DayLocal = (Get-Date).Date }

    $dayStartUtc = $DayLocal.ToUniversalTime()
    $dayEndUtc = $DayLocal.AddDays(1).ToUniversalTime()

    $cache = @{}
    $grand = 0

    foreach ($timer in @($State.data.timers | Where-Object { $_ })) {
        $timerId = [string](Get-PropValue -Obj $timer -Name "id")
        if (-not $timerId) { continue }

        $timerTotal = 0
        foreach ($session in @($timer.sessions | Where-Object { $_ })) {
            $startUtc = From-UtcString (Get-PropValue -Obj $session -Name "start")
            if (-not $startUtc) { continue }

            # Base cache only includes ended/paused sessions. Running session is added separately.
            $endStr = Get-PropValue -Obj $session -Name "end"
            if (-not $endStr) { continue }

            $durationSeconds = $null
            $durVal = Get-PropValue -Obj $session -Name "duration_seconds"
            if ($null -ne $durVal) {
                try { $durationSeconds = [int]$durVal } catch { $durationSeconds = $null }
            }

            $endUtc = $null
            if ($null -ne $durationSeconds -and $durationSeconds -ge 0) {
                $endUtc = $startUtc.AddSeconds([int]$durationSeconds)
            } else {
                $endUtc = From-UtcString $endStr
            }
            if (-not $endUtc) { continue }

            if ($endUtc -le $dayStartUtc -or $startUtc -ge $dayEndUtc) { continue }
            $clipStart = $startUtc
            if ($clipStart -lt $dayStartUtc) { $clipStart = $dayStartUtc }
            $clipEnd = $endUtc
            if ($clipEnd -gt $dayEndUtc) { $clipEnd = $dayEndUtc }

            $secs = [int][math]::Floor([math]::Max(0, ($clipEnd - $clipStart).TotalSeconds))
            if ($secs -gt 0) { $timerTotal += $secs }
        }

        $cache[$timerId] = [int]$timerTotal
        $grand += [int]$timerTotal
    }

    $State.runtime.cache_day_local = $DayLocal
    $State.runtime.cache_by_timer = $cache
    $State.runtime.cache_total = [int]$grand
    $State.runtime.cache_dirty = $false
}

function Tick-Timekeeping {
    # No-op: time is tracked using a monotonic anchor (TickCount64) and computed on-demand.
    return
}

function Timer-ElapsedTodaySeconds {
    param($Timer)
    if (-not $Timer) { return 0 }
    Ensure-RuntimeClock
    Ensure-TodayCache

    $timerId = [string](Get-PropValue -Obj $Timer -Name "id")
    if (-not $timerId) { return 0 }

    $base = 0
    $cache = $State.runtime.cache_by_timer
    if ($cache -is [System.Collections.IDictionary] -and $cache.Contains($timerId)) {
        try { $base = [int]$cache[$timerId] } catch { $base = 0 }
    }

    if ($State.data.active_timer_id -and $State.data.active_session_id -and ([string]$State.data.active_timer_id -eq $timerId)) {
        $base += Get-ActiveSessionTodaySeconds
    }
    if ($base -lt 0) { $base = 0 }
    return [int]$base
}

function Total-ForToday {
    Ensure-RuntimeClock
    Ensure-TodayCache

    $total = 0
    try { $total = [int]$State.runtime.cache_total } catch { $total = 0 }
    if ($State.data.active_timer_id -and $State.data.active_session_id) {
        $total += Get-ActiveSessionTodaySeconds
    }
    if ($total -lt 0) { $total = 0 }
    return [int]$total
}

function Clear-ActiveAnchor {
    Ensure-RuntimeClock
    $State.runtime.active_anchor = $null
    $State.runtime.last_active_elapsed_s = 0
    $State.runtime.last_active_timer_id = ""
    $State.runtime.last_active_session_id = ""
}

function Start-ActiveAnchor {
    param(
        [Parameter(Mandatory=$true)][string]$TimerId,
        [Parameter(Mandatory=$true)][string]$SessionId,
        [int]$BaseSeconds = 0
    )
    Ensure-RuntimeClock
    if ($BaseSeconds -lt 0) { $BaseSeconds = 0 }
    $State.runtime.active_anchor = [ordered]@{
        timer_id = $TimerId
        session_id = $SessionId
        start_mono_ms = (Mono-NowMs)
        base_seconds = [int]$BaseSeconds
    }
}

function Ensure-ActiveAnchor {
    Ensure-RuntimeClock
    if (-not $State.data.active_timer_id -or -not $State.data.active_session_id) { return $null }
    $timer = Get-ActiveTimer
    $session = Get-ActiveSession
    if (-not $timer -or -not $session) { return $null }

    $timerId = [string](Get-PropValue -Obj $timer -Name "id")
    $sessionId = [string](Get-PropValue -Obj $session -Name "id")
    if (-not $timerId -or -not $sessionId) { return $null }

    $anchor = $State.runtime.active_anchor
    if ($anchor -and $anchor.timer_id -eq $timerId -and $anchor.session_id -eq $sessionId) {
        return $anchor
    }

    $base = 0
    $durVal = Get-PropValue -Obj $session -Name "duration_seconds"
    if ($null -ne $durVal) {
        try { $base = [int]$durVal } catch { $base = 0 }
    }
    if ($base -lt 0) { $base = 0 }
    if ($base -eq 0) {
        $startUtc = From-UtcString (Get-PropValue -Obj $session -Name "start")
        if ($startUtc) {
            $base = [int][math]::Max(0, [math]::Floor(((Utc-Now) - $startUtc).TotalSeconds))
        }
    }

    Start-ActiveAnchor -TimerId $timerId -SessionId $sessionId -BaseSeconds $base
    return $State.runtime.active_anchor
}

function Get-ActiveSessionElapsedSeconds {
    Ensure-RuntimeClock
    if (-not $State.data.active_timer_id -or -not $State.data.active_session_id) { return 0 }
    $anchor = Ensure-ActiveAnchor
    if (-not $anchor) { return 0 }

    $timerId = [string]$anchor.timer_id
    $sessionId = [string]$anchor.session_id
    if ($State.runtime.last_active_timer_id -ne $timerId -or $State.runtime.last_active_session_id -ne $sessionId) {
        # New session (or anchor replaced): reset monotonic display clamp.
        $State.runtime.last_active_timer_id = $timerId
        $State.runtime.last_active_session_id = $sessionId
        $State.runtime.last_active_elapsed_s = 0
    }

    $nowMs = Mono-NowMs
    $startMs = [long]$anchor.start_mono_ms
    $deltaMs = $nowMs - $startMs
    if ($deltaMs -lt 0) { $deltaMs = 0 }
    $deltaSeconds = [int]($deltaMs / 1000)
    $total = [int]$anchor.base_seconds + $deltaSeconds
    if ($total -lt 0) { $total = 0 }

    # Hard clamp: never allow the displayed elapsed time to go backwards.
    $prev = 0
    try { $prev = [int]$State.runtime.last_active_elapsed_s } catch { $prev = 0 }
    if ($total -lt $prev) { $total = $prev } else { $State.runtime.last_active_elapsed_s = [int]$total }
    return $total
}

function Get-ActiveSessionTodaySeconds {
    Ensure-RuntimeClock
    if (-not $State.data.active_timer_id -or -not $State.data.active_session_id) { return 0 }

    $session = Get-ActiveSession
    if (-not $session) { return 0 }

    $startUtc = From-UtcString (Get-PropValue -Obj $session -Name "start")
    if (-not $startUtc) { return 0 }

    $dayLocal = (Get-Date).Date
    $dayStartUtc = $dayLocal.ToUniversalTime()
    $dayEndUtc = $dayLocal.AddDays(1).ToUniversalTime()

    $elapsedSeconds = Get-ActiveSessionElapsedSeconds
    $endUtc = $startUtc.AddSeconds([int]$elapsedSeconds)

    if ($endUtc -le $dayStartUtc -or $startUtc -ge $dayEndUtc) { return 0 }
    $clipStart = $startUtc
    if ($clipStart -lt $dayStartUtc) { $clipStart = $dayStartUtc }
    $clipEnd = $endUtc
    if ($clipEnd -gt $dayEndUtc) { $clipEnd = $dayEndUtc }

    $secs = [int][math]::Floor([math]::Max(0, ($clipEnd - $clipStart).TotalSeconds))
    if ($secs -lt 0) { $secs = 0 }
    return [int]$secs
}

function Format-Hms {
    param([int]$TotalSeconds)
    if ($TotalSeconds -lt 0) { $TotalSeconds = 0 }

    # Use TimeSpan to avoid any surprising integer math/conversion edge cases in PowerShell.
    $ts = [TimeSpan]::FromSeconds([double]$TotalSeconds)
    $h = [int]$ts.TotalHours
    return ("{0:d2}:{1:d2}:{2:d2}" -f $h, $ts.Minutes, $ts.Seconds)
}

function Desktop-TargetDir {
    $override = $State.settings.desktop_path_override
    if ($override -and -not [string]::IsNullOrWhiteSpace($override)) {
        if (-not (Test-Path $override)) {
            New-Item -ItemType Directory -Path $override -Force | Out-Null
        }
        return $override
    }

    # Default export location (instead of Desktop).
    $path = "C:\\MyTime"
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
    return $path
}

function Write-LogLine {
    param(
        [string]$Label,
        [int]$DurationSeconds,
        [int]$SessionSeconds,
        [string]$Note,
        [DateTime]$EndedAtUtc
    )
    $targetDir = Desktop-TargetDir
    $localEnd = $EndedAtUtc.ToLocalTime()
    $dateStamp = $localEnd.ToString("yyyy-MM-dd")
    $ts = $localEnd.ToString("yyyy-MM-dd HH:mm")
    $sessionText = ""
    if ($SessionSeconds -gt 0) {
        $sessionText = " | Session: {0}" -f (Format-Hms -TotalSeconds $SessionSeconds)
    }
    $line = "[{0}] {1} | Total: {2}{3} | Note: {4}`n" -f $ts, $Label, (Format-Hms -TotalSeconds $DurationSeconds), $sessionText, $Note.Trim()

    if ($State.settings.write_daily_desktop_log) {
        $daily = Join-Path $targetDir ("TicketTimeLog_{0}.txt" -f $dateStamp)
        Add-Content -Path $daily -Value $line
    }
    if ($State.settings.write_per_ticket_desktop_log) {
        $safeLabel = Sanitize-Filename $Label
        $per = Join-Path $targetDir ("{0}_{1}.txt" -f $safeLabel, $dateStamp)
        Add-Content -Path $per -Value $line
    }
}

function Log-StoppedSession {
    param($Summary, [int]$TotalSeconds, [string]$Note)
    $endUtc = From-UtcString $Summary.end
    $sessionSeconds = [int]$Summary.duration_seconds
    Write-LogLine -Label $Summary.label -DurationSeconds $TotalSeconds -SessionSeconds $sessionSeconds -Note $Note -EndedAtUtc $endUtc

    $timer = Get-TimerById -TimerId $Summary.timer_id
    if ($timer) {
        $session = $timer.sessions | Where-Object { $_.id -eq $Summary.session_id } | Select-Object -First 1
        if ($session) {
            Set-PropValue -Obj $session -Name "note" -Value $Note
            Set-PropValue -Obj $session -Name "exported" -Value $true
        }
    }
    Save-Data
}

function Save-PendingSessions {
    $pending = @()
    foreach ($timer in $State.data.timers) {
        foreach ($session in $timer.sessions) {
            if (-not $session.exported -and $session.end) {
                $endUtc = From-UtcString $session.end
                $durationSeconds = $null
                $durVal = Get-PropValue -Obj $session -Name "duration_seconds"
                if ($null -ne $durVal) {
                    try { $durationSeconds = [int]$durVal } catch { $durationSeconds = $null }
                }
                if ($null -eq $durationSeconds -or $durationSeconds -lt 0) {
                    $startUtc = From-UtcString $session.start
                    $durationSeconds = 0
                    if ($startUtc -and $endUtc) {
                        $durationSeconds = [int][math]::Max(0, ($endUtc - $startUtc).TotalSeconds)
                    }
                    Set-PropValue -Obj $session -Name "duration_seconds" -Value $durationSeconds
                }
                $pending += [ordered]@{
                    timer = $timer
                    session = $session
                    label = $timer.label
                    duration_seconds = $durationSeconds
                    note = $session.note
                    end = $endUtc
                }
            }
        }
    }

    foreach ($item in $pending) {
                Write-LogLine -Label $item.label -DurationSeconds $item.duration_seconds -SessionSeconds 0 -Note $item.note -EndedAtUtc $item.end
        Set-PropValue -Obj $item.session -Name "exported" -Value $true
    }
    if ($pending.Count -gt 0) {
        Invalidate-TodayCache
        Save-Data
    }
}

function Stop-And-SaveAll {
    $summary = Close-ActiveSession -At (Utc-Now)
    if ($summary) {
        $timer = Get-TimerById -TimerId $summary.timer_id
        $totalSeconds = if ($timer) { Timer-TotalSeconds -Timer $timer } else { [int]$summary.duration_seconds }
        Log-StoppedSession -Summary $summary -TotalSeconds $totalSeconds -Note ""
    }
    Save-PendingSessions
}

function Handle-Recovery {
    if (-not $State.runtime.recovery_prompt) { return }
    $prompt = $State.runtime.recovery_prompt
    $label = $prompt.label
    $choice = New-SelectDialog -Title "MyTime Recovery" -Prompt ("Resume previous timer: {0}?" -f $label) -Items @("Resume", "Stop")
    $timer = Get-TimerById -TimerId $prompt.timer_id
    if ($timer) {
        $session = $timer.sessions | Where-Object { $_.id -eq $prompt.session_id } | Select-Object -First 1
        if ($session) {
            $session.end = To-UtcString $prompt.startup_at
            if (-not ($session.PSObject.Properties.Name -contains "duration_seconds")) {
                Set-PropValue -Obj $session -Name "duration_seconds" -Value 0
            }
            $startUtc = From-UtcString $session.start
            $endUtc = From-UtcString $session.end
            if ($startUtc -and $endUtc) {
                Set-PropValue -Obj $session -Name "duration_seconds" -Value ([int][math]::Max(0, ($endUtc - $startUtc).TotalSeconds))
            }
            if (-not $session.note) {
                Set-PropValue -Obj $session -Name "note" -Value "Recovered after app restart"
            } else {
                Set-PropValue -Obj $session -Name "note" -Value ($session.note + " | Recovered after app restart")
            }
        }
    }
    $State.data.active_timer_id = $null
    $State.data.active_session_id = $null
    if ($choice -eq "Resume") {
        Start-TimerInternal -TimerId $prompt.timer_id
    } else {
        Save-Data
    }
    $State.runtime.recovery_prompt = $null
    Rebuild-TodayCache -DayLocal (Get-Date).Date
}

function Load-State {
    $State.data = Read-JsonOrDefault -Path $dataPath -DefaultValue (New-DefaultData)
    $State.settings = Read-JsonOrDefault -Path $settingsPath -DefaultValue (New-DefaultSettings)

    if (-not $State.data.timers) { $State.data.timers = @() }
    if ($null -eq $State.settings.write_daily_desktop_log) { $State.settings.write_daily_desktop_log = $true }
    if ($null -eq $State.settings.write_per_ticket_desktop_log) { $State.settings.write_per_ticket_desktop_log = $true }
    if (-not $State.settings.desktop_path_override) { $State.settings.desktop_path_override = "" }
    if (-not ($State.settings.PSObject.Properties.Name -contains "floating_font_scale")) {
        $State.settings | Add-Member -NotePropertyName floating_font_scale -NotePropertyValue 1.0
    }
    $fontScale = Get-PropValue -Obj $State.settings -Name "floating_font_scale"
    if ($null -eq $fontScale -or [double]$fontScale -le 0) { $State.settings.floating_font_scale = 1.0 }

    # Normalize timers loaded from JSON to prevent missing properties
    $normalizedTimers = @()
    foreach ($t in $State.data.timers) {
        if (-not ($t -is [psobject])) { continue }
        if (-not $t.PSObject.Properties["id"]) { continue }
        if (-not $t.PSObject.Properties["label"]) { $t | Add-Member -NotePropertyName label -NotePropertyValue "" }
        if (-not $t.PSObject.Properties["created_at"]) { $t | Add-Member -NotePropertyName created_at -NotePropertyValue (To-UtcString (Utc-Now)) }
        if (-not $t.PSObject.Properties["sessions"]) { $t | Add-Member -NotePropertyName sessions -NotePropertyValue @() }
        if (-not ($t.sessions -is [System.Collections.IEnumerable])) { $t.sessions = @() }

        # Normalize sessions (older data may be missing properties).
        $normalizedSessions = @()
        foreach ($s in @($t.sessions | Where-Object { $_ })) {
            if (-not ($s -is [psobject])) { continue }
            if (-not $s.PSObject.Properties["id"]) { $s | Add-Member -NotePropertyName id -NotePropertyValue (New-Id) }
            if (-not $s.PSObject.Properties["start"]) { continue }
            if (-not $s.PSObject.Properties["end"]) { $s | Add-Member -NotePropertyName end -NotePropertyValue $null }
            if (-not $s.PSObject.Properties["note"]) { $s | Add-Member -NotePropertyName note -NotePropertyValue "" }
            if (-not $s.PSObject.Properties["exported"]) { $s | Add-Member -NotePropertyName exported -NotePropertyValue $false }
            if (-not $s.PSObject.Properties["duration_seconds"]) {
                $dur = 0
                if ($s.end) {
                    $startUtc = From-UtcString $s.start
                    $endUtc = From-UtcString $s.end
                    if ($startUtc -and $endUtc) {
                        $dur = [int][math]::Max(0, ($endUtc - $startUtc).TotalSeconds)
                    }
                }
                $s | Add-Member -NotePropertyName duration_seconds -NotePropertyValue $dur
            }
            $normalizedSessions += $s
        }
        $t.sessions = $normalizedSessions

        $normalizedTimers += $t
    }
    $State.data.timers = $normalizedTimers

    $startupNow = Utc-Now
    if ($State.data.active_timer_id -and $State.data.active_session_id) {
        $timer = Get-TimerById -TimerId $State.data.active_timer_id
        if ($timer) {
            $session = $timer.sessions | Where-Object { $_.id -eq $State.data.active_session_id } | Select-Object -First 1
            if ($session -and -not $session.end) {
                $State.runtime.recovery_prompt = [ordered]@{
                    timer_id = $timer.id
                    session_id = $session.id
                    label = $timer.label
                    startup_at = $startupNow
                }
                $State.data.active_timer_id = $null
                $State.data.active_session_id = $null
            }
        }
    }

    Save-All
    Ensure-RuntimeClock
    Rebuild-TodayCache -DayLocal (Get-Date).Date
}

function Update-TrayTooltip {
    param([System.Windows.Forms.NotifyIcon]$Tray)
    $activeLabel = $null
    $activeTimer = Get-ActiveTimer
    if ($activeTimer) { $activeLabel = $activeTimer.label }
    $activeElapsed = Active-ElapsedSeconds
    $activeTotal = 0
    if ($activeTimer) { $activeTotal = Timer-TotalSeconds -Timer $activeTimer }
    $symbol = "Stopped"
    if ($activeTimer) {
        if ($State.data.active_session_id) {
            $symbol = "Running"
        } else {
            $symbol = "Paused"
        }
    }
    $label = if ($activeLabel) { $activeLabel } else { "No active timer" }
    $tooltip = "{0} {1} {2}" -f $symbol, $label, (Format-Hms -TotalSeconds $activeElapsed)
    if ($activeTimer) {
        $tooltip = "{0} | Total {1}" -f $tooltip, (Format-Hms -TotalSeconds $activeTotal)
    }

    # Trace what we actually assign near the reported jump point.
    try {
        $mod = [int]($activeElapsed % 60)
        if ($mod -ge 28 -and $mod -le 32) {
            Perf-Note -Line ("[{0}] tray_assign elapsed_s={1} total_s={2} text='{3}'" -f (Get-Date).ToString("s"), [int]$activeElapsed, [int]$activeTotal, $tooltip)
        }
    } catch {
    }

    if ($tooltip.Length -gt 127) { $tooltip = $tooltip.Substring(0, 127) }
    $Tray.Text = $tooltip
}

function Update-FloatingWindow {
    param([System.Windows.Forms.Form]$Window, $Panel)
    if (-not $Window -or -not $Panel) { return }
    if ($Window.IsDisposed) { return }

    $timers = @($State.data.timers | Where-Object { $_ })
    if (-not (Get-Variable -Name floatingControls -Scope Script -ErrorAction SilentlyContinue) -or -not ($script:floatingControls -is [System.Collections.IDictionary])) { $script:floatingControls = @{} }
    if (-not (Get-Variable -Name floatingOrder -Scope Script -ErrorAction SilentlyContinue)) { $script:floatingOrder = @() }
    if (-not (Get-Variable -Name floatingLastSeconds -Scope Script -ErrorAction SilentlyContinue) -or -not ($script:floatingLastSeconds -is [System.Collections.IDictionary])) { $script:floatingLastSeconds = @{} }
    if (-not (Get-Variable -Name floatingLastLabels -Scope Script -ErrorAction SilentlyContinue) -or -not ($script:floatingLastLabels -is [System.Collections.IDictionary])) { $script:floatingLastLabels = @{} }
    $scale = [double]$State.settings.floating_font_scale
    if ($scale -le 0) { $scale = 1.0 }
    $needsFontRefresh = $false
    if (-not (Get-Variable -Name floatingFonts -Scope Script -ErrorAction SilentlyContinue) -or -not (Get-Variable -Name floatingFontsScale -Scope Script -ErrorAction SilentlyContinue) -or ([math]::Abs([double]$script:floatingFontsScale - $scale) -gt 0.001)) {
        $monoFontName = "Cascadia Mono"
        try {
            $null = New-Object System.Drawing.Font($monoFontName, 9)
        } catch {
            $monoFontName = "Segoe UI"
        }
        $titleSize = [math]::Round(10 * $scale, 1)
        $timeSize = [math]::Round(13 * $scale, 1)
        $script:floatingFonts = @{
            title = New-Object System.Drawing.Font($monoFontName, $titleSize, [System.Drawing.FontStyle]::Bold)
            time = New-Object System.Drawing.Font($monoFontName, $timeSize, [System.Drawing.FontStyle]::Bold)
        }
        $script:floatingFontsScale = $scale
        $needsFontRefresh = $true
    }

    $ids = @($timers | ForEach-Object { [string](Get-PropValue -Obj $_ -Name "id") })
    $known = @($script:floatingControls.Keys)
    $needsRebuild = ($ids.Count -ne $known.Count) -or (@($ids | Where-Object { $known -notcontains $_ }).Count -gt 0)

    if ($needsRebuild) {
        try { $Window.SuspendLayout() } catch {}
        try { $Panel.SuspendLayout() } catch {}
        try { foreach ($control in @($Panel.Controls)) { $control.Dispose() } } catch {}
        try { $Panel.Controls.Clear() } catch {}
        $script:floatingControls = @{}
        $script:floatingOrder = @()
        $script:floatingLastSeconds = @{}
        $script:floatingLastLabels = @{}
        $palette = @(
            [System.Drawing.Color]::FromArgb(255, 99, 71),
            [System.Drawing.Color]::FromArgb(100, 149, 237),
            [System.Drawing.Color]::FromArgb(60, 179, 113),
            [System.Drawing.Color]::FromArgb(255, 165, 0),
            [System.Drawing.Color]::FromArgb(186, 85, 211),
            [System.Drawing.Color]::FromArgb(72, 209, 204),
            [System.Drawing.Color]::FromArgb(238, 130, 238),
            [System.Drawing.Color]::FromArgb(255, 215, 0)
        )
        $index = 0

        # Compute a width that fits inside the FlowLayoutPanel's padded display area,
        # accounting for the per-block left/right margin (so right padding doesn't look "off").
        $avail = 0
        try { $avail = [int]$Panel.DisplayRectangle.Width } catch { $avail = 0 }
        if ($avail -le 0) {
            try { $avail = [int]($Panel.ClientSize.Width) } catch { $avail = 0 }
        }
        if ($avail -le 0) {
            $avail = [int]($Window.ClientSize.Width - 24)
        }
        $blockMarginLR = 12 # 6px left + 6px right (matches the margin we set below)
        $blockWidth = [math]::Max(200, ($avail - $blockMarginLR))
        foreach ($t in $timers) {
            $timerId = [string](Get-PropValue -Obj $t -Name "id")
            if (-not $timerId) { continue }
            $labelText = [string](Get-PropValue -Obj $t -Name "label")
            if ([string]::IsNullOrWhiteSpace($labelText)) { $labelText = "Untitled" }
            $color = $palette[$index % $palette.Count]
            $index += 1

            $block = New-Object System.Windows.Forms.Panel
            $block.Height = 54
            $block.Width = $blockWidth
            $block.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
            $block.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
            $block.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
            Enable-DoubleBuffering -Control $block
            Set-RoundedCorners -Control $block -Radius 5

            $title = New-Object System.Windows.Forms.Label
            $title.AutoSize = $false
            $title.Dock = "Top"
            $title.Height = 22
            $title.TextAlign = "MiddleLeft"
            $title.Font = $script:floatingFonts.title
            $title.ForeColor = $color
            $title.BackColor = $block.BackColor
            $title.Text = $labelText

            $time = New-Object System.Windows.Forms.Label
            $time.AutoSize = $false
            $time.Dock = "Fill"
            $time.TextAlign = "MiddleLeft"
            $time.Font = $script:floatingFonts.time
            $time.ForeColor = [System.Drawing.Color]::White
            $time.BackColor = $block.BackColor

            $block.Controls.Add($time)
            $block.Controls.Add($title)
            $Panel.Controls.Add($block) | Out-Null

            $script:floatingControls[$timerId] = @{
                block = $block
                title = $title
                time = $time
            }
            $script:floatingOrder += $timerId
        }

        # Height changes only when the timer list changes; avoid doing this work on every tick.
        $panelPadding = $Panel.Padding
        $panelHeight = 0
        foreach ($id in $script:floatingOrder) {
            $controls = $script:floatingControls[$id]
            if ($controls) {
                $panelHeight += $controls.block.Height + $controls.block.Margin.Top + $controls.block.Margin.Bottom
            }
        }
        $panelHeight += $panelPadding.Top + $panelPadding.Bottom + 4
        $targetHeight = [int]$panelHeight
        if ($Window.MaximumSize.Height -gt 0) { $targetHeight = [int][math]::Min($targetHeight, $Window.MaximumSize.Height) }
        if ($Window.MinimumSize.Height -gt 0) { $targetHeight = [int][math]::Max($targetHeight, $Window.MinimumSize.Height) }
        $Window.Height = $targetHeight

        Ensure-RuntimeClock
        $State.runtime.ui_times_dirty = $true

        try { $Panel.ResumeLayout($true) } catch {}
        try { $Window.ResumeLayout($true) } catch {}
    }

    if (-not $timers -or $timers.Count -eq 0) {
        foreach ($control in @($Panel.Controls)) { $control.Dispose() }
        $Panel.Controls.Clear()
        $script:floatingControls = @{}
        $script:floatingOrder = @()
        $script:floatingLastSeconds = @{}
        $script:floatingLastLabels = @{}
        $empty = New-Object System.Windows.Forms.Label
        $empty.AutoSize = $true
        $empty.Text = "No timers"
        $empty.ForeColor = [System.Drawing.Color]::White
        $empty.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
        $Panel.Controls.Add($empty) | Out-Null
        return
    }

    Ensure-RuntimeClock
    $forceAll = $needsFontRefresh
    if (-not $forceAll) {
        try { $forceAll = [bool]$State.runtime.ui_times_dirty } catch { $forceAll = $false }
    }
    if ($forceAll) { $State.runtime.ui_times_dirty = $false }

    $activeId = [string]$State.data.active_timer_id
    foreach ($t in $timers) {
        $timerId = [string](Get-PropValue -Obj $t -Name "id")
        if (-not $timerId) { continue }
        $controls = $script:floatingControls[$timerId]
        if (-not $controls) { continue }

        $labelText = [string](Get-PropValue -Obj $t -Name "label")
        if ([string]::IsNullOrWhiteSpace($labelText)) { $labelText = "Untitled" }

        $prevLabel = $null
        if ($script:floatingLastLabels.Contains($timerId)) { $prevLabel = [string]$script:floatingLastLabels[$timerId] }
        if ($forceAll -or $prevLabel -ne $labelText) {
            $controls.title.Text = $labelText
            $script:floatingLastLabels[$timerId] = $labelText
        }

        if ($needsFontRefresh) {
            $controls.title.Font = $script:floatingFonts.title
            $controls.time.Font = $script:floatingFonts.time
        }

        # Only the active timer changes every tick; update others only when forced (cache changes, rebuilds, etc.).
        if (-not $forceAll -and $timerId -ne $activeId) { continue }

        $elapsedSeconds = [int](Timer-TotalSeconds -Timer $t)
        $prevSeconds = $null
        if ($script:floatingLastSeconds.Contains($timerId)) { $prevSeconds = $script:floatingLastSeconds[$timerId] }
        if ($null -eq $prevSeconds -or [int]$prevSeconds -ne $elapsedSeconds) {
            $text = (Format-Hms -TotalSeconds $elapsedSeconds)
            $controls.time.Text = $text
            $script:floatingLastSeconds[$timerId] = [int]$elapsedSeconds

            # Trace what we actually assign for the active timer near the reported jump point.
            if ($timerId -eq $activeId) {
                try {
                    $mod = [int]($elapsedSeconds % 60)
                    if ($mod -ge 28 -and $mod -le 32) {
                        Perf-Note -Line ("[{0}] float_assign elapsed_s={1} text='{2}'" -f (Get-Date).ToString("s"), [int]$elapsedSeconds, $text)
                    }
                } catch {
                }
            }
        }
    }
}

function Pause-ActiveSession {
    $timer = Get-ActiveTimer
    if (-not $timer) { return }
    if (-not $State.data.active_session_id) { return }

    $session = Get-ActiveSession
    if (-not $session) {
        $State.data.active_session_id = $null
        Clear-ActiveAnchor
        Invalidate-TodayCache
        Save-Data
        return
    }

    $durationSeconds = Get-ActiveSessionElapsedSeconds
    Set-PropValue -Obj $session -Name "duration_seconds" -Value ([int]$durationSeconds)
    if (-not (Get-PropValue -Obj $session -Name "end")) {
        Set-PropValue -Obj $session -Name "end" -Value (To-UtcString (Utc-Now))
    }

    $State.data.active_session_id = $null
    Clear-ActiveAnchor
    Invalidate-TodayCache
    Save-Data
}

function Start-OrPause {
    $activeTimer = Get-ActiveTimer
    if ($activeTimer) {
        if ($State.data.active_session_id) {
            Pause-ActiveSession
            return
        }
        Start-TimerInternal -TimerId $activeTimer.id
        return
    }
    $id = Most-RecentTimerId
    if ($id) {
        Start-TimerInternal -TimerId $id
    } else {
        $label = New-InputDialog -Title "New Timer" -Prompt "Ticket / label:" -DefaultValue ""
        if ($label) { Create-OrStartTimer -Label $label }
    }
}

function New-Timer {
    $label = New-InputDialog -Title "New Timer" -Prompt "Ticket / label:" -DefaultValue ""
    if ($label) { Create-OrStartTimer -Label $label }
}

function Switch-Timer {
    $choices = @()
    $choiceMap = @{}
    foreach ($t in ($State.data.timers | Where-Object { $_ })) {
        $label = [string](Get-PropValue -Obj $t -Name "label")
        if ([string]::IsNullOrWhiteSpace($label)) {
            $idVal = Get-PropValue -Obj $t -Name "id"
            $shortId = if ($idVal) { $idVal.ToString().Substring(0, 8) } else { "unknown" }
            $label = "Untitled $shortId"
        }
        $idVal = Get-PropValue -Obj $t -Name "id"
        $idText = if ($idVal) { [string]$idVal } else { "unknown" }
        $elapsed = Format-Hms -TotalSeconds (Timer-TotalSeconds -Timer $t)
        $display = "{0} - {1} - {2}" -f $label, $elapsed, $idText
        $choices += $display
        $choiceMap[$display] = $idVal
    }

    if (-not $choices -or $choices.Count -eq 0) {
        New-Timer
        return
    }
    $selected = New-SelectDialog -Title "Switch Timer" -Prompt "Select timer:" -Items $choices
    if (-not $selected) { return }
    $timerId = $choiceMap[$selected]
    if ($timerId) { Start-TimerInternal -TimerId $timerId }
}

function Stop-Timer {
    $summary = $null
    if ($State.data.active_session_id) {
        $summary = Close-ActiveSession -At (Utc-Now)
    } else {
        $timer = Get-ActiveTimer
        if ($timer) {
            $latest = $timer.sessions | Sort-Object -Property @{
                Expression = { if ($_.end) { From-UtcString $_.end } else { From-UtcString $_.start } }
                Descending = $true
            } | Select-Object -First 1
            if ($latest) {
                if (-not (Get-PropValue -Obj $latest -Name "end")) {
                    Set-PropValue -Obj $latest -Name "end" -Value (To-UtcString (Utc-Now))
                }
                $durationSeconds = $null
                $durVal = Get-PropValue -Obj $latest -Name "duration_seconds"
                if ($null -ne $durVal) {
                    try { $durationSeconds = [int]$durVal } catch { $durationSeconds = $null }
                }
                if ($null -eq $durationSeconds -or $durationSeconds -lt 0) {
                    $start = From-UtcString $latest.start
                    $end = From-UtcString $latest.end
                    if ($start -and $end) {
                        $durationSeconds = [int][math]::Max(0, ($end - $start).TotalSeconds)
                        Set-PropValue -Obj $latest -Name "duration_seconds" -Value $durationSeconds
                    } else {
                        $durationSeconds = 0
                    }
                }
                $summary = [ordered]@{
                    timer_id = $timer.id
                    session_id = $latest.id
                    label = $timer.label
                    duration_seconds = [int]$durationSeconds
                    start = $latest.start
                    end = $latest.end
                }
            }
        }
    }
    if ($summary) {
        $timerForTotal = Get-TimerById -TimerId $summary.timer_id
        $totalSeconds = if ($timerForTotal) { Timer-TotalSeconds -Timer $timerForTotal } else { [int]$summary.duration_seconds }
        $note = New-NoteDialog -Label $summary.label -Duration (Format-Hms -TotalSeconds $totalSeconds) -SessionDuration (Format-Hms -TotalSeconds $summary.duration_seconds)
        if ($null -eq $note) { $note = "" }
        Log-StoppedSession -Summary $summary -TotalSeconds $totalSeconds -Note $note
    }

    $removeId = $null
    if ($summary) {
        $removeId = $summary.timer_id
    } else {
        $removeId = $State.data.active_timer_id
    }
    if ($removeId) {
        $State.data.timers = @(
            $State.data.timers | Where-Object { $_ -and (Get-PropValue -Obj $_ -Name "id") -ne $removeId }
        )
    }
    $State.data.active_timer_id = $null
    $State.data.active_session_id = $null
    if ($State.data.timers.Count -eq 1) {
        $remaining = $State.data.timers[0]
        $remainingId = Get-PropValue -Obj $remaining -Name "id"
        if ($remainingId) { $State.data.active_timer_id = $remainingId }
    }
    Invalidate-TodayCache
    Save-Data
}

function Reset-AllData {
    $confirm = New-ConfirmDialog -Title "Reset All Data" -Message "Clear all timers and totals? This cannot be undone."
    if (-not $confirm) { return }
    $State.data = New-DefaultData
    $State.runtime.last_autosave_at = $null
    $State.runtime.recovery_prompt = $null
    $State.runtime.mono_epoch_ms = Mono-NowMs
    $State.runtime.active_anchor = $null
    $State.runtime.cache_day_local = (Get-Date).Date
    $State.runtime.cache_by_timer = @{}
    $State.runtime.cache_total = 0
    $State.runtime.cache_dirty = $false
    $State.runtime.last_autosave_mono_s = $null
    Save-Data
}

function Maybe-Autosave {
    Ensure-RuntimeClock

    # Data doesn't change while a session is running (elapsed time is computed from start + monotonic clock).
    # Avoid disk writes on the UI timer tick: they can stall WinForms timers and cause visible "freeze then jump".
    if ($State.data.active_timer_id -and $State.data.active_session_id) { return }
    if (-not $AutoSaveSeconds -or $AutoSaveSeconds -le 0) { return }

    $nowS = Mono-NowSeconds
    if (-not $State.runtime.Contains("last_autosave_mono_s") -or $null -eq $State.runtime.last_autosave_mono_s) {
        $State.runtime.last_autosave_mono_s = $nowS
        return
    }
    $last = [long]$State.runtime.last_autosave_mono_s
    if ($nowS -ge ($last + $AutoSaveSeconds)) {
        # Autosave should be cheap and never stall the UI; export/logging is handled on Stop or manually.
        Save-Data
        $State.runtime.last_autosave_mono_s = $nowS
    }
}

Load-State
Perf-Init

[System.Windows.Forms.Application]::EnableVisualStyles()

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$trayItemStartPause = $trayMenu.Items.Add("Start/Pause")
$trayItemNew = $trayMenu.Items.Add("New Timer")
$trayItemSwitch = $trayMenu.Items.Add("Switch Timer")
$trayItemStop = $trayMenu.Items.Add("Stop Timer")
$trayMenu.Items.Add("-") | Out-Null
$trayItemFont = New-Object System.Windows.Forms.ToolStripMenuItem("Floating Font Size")
$trayItemFontSmall = New-Object System.Windows.Forms.ToolStripMenuItem("Small (Default)")
$trayItemFontMed = New-Object System.Windows.Forms.ToolStripMenuItem("Medium (+20%)")
$trayItemFontLarge = New-Object System.Windows.Forms.ToolStripMenuItem("Large (+40%)")
$trayItemFont.DropDownItems.AddRange(@($trayItemFontSmall, $trayItemFontMed, $trayItemFontLarge))
$trayMenu.Items.Add($trayItemFont) | Out-Null
$trayItemResetAll = $trayMenu.Items.Add("Reset All Data")
$trayItemQuit = $trayMenu.Items.Add("Quit")

$tray = New-Object System.Windows.Forms.NotifyIcon
$tray.Icon = Load-AppIcon
$tray.Visible = $true
$tray.ContextMenuStrip = $trayMenu
$tray.Text = "MyTime"

$trayItemStartPause.Add_Click({ Invoke-Safe { Start-OrPause } })
$trayItemNew.Add_Click({ Invoke-Safe { New-Timer } })
$trayItemSwitch.Add_Click({ Invoke-Safe { Switch-Timer } })
$trayItemStop.Add_Click({ Invoke-Safe { Stop-Timer } })
$trayItemFontSmall.Add_Click({
    Invoke-Safe {
        $State.settings.floating_font_scale = 1.0
        Save-Settings
    }
})
$trayItemFontMed.Add_Click({
    Invoke-Safe {
        $State.settings.floating_font_scale = 1.2
        Save-Settings
    }
})
$trayItemFontLarge.Add_Click({
    Invoke-Safe {
        $State.settings.floating_font_scale = 1.4
        Save-Settings
    }
})
$trayItemResetAll.Add_Click({ Invoke-Safe { Reset-AllData } })
$trayItemQuit.Add_Click({
    Stop-And-SaveAll
    Perf-Flush
    $tray.Visible = $false
    $tray.Dispose()
    if ($floatingWindow -and -not $floatingWindow.IsDisposed) {
        $floatingWindow.Close()
        $floatingWindow.Dispose()
    }
    $hotKeyForm.Close()
})

$hotKeyForm = New-Object HotKeyWindow
$hotKeyForm.ShowInTaskbar = $false
$hotKeyForm.WindowState = "Minimized"
$hotKeyForm.FormBorderStyle = "FixedToolWindow"
$hotKeyForm.Opacity = 0
$hotKeyForm.Icon = Load-AppIcon

$MOD_ALT = 0x0001
$MOD_CONTROL = 0x0002

$HOTKEY_STARTPAUSE = 1
$HOTKEY_NEW = 2
$HOTKEY_SWITCH = 3
$HOTKEY_STOP = 4

$hotKeyForm.add_HotKeyPressed({
    param($id)
    Invoke-Safe {
        switch ($id) {
            $HOTKEY_STARTPAUSE { Start-OrPause }
            $HOTKEY_NEW { New-Timer }
            $HOTKEY_SWITCH { Switch-Timer }
            $HOTKEY_STOP { Stop-Timer }
        }
    }
})

$hotKeyForm.Add_Shown({
    $handle = $hotKeyForm.Handle
    [HotKeyWindow]::RegisterHotKey($handle, $HOTKEY_STARTPAUSE, $MOD_CONTROL -bor $MOD_ALT, [int][System.Windows.Forms.Keys]::S) | Out-Null
    [HotKeyWindow]::RegisterHotKey($handle, $HOTKEY_NEW, $MOD_CONTROL -bor $MOD_ALT, [int][System.Windows.Forms.Keys]::N) | Out-Null
    [HotKeyWindow]::RegisterHotKey($handle, $HOTKEY_SWITCH, $MOD_CONTROL -bor $MOD_ALT, [int][System.Windows.Forms.Keys]::X) | Out-Null
    [HotKeyWindow]::RegisterHotKey($handle, $HOTKEY_STOP, $MOD_CONTROL -bor $MOD_ALT, [int][System.Windows.Forms.Keys]::E) | Out-Null

    Handle-Recovery
    Update-TrayTooltip -Tray $tray
})

$hotKeyForm.Add_FormClosing({
    $handle = $hotKeyForm.Handle
    [HotKeyWindow]::UnregisterHotKey($handle, $HOTKEY_STARTPAUSE) | Out-Null
    [HotKeyWindow]::UnregisterHotKey($handle, $HOTKEY_NEW) | Out-Null
    [HotKeyWindow]::UnregisterHotKey($handle, $HOTKEY_SWITCH) | Out-Null
    [HotKeyWindow]::UnregisterHotKey($handle, $HOTKEY_STOP) | Out-Null
})

$floatingWindow = New-Object System.Windows.Forms.Form
$floatingWindow.Text = "MyTime"
$floatingWindow.StartPosition = "Manual"
$floatingWindow.TopMost = $true
$floatingWindow.FormBorderStyle = "None"
$floatingWindow.ShowInTaskbar = $false
$floatingWindow.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$floatingWindow.ForeColor = [System.Drawing.Color]::White
$floatingWindow.Opacity = 0.9
$floatingWindow.Location = New-Object System.Drawing.Point(20, 20)
$floatingWindow.AutoSize = $false
$floatingWindow.AutoSizeMode = "GrowAndShrink"
$floatingWindow.MinimumSize = New-Object System.Drawing.Size(260, 60)
$floatingWindow.MaximumSize = New-Object System.Drawing.Size(260, 2000)
$floatingWindow.Add_Shown({ Invoke-Safe { Set-RoundedCorners -Control $floatingWindow -Radius 5 } })
$floatingWindow.Add_SizeChanged({ Invoke-Safe { Set-RoundedCorners -Control $floatingWindow -Radius 5 } })

$floatingInner = New-Object System.Windows.Forms.Panel
$floatingInner.Dock = "Fill"
$floatingInner.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
$floatingInner.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
$floatingWindow.Controls.Add($floatingInner)
Enable-DoubleBuffering -Control $floatingInner
$floatingInner.Add_SizeChanged({ Invoke-Safe { Set-RoundedCorners -Control $floatingInner -Radius 5 } })

$floatingPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$floatingPanel.Dock = "Fill"
$floatingPanel.AutoScroll = $false
$floatingPanel.AutoSize = $false
$floatingPanel.AutoSizeMode = "GrowAndShrink"
$floatingPanel.WrapContents = $false
$floatingPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$floatingPanel.Padding = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
$floatingPanel.BackColor = $floatingInner.BackColor
$floatingInner.Controls.Add($floatingPanel)
Enable-DoubleBuffering -Control $floatingWindow
Enable-DoubleBuffering -Control $floatingPanel

$floatingWindow.Add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        [Win32Drag]::ReleaseCapture() | Out-Null
        [Win32Drag]::SendMessage($floatingWindow.Handle, 0x00A1, [IntPtr]2, [IntPtr]0) | Out-Null
    }
})
$floatingPanel.Add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        [Win32Drag]::ReleaseCapture() | Out-Null
        [Win32Drag]::SendMessage($floatingWindow.Handle, 0x00A1, [IntPtr]2, [IntPtr]0) | Out-Null
    }
})
$floatingInner.Add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        [Win32Drag]::ReleaseCapture() | Out-Null
        [Win32Drag]::SendMessage($floatingWindow.Handle, 0x00A1, [IntPtr]2, [IntPtr]0) | Out-Null
    }
})
Update-FloatingWindow -Window $floatingWindow -Panel $floatingPanel
$floatingWindow.Show()

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    Invoke-Safe {
        Ensure-RuntimeClock
        $tickNowMs = Mono-NowMs
        $lastTickMs = $null
        try { $lastTickMs = [long]$State.runtime.last_ui_tick_ms } catch { $lastTickMs = $null }
        if ($null -ne $lastTickMs) {
            $gap = $tickNowMs - $lastTickMs
            if ($gap -gt 1500) {
                Perf-Note -Line ("[{0}] ui_tick_gap_ms={1}" -f (Get-Date).ToString("s"), [int]$gap)
            }
        }
        $State.runtime.last_ui_tick_ms = [long]$tickNowMs

        # Low-volume trace near the point you see the bug (around :30 each minute) to prove
        # whether computed elapsed seconds are jumping or only the UI is missing updates.
        if ($State.data.active_timer_id -and $State.data.active_session_id) {
            $e = 0
            try { $e = [int](Get-ActiveSessionElapsedSeconds) } catch { $e = 0 }
            $mod = 0
            try { $mod = [int]($e % 60) } catch { $mod = 0 }
            if ($mod -ge 28 -and $mod -le 32) {
                $ae = 0
                $tt = 0
                $cacheTotal = 0
                $cacheBy = $null
                $activeId = [string]$State.data.active_timer_id
                try { $ae = [int](Active-ElapsedSeconds) } catch { $ae = 0 }
                try { $tt = [int](Total-ForToday) } catch { $tt = 0 }
                try { $cacheTotal = [int]$State.runtime.cache_total } catch { $cacheTotal = 0 }
                try { $cacheBy = $State.runtime.cache_by_timer } catch { $cacheBy = $null }
                $baseForActive = 0
                if ($cacheBy -is [System.Collections.IDictionary] -and $activeId -and $cacheBy.Contains($activeId)) {
                    try { $baseForActive = [int]$cacheBy[$activeId] } catch { $baseForActive = 0 }
                }

                Perf-Note -Line ("[{0}] session_elapsed_s={1} active_elapsed_s={2} today_total_s={3} base_active_s={4} cache_total_s={5} mod60={6}" -f (Get-Date).ToString("s"), $e, $ae, $tt, $baseForActive, $cacheTotal, $mod)
            }
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $trayItemFontSmall.Checked = ($State.settings.floating_font_scale -eq 1.0)
        $trayItemFontMed.Checked = ($State.settings.floating_font_scale -eq 1.2)
        $trayItemFontLarge.Checked = ($State.settings.floating_font_scale -eq 1.4)
        $sw.Stop()
        $msMenu = [int]$sw.ElapsedMilliseconds

        $sw.Restart()
        Maybe-Autosave
        $sw.Stop()
        $msAutosave = [int]$sw.ElapsedMilliseconds

        $sw.Restart()
        Update-TrayTooltip -Tray $tray
        $sw.Stop()
        $msTray = [int]$sw.ElapsedMilliseconds

        $sw.Restart()
        Update-FloatingWindow -Window $floatingWindow -Panel $floatingPanel
        $sw.Stop()
        $msFloat = [int]$sw.ElapsedMilliseconds

        $maxMs = [math]::Max([math]::Max($msMenu, $msAutosave), [math]::Max($msTray, $msFloat))
        if ($maxMs -ge 200) {
            Perf-Note -Line ("[{0}] ui_tick_ms menu={1} autosave={2} tray={3} float={4}" -f (Get-Date).ToString("s"), $msMenu, $msAutosave, $msTray, $msFloat)
        }
    }
})
$timer.Start()

[System.Windows.Forms.Application]::Run($hotKeyForm)
