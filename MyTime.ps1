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
    return (Get-Date).ToUniversalTime()
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
            $line = "[{0}] {1}`r`n" -f (Get-Date).ToString("s"), $msg
            Add-Content -Path $logPath -Value $line
        } catch {
        }
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
            Save-Data
        }
        return $null
    }

    $session = Get-ActiveSession
    if (-not $session) {
        Clear-ActiveTracking
        Save-Data
        return $null
    }

    if (-not $session.end) {
        $session.end = To-UtcString $At
    }

    $start = From-UtcString $session.start
    $end = From-UtcString $session.end
    $durationSeconds = [math]::Max(0, ($end - $start).TotalSeconds)

    $summary = [ordered]@{
        timer_id = $timer.id
        session_id = $session.id
        label = $timer.label
        duration_seconds = [int]$durationSeconds
        start = $session.start
        end = $session.end
    }

    Clear-ActiveTracking
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
        note = ""
        exported = $false
    }
    $timer.sessions += $session

    $State.data.active_timer_id = [string]$TimerId
    $State.data.active_session_id = $session.id
    $State.runtime.last_autosave_at = $now
    Save-Data
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
    return (Timer-ElapsedTodaySeconds -Timer $timer)
}

function Timer-ElapsedTodaySeconds {
    param($Timer)
    if (-not $Timer) { return 0 }
    $activeTimerId = $State.data.active_timer_id
    $activeSessionId = $State.data.active_session_id

    $todayLocal = (Get-Date).Date
    $dayStartUtc = $todayLocal.ToUniversalTime()
    $dayEndUtc = $todayLocal.AddDays(1).ToUniversalTime()

    $total = 0
    foreach ($session in $Timer.sessions) {
        $start = From-UtcString $session.start
        if (-not $start) { continue }
        if ($session.end) {
            $end = From-UtcString $session.end
        } elseif ($activeTimerId -and $activeSessionId -and $Timer.id -eq $activeTimerId -and $session.id -eq $activeSessionId) {
            $end = Utc-Now
        } else {
            continue
        }
        if (-not $end) { continue }
        if ($end -le $dayStartUtc -or $start -ge $dayEndUtc) { continue }
        if ($start -lt $dayStartUtc) { $start = $dayStartUtc }
        if ($end -gt $dayEndUtc) { $end = $dayEndUtc }
        $total += [math]::Max(0, ($end - $start).TotalSeconds)
    }
    return [int]$total
}

function Total-ForToday {
    $todayLocal = (Get-Date).Date
    $dayStartLocal = $todayLocal
    $dayEndLocal = $todayLocal.AddDays(1)
    $dayStartUtc = $dayStartLocal.ToUniversalTime()
    $dayEndUtc = $dayEndLocal.ToUniversalTime()
    $activeTimerId = $State.data.active_timer_id
    $activeSessionId = $State.data.active_session_id
    $total = 0
    foreach ($timer in $State.data.timers) {
        foreach ($session in $timer.sessions) {
            $start = From-UtcString $session.start
            if (-not $start) { continue }
            if ($session.end) {
                $end = From-UtcString $session.end
            } elseif ($activeTimerId -and $activeSessionId -and $timer.id -eq $activeTimerId -and $session.id -eq $activeSessionId) {
                $end = Utc-Now
            } else {
                continue
            }
            if (-not $end) { continue }
            if ($end -le $dayStartUtc -or $start -ge $dayEndUtc) { continue }
            if ($start -lt $dayStartUtc) { $start = $dayStartUtc }
            if ($end -gt $dayEndUtc) { $end = $dayEndUtc }
            $total += [math]::Max(0, ($end - $start).TotalSeconds)
        }
    }
    return [int]$total
}

function Format-Hms {
    param([int]$TotalSeconds)
    $h = [int]($TotalSeconds / 3600)
    $m = [int](($TotalSeconds % 3600) / 60)
    $s = [int]($TotalSeconds % 60)
    return ("{0:d2}:{1:d2}:{2:d2}" -f $h, $m, $s)
}

function Desktop-TargetDir {
    $override = $State.settings.desktop_path_override
    if ($override -and -not [string]::IsNullOrWhiteSpace($override)) {
        if (-not (Test-Path $override)) {
            New-Item -ItemType Directory -Path $override -Force | Out-Null
        }
        return $override
    }
    $path = Get-DesktopDir
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
            $session.note = $Note
            $session.exported = $true
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
                $startUtc = From-UtcString $session.start
                $durationSeconds = [int][math]::Max(0, ($endUtc - $startUtc).TotalSeconds)
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
        $item.session.exported = $true
    }
    if ($pending.Count -gt 0) {
        Save-Data
    }
}

function Stop-And-SaveAll {
    $summary = Close-ActiveSession -At (Utc-Now)
    if ($summary) {
        $timer = Get-TimerById -TimerId $summary.timer_id
        $totalSeconds = if ($timer) { Timer-ElapsedTodaySeconds -Timer $timer } else { [int]$summary.duration_seconds }
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
            if (-not $session.note) {
                $session.note = "Recovered after app restart"
            } else {
                $session.note = ($session.note + " | Recovered after app restart")
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
}

function Update-TrayTooltip {
    param([System.Windows.Forms.NotifyIcon]$Tray)
    $activeLabel = $null
    $activeTimer = Get-ActiveTimer
    if ($activeTimer) { $activeLabel = $activeTimer.label }
    $activeElapsed = Active-ElapsedSeconds
    $todayTotal = Total-ForToday
    $symbol = "Stopped"
    if ($activeTimer) {
        if ($State.data.active_session_id) {
            $symbol = "Running"
        } else {
            $symbol = "Paused"
        }
    }
    $label = if ($activeLabel) { $activeLabel } else { "No active timer" }
    $tooltip = "{0} {1} {2} | Today {3}" -f $symbol, $label, (Format-Hms -TotalSeconds $activeElapsed), (Format-Hms -TotalSeconds $todayTotal)
    if ($tooltip.Length -gt 127) { $tooltip = $tooltip.Substring(0, 127) }
    $Tray.Text = $tooltip
}

function Update-FloatingWindow {
    param([System.Windows.Forms.Form]$Window, $Panel)
    if (-not $Window -or -not $Panel) { return }
    if ($Window.IsDisposed) { return }

    $timers = @($State.data.timers | Where-Object { $_ })
    if (-not (Get-Variable -Name floatingControls -Scope Script -ErrorAction SilentlyContinue)) { $script:floatingControls = @{} }
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
        foreach ($control in @($Panel.Controls)) { $control.Dispose() }
        $Panel.Controls.Clear()
        $script:floatingControls = @{}
        $script:floatingOrder = @()
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
        foreach ($t in $timers) {
            $timerId = [string](Get-PropValue -Obj $t -Name "id")
            if (-not $timerId) { continue }
            $labelText = [string](Get-PropValue -Obj $t -Name "label")
            if ([string]::IsNullOrWhiteSpace($labelText)) { $labelText = "Untitled" }
            $color = $palette[$index % $palette.Count]
            $index += 1

            $block = New-Object System.Windows.Forms.Panel
            $block.Height = 54
            $block.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
            $block.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)

            $title = New-Object System.Windows.Forms.Label
            $title.AutoSize = $false
            $title.Dock = "Top"
            $title.Height = 22
            $title.TextAlign = "MiddleLeft"
            $title.Font = $script:floatingFonts.title
            $title.ForeColor = $color
            $title.Text = $labelText

            $time = New-Object System.Windows.Forms.Label
            $time.AutoSize = $false
            $time.Dock = "Fill"
            $time.TextAlign = "MiddleLeft"
            $time.Font = $script:floatingFonts.time
            $time.ForeColor = [System.Drawing.Color]::White

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
    }

    if (-not $timers -or $timers.Count -eq 0) {
        foreach ($control in @($Panel.Controls)) { $control.Dispose() }
        $Panel.Controls.Clear()
        $script:floatingControls = @{}
        $empty = New-Object System.Windows.Forms.Label
        $empty.AutoSize = $true
        $empty.Text = "No timers"
        $empty.ForeColor = [System.Drawing.Color]::White
        $empty.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 6)
        $Panel.Controls.Add($empty) | Out-Null
        return
    }

    $maxWidth = 240
    foreach ($t in $timers) {
        $timerId = [string](Get-PropValue -Obj $t -Name "id")
        if (-not $timerId) { continue }
        $controls = $script:floatingControls[$timerId]
        if (-not $controls) { continue }
        $labelText = [string](Get-PropValue -Obj $t -Name "label")
        if ([string]::IsNullOrWhiteSpace($labelText)) { $labelText = "Untitled" }
        $elapsed = Format-Hms -TotalSeconds (Timer-ElapsedTodaySeconds -Timer $t)
        $controls.title.Text = $labelText
        $controls.time.Text = $elapsed
        if ($needsFontRefresh) {
            $controls.title.Font = $script:floatingFonts.title
            $controls.time.Font = $script:floatingFonts.time
        }

        $titleSize = [System.Windows.Forms.TextRenderer]::MeasureText($labelText, $script:floatingFonts.title)
        $timeSize = [System.Windows.Forms.TextRenderer]::MeasureText($elapsed, $script:floatingFonts.time)
        $blockWidth = [math]::Max($titleSize.Width, $timeSize.Width) + 24
        if ($blockWidth -gt $maxWidth) { $maxWidth = $blockWidth }
    }

    foreach ($id in $script:floatingOrder) {
        $controls = $script:floatingControls[$id]
        if ($controls) {
            $controls.block.Width = $maxWidth
        }
    }

    $panelPadding = $Panel.Padding
    $panelWidth = $maxWidth + $panelPadding.Left + $panelPadding.Right + 12
    $panelHeight = 0
    foreach ($id in $script:floatingOrder) {
        $controls = $script:floatingControls[$id]
        if ($controls) {
            $panelHeight += $controls.block.Height + $controls.block.Margin.Top + $controls.block.Margin.Bottom
        }
    }
    $panelHeight += $panelPadding.Top + $panelPadding.Bottom + 4
    $Window.Width = 260
    $Window.Height = $panelHeight
}

function Pause-ActiveSession {
    $timer = Get-ActiveTimer
    if (-not $timer) { return }
    if (-not $State.data.active_session_id) { return }

    $session = Get-ActiveSession
    if (-not $session) {
        $State.data.active_session_id = $null
        Save-Data
        return
    }
    if (-not $session.end) {
        $session.end = To-UtcString (Utc-Now)
    }

    $State.data.active_session_id = $null
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
        $elapsed = Format-Hms -TotalSeconds (Timer-ElapsedTodaySeconds -Timer $t)
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
                if (-not $latest.end) {
                    $latest.end = To-UtcString (Utc-Now)
                }
                $start = From-UtcString $latest.start
                $end = From-UtcString $latest.end
                if ($start -and $end) {
                    $durationSeconds = [int][math]::Max(0, ($end - $start).TotalSeconds)
                    $summary = [ordered]@{
                        timer_id = $timer.id
                        session_id = $latest.id
                        label = $timer.label
                        duration_seconds = $durationSeconds
                        start = $latest.start
                        end = $latest.end
                    }
                }
            }
        }
    }
    if ($summary) {
        $timerForTotal = Get-TimerById -TimerId $summary.timer_id
        $totalSeconds = if ($timerForTotal) { Timer-ElapsedTodaySeconds -Timer $timerForTotal } else { [int]$summary.duration_seconds }
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
    Save-Data
}

function Reset-AllData {
    $confirm = New-ConfirmDialog -Title "Reset All Data" -Message "Clear all timers and totals? This cannot be undone."
    if (-not $confirm) { return }
    $State.data = New-DefaultData
    $State.runtime.last_autosave_at = $null
    $State.runtime.recovery_prompt = $null
    Save-Data
}

function Maybe-Autosave {
    $now = Utc-Now
    if (-not $State.runtime.last_autosave_at) {
        $State.runtime.last_autosave_at = $now
        return
    }
    if ($now -ge $State.runtime.last_autosave_at.AddSeconds($AutoSaveSeconds)) {
        Save-PendingSessions
        $State.runtime.last_autosave_at = $now
    }
}

Load-State

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

$floatingPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$floatingPanel.Dock = "Fill"
$floatingPanel.AutoScroll = $false
$floatingPanel.AutoSize = $false
$floatingPanel.AutoSizeMode = "GrowAndShrink"
$floatingPanel.WrapContents = $false
$floatingPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$floatingPanel.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
$floatingPanel.BackColor = $floatingWindow.BackColor
$floatingWindow.Controls.Add($floatingPanel)

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
Update-FloatingWindow -Window $floatingWindow -Panel $floatingPanel
$floatingWindow.Show()

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    Invoke-Safe {
        $trayItemFontSmall.Checked = ($State.settings.floating_font_scale -eq 1.0)
        $trayItemFontMed.Checked = ($State.settings.floating_font_scale -eq 1.2)
        $trayItemFontLarge.Checked = ($State.settings.floating_font_scale -eq 1.4)
        Maybe-Autosave
        Update-TrayTooltip -Tray $tray
        Update-FloatingWindow -Window $floatingWindow -Panel $floatingPanel
    }
})
$timer.Start()

[System.Windows.Forms.Application]::Run($hotKeyForm)
