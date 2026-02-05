```
 __  __       _______ _                 
|  \/  |     |__   __(_)                
| \  / |_   _   | |   _ _ __ ___   ___  
| |\/| | | | |  | |  | | '_ ` _ \ / _ \ 
| |  | | |_| |  | |  | | | | | | |  __/ 
|_|  |_|\__, |  |_|  |_|_| |_| |_|\___| 
         __/ |                           
        |___/                            
```

# MyTime (PowerShell Tray App)

Global-hotkey ticket timer with desktop logging and an always-on-top floating timer window. Runs as a PowerShell script so it avoids unsigned executable blocks.

**Status:** Active development. Not ready for production use.

Important: timekeeping accuracy still needs verification. More testing is required.

## Run

```powershell
Set-Location c:\DEV\MyTime
powershell.exe -ExecutionPolicy Bypass -File .\MyTime.ps1
```

## Hotkeys

- `Ctrl+Alt+S` Start/Pause
- `Ctrl+Alt+N` New timer
- `Ctrl+Alt+X` Switch timer
- `Ctrl+Alt+E` Stop timer

## Floating Window

- Always-on-top, draggable window listing all timers.
- Each timer shows its label and today’s total time for that ticket.
- Window height grows/shrinks as timers are added/removed.
- Font: Cascadia Mono (falls back to Segoe UI).
- Tray menu: `Floating Font Size` with Default, +20%, +40%.

## Switch Timer Display

Switch Timer shows each entry as:

`Timer Name - Time - ID`

Time is the total for today for that ticket.

## Data + Logs

- Data: `%APPDATA%\MyTime\data.json`
- Settings: `%APPDATA%\MyTime\settings.json`
- Logs on Desktop:
  - `TicketTimeLog_YYYY-MM-DD.txt`
  - `<TICKETLABEL>_YYYY-MM-DD.txt`

Stop log entries include:

- Total (today’s total for that ticket)
- Session (last session duration)

## Reset

Use the tray menu item **Reset All Data** to clear all timers and today totals.

## Development Notes

This project is under active development and not ready for general use. Known gaps:

- Timekeeping accuracy requires additional testing.
