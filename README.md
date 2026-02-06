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

Global-hotkey ticket timer with export logging and an always-on-top floating timer window. Runs as a PowerShell script so it avoids unsigned executable blocks.

**Status:** Active development. Not ready for production use.

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
- Each timer shows its label and total time for that ticket (ended sessions + current session if running).
- Window height grows/shrinks as timers are added/removed.
- Font: Cascadia Mono (falls back to Segoe UI).
- Tray menu: `Floating Font Size` with Default, +20%, +40%.
- Rounded corners for the window and timer cards.

## Switch Timer Display

Switch Timer shows each entry as:

`Timer Name - Time - ID`

Time is the total for that ticket.

## Data + Logs

- Data: `%APPDATA%\MyTime\data.json`
- Settings: `%APPDATA%\MyTime\settings.json`
- Export logs (default): `C:\MyTime`
  - `TicketTimeLog_YYYY-MM-DD.txt`
  - `<TICKETLABEL>_YYYY-MM-DD.txt`
- Export directory override: set `desktop_path_override` in `%APPDATA%\MyTime\settings.json` to an absolute path.

Stop log entries include:

- Total (ticket total)
- Session (last session duration)

## Reset

Use the tray menu item **Reset All Data** to clear all timers and totals.

## Development Notes

This project is under active development and not ready for general use. Known gaps:

- This is a single-script tray app; there is no installer and it runs in the background as a PowerShell process.
