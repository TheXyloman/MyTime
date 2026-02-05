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

Global-hotkey ticket timer with desktop logging. Runs as a PowerShell script so it avoids unsigned executable blocks.

**Status:** Active development. Not ready for production use.

Important: timekeeping accuracy and 15-minute reminders still need verification. More testing is required.

## Run

```powershell
Set-Location c:\dev\MyTimeCLI\MyTime
powershell.exe -ExecutionPolicy Bypass -File .\MyTime.ps1
```

## Hotkeys

- `Ctrl+Alt+S` Start/Pause
- `Ctrl+Alt+N` New timer
- `Ctrl+Alt+X` Switch timer
- `Ctrl+Alt+E` Stop timer

## Data + Logs

- Data: `%APPDATA%\MyTime\data.json`
- Settings: `%APPDATA%\MyTime\settings.json`
- Logs on Desktop:
  - `TicketTimeLog_YYYY-MM-DD.txt`
  - `<TICKETLABEL>_YYYY-MM-DD.txt`

## Reminders

Every 15 minutes while a timer is running, a reminder appears:

- **Continue**: keeps timer running and schedules the next reminder in 15 minutes
- **Stop**: stops the timer and writes the final log

## Reset

Use the tray menu item **Reset All Data** to clear all timers and today totals.

## Development Notes

This project is under active development and not ready for general use. Known gaps:

- Timekeeping accuracy requires additional testing.
- 15-minute reminder behavior requires additional testing.
