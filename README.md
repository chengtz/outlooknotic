# Outlook Feishu Mail Monitor

Console test program for monitoring the local Outlook Inbox on Windows and
sending a Feishu bot webhook alert when new mail arrives.

## How it works

- Uses local Outlook COM automation, so it relies on the Outlook profile already
  logged in on the current Windows user.
- Does not connect to Exchange from the public internet.
- Prompts for the Feishu webhook URL after startup.
- Polls the default Outlook Inbox every 15 seconds.
- Ignores mail already present when the program starts.
- Stops completely when the console program is closed or `Ctrl+C` is pressed.

## Requirements

- Windows
- Microsoft Outlook installed and logged in
- Go installed
- Network access from this computer to the Feishu webhook URL

## Run

```powershell
go mod tidy
go run .
```

Then paste the Feishu bot webhook URL when prompted.

## Build EXE

```powershell
go mod tidy
go build -o outlook-feishu-monitor.exe .
.\outlook-feishu-monitor.exe
```

## Notes

If the company Feishu bot has signature verification enabled, the program will
need the webhook secret added later. The current version supports a plain Feishu
custom bot webhook.
