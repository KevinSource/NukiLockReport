# NukiLockReport

A Windows console program that lists every Nuki smart lock on your account and every entry on each lock (app users, fobs, keypad codes) in an Excel workbook, then opens it in Excel.

It only reads from the [Nuki Web API](https://api.nuki.io); it never changes anything on your locks.

## What the report contains

| Sheet | Contents |
|---|---|
| **Entries** | One row per entry: lock, name, type, **keypad code**, enabled, status (Active / Upcoming / Expired / Disabled, colour-coded), date limits, allowed weekdays, time window, remote access, last active and created dates |
| **Locks** | Name, ID, model, server state, battery level (colour-coded) and battery-critical flag, number of entries |
| **Settings** | The time zone used for the dates, where it was set, and how to change it |

The report is saved as `Documents\Nuki Reports\Nuki Locks <date time>.xlsx`.

> **The report contains your keypad codes.** Store and share it accordingly. `.gitignore` keeps spreadsheets out of this repository.

If anything goes wrong (bad token, missing permission, lock offline, unknown time zone), each error is printed in red with an explanation, and the window stays open until you press a key.

## Requirements

- Windows 10 or 11
- [.NET 8 Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) to run it, or the .NET 8 SDK (or later) to build it
- A Nuki Web account and an API token created in Nuki Web, with permission to read smart locks and their authorizations
- Excel or another program that opens `.xlsx` files

## Setup

The API token and settings are stored with [.NET user-secrets](https://learn.microsoft.com/aspnet/core/security/app-secrets), in your Windows profile (`%APPDATA%\Microsoft\UserSecrets\<id>\secrets.json`), never in the project folder. From the `NukiLockReport` project folder:

```
dotnet user-secrets set "Nuki:ApiToken" "<your token>"
dotnet user-secrets set "Nuki:TimeZone" "Europe/Berlin"
```

You can also use environment variables instead, which take priority: `NUKI_API_TOKEN` and `NUKI_TIME_ZONE`.

### Time zone

Nuki returns all dates in UTC. The report converts them to the time zone you set, so use **the lock's time zone** to match what Nuki Web and the Nuki app show.

- IANA names such as `America/New_York` or `Europe/Berlin`, or Windows names such as `Eastern Standard Time` (list them with `tzutil /l`)
- With no setting, the PC's own time zone is used
- An unknown name is reported as an error, and the PC's time zone is used for that report

The **Settings** sheet in every report shows the time zone used and repeats these instructions.

## Build and run

Open `NukiLockReport.slnx` in Visual Studio (2022 version 17.14 or later opens `.slnx` solutions) and press F5, or from the repository folder:

```
dotnet run --project NukiLockReport
```

## Install

Publish it as a single `.exe` into any folder:

```
dotnet publish NukiLockReport\NukiLockReport.csproj -c Release -r win-x64 -o "C:\Tools\NukiLockReport"
```

The installed `.exe` finds the token and settings in the same place, so nothing else needs to be copied. To change a setting without the .NET SDK, edit `secrets.json` directly: the path is shown on the **Settings** sheet of any report.

## License

[MIT](LICENSE)
