using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using NukiLockReport;

// Lists every Nuki lock on the account and all of its entries (with date limits and keypad codes),
// writes them to an .xlsx in Documents\Nuki Reports and opens it in Excel.
// Errors are printed in red; if there were any, the window waits for a key before closing.

var errors = new List<string>();

// Dates are converted from UTC to this zone. Set it to the lock's time zone so they match Nuki Web.
var (zoneSetting, zoneSource) = Secrets.Get(Secrets.TimeZoneKey, Secrets.TimeZoneEnv);
var lockZone = TimeZoneInfo.Local;
string zoneSetBy = "Not set - using this PC's time zone";
if (zoneSetting != null)
{
    try
    {
        lockZone = TimeZoneInfo.FindSystemTimeZoneById(zoneSetting);
        zoneSetBy = zoneSource!;
    }
    catch (Exception ex) when (ex is TimeZoneNotFoundException or InvalidTimeZoneException)
    {
        Error($"Time zone \"{zoneSetting}\" ({zoneSource}) is not a known time zone, so dates use this PC's time zone. " +
              "See the Settings sheet in the report for how to fix it.");
        zoneSetBy = $"\"{zoneSetting}\" ({zoneSource}) is not a known time zone - using this PC's time zone";
    }
}
string zoneLabel = lockZone.Id;

try
{
    await Run();
}
catch (Exception ex)
{
    Error($"Unexpected error: {ex}");
}

if (errors.Count > 0)
{
    Header($"{errors.Count} error(s)");
    foreach (string e in errors)
        Error(e, record: false);
    if (!Console.IsInputRedirected)
    {
        Console.WriteLine("\nPress any key to close...");
        Console.ReadKey(intercept: true);
    }
    return 1;
}
return 0;

async Task Run()
{
    string? token = Secrets.GetToken();
    if (token == null)
    {
        Error("No API token found. Set it once from the project folder with:  dotnet user-secrets set \"Nuki:ApiToken\" \"<your token>\"  (or set NUKI_API_TOKEN).");
        return;
    }

    using var api = new NukiApi(token);

    Header("Reading locks and entries");
    var lockList = await api.GetAsync("smartlock");
    if (!ApiOk("List locks", lockList))
        return;

    using var locksDoc = JsonDocument.Parse(lockList.Body);
    var locks = locksDoc.RootElement.EnumerateArray().ToList();
    if (locks.Count == 0)
    {
        Error("The token can see no locks.");
        return;
    }

    var entries = new List<(string LockName, JsonElement Auth)>();
    var authDocs = new List<JsonDocument>();
    try
    {
        foreach (var l in locks)
        {
            string name = LockName(l);
            var auths = await api.GetAsync($"smartlock/{l.GetProperty("smartlockId").GetInt64()}/auth");
            if (!ApiOk($"Read entries for '{name}'", auths))
                continue;

            var doc = JsonDocument.Parse(auths.Body);
            authDocs.Add(doc);
            var list = doc.RootElement.EnumerateArray().ToList();
            entries.AddRange(list.Select(a => (name, a)));
            Console.WriteLine($"  {name,-30} {list.Count} entries");
        }

        Header("Building workbook");
        string path = BuildWorkbook(locks, entries);
        Console.WriteLine($"  Saved {path}");

        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            Error($"Could not open the file in Excel: {ex.Message}");
        }
    }
    finally
    {
        foreach (var d in authDocs) d.Dispose();
    }
}

string BuildWorkbook(List<JsonElement> locks, List<(string LockName, JsonElement Auth)> entries)
{
    using var wb = new XLWorkbook();

    // ---------------------------------------------------------------- Entries
    var ws = wb.Worksheets.Add("Entries");
    string[] headers =
    [
        "Lock", "Name", "Type", "Code", "Enabled", "Status", "Date limited",
        $"Valid from ({zoneLabel})", $"Valid until ({zoneLabel})", "Days", "Time window",
        "Remote allowed", $"Last active ({zoneLabel})", $"Created ({zoneLabel})"
    ];
    WriteHeaders(ws, headers);

    var now = DateTime.UtcNow;
    int row = 2;
    foreach (var (lockName, a) in entries
                 .OrderBy(e => e.LockName)
                 .ThenBy(e => Date(e.Auth, "allowedFromDate") ?? DateTime.MaxValue)
                 .ThenBy(e => Str(e.Auth, "name")))
    {
        DateTime? from = Date(a, "allowedFromDate");
        DateTime? until = Date(a, "allowedUntilDate");
        bool enabled = Bool(a, "enabled") ?? false;
        string status = !enabled ? "Disabled"
            : until != null && until < now ? "Expired"
            : from != null && from > now ? "Upcoming"
            : "Active";

        ws.Cell(row, 1).Value = lockName;
        ws.Cell(row, 2).Value = Str(a, "name") ?? "";
        ws.Cell(row, 3).Value = AuthType(Int(a, "type") ?? -1);
        ws.Cell(row, 4).Value = Int(a, "code")?.ToString() ?? "";
        ws.Cell(row, 5).Value = enabled ? "Yes" : "No";
        ws.Cell(row, 6).Value = status;
        ws.Cell(row, 7).Value = from != null || until != null ? "Yes" : "No";
        SetDate(ws.Cell(row, 8), from);
        SetDate(ws.Cell(row, 9), until);
        ws.Cell(row, 10).Value = WeekDays(Int(a, "allowedWeekDays"));
        ws.Cell(row, 11).Value = TimeWindow(Int(a, "allowedFromTime"), Int(a, "allowedUntilTime"));
        ws.Cell(row, 12).Value = Bool(a, "remoteAllowed") == true ? "Yes" : "No";
        SetDate(ws.Cell(row, 13), Date(a, "lastActiveDate"));
        SetDate(ws.Cell(row, 14), Date(a, "creationDate"));

        ws.Cell(row, 6).Style.Fill.BackgroundColor = status switch
        {
            "Active" => XLColor.FromHtml("#C6EFCE"),
            "Upcoming" => XLColor.FromHtml("#DDEBF7"),
            "Expired" => XLColor.FromHtml("#EDEDED"),
            _ => XLColor.FromHtml("#FFC7CE")
        };
        row++;
    }
    Finish(ws, headers.Length, row - 1);

    // ---------------------------------------------------------------- Locks
    var ls = wb.Worksheets.Add("Locks");
    string[] lockHeaders = ["Name", "Smartlock ID", "Type", "Server state", "Battery level", "Battery critical", "Entries"];
    WriteHeaders(ls, lockHeaders);
    row = 2;
    foreach (var l in locks.OrderBy(LockName))
    {
        string name = LockName(l);
        bool hasState = l.TryGetProperty("state", out var state);
        bool? battery = hasState ? Bool(state, "batteryCritical") : null;
        int? charge = hasState ? Int(state, "batteryCharge") : null;
        ls.Cell(row, 1).Value = name;
        ls.Cell(row, 2).Value = l.GetProperty("smartlockId").GetInt64().ToString();
        ls.Cell(row, 3).Value = LockType(Int(l, "type") ?? -1);
        ls.Cell(row, 4).Value = ServerState(Int(l, "serverState") ?? -1);
        if (charge != null)
        {
            // batteryCharge is a whole percentage; store it as a fraction so Excel shows "29%"
            ls.Cell(row, 5).Value = charge.Value / 100.0;
            ls.Cell(row, 5).Style.NumberFormat.Format = "0%";
            ls.Cell(row, 5).Style.Fill.BackgroundColor = charge switch
            {
                <= 20 => XLColor.FromHtml("#FFC7CE"),
                <= 40 => XLColor.FromHtml("#FFEB9C"),
                _ => XLColor.FromHtml("#C6EFCE")
            };
        }
        ls.Cell(row, 6).Value = battery == null ? "" : battery.Value ? "Yes" : "No";
        ls.Cell(row, 7).Value = entries.Count(e => e.LockName == name);
        row++;
    }
    Finish(ls, lockHeaders.Length, row - 1);

    // ---------------------------------------------------------------- Settings
    var ss = wb.Worksheets.Add("Settings");
    WriteHeaders(ss, ["Setting", "Value"]);
    TimeSpan offset = lockZone.GetUtcOffset(DateTime.UtcNow);
    string[,] settings =
    {
        { "Time zone", lockZone.Id },
        { "Description", lockZone.DisplayName },
        { "UTC offset now", $"UTC{(offset < TimeSpan.Zero ? "-" : "+")}{offset:hh\\:mm}" },
        { "Set by", zoneSetBy },
        { "Report created", $"{DateTime.Now:yyyy-MM-dd HH:mm} ({TimeZoneInfo.Local.Id}, this PC)" }
    };
    for (int i = 0; i < settings.GetLength(0); i++)
    {
        ss.Cell(i + 2, 1).Value = settings[i, 0];
        ss.Cell(i + 2, 2).Value = settings[i, 1];
    }

    string example = zoneSetting ?? "America/New_York";
    string[] help =
    [
        "All dates in this report are converted from UTC to the time zone above.",
        "Set it to the lock's time zone so the dates match what Nuki Web and the Nuki app show.",
        "",
        "Option 1 - edit the settings file (works for the installed .exe). Add or change this line:",
        $"    \"{Secrets.TimeZoneKey}\": \"{example}\",",
        $"    in {Secrets.FilePath}",
        "Option 2 - from the NukiLockReport project folder (needs the .NET SDK):",
        $"    dotnet user-secrets set \"{Secrets.TimeZoneKey}\" \"{example}\"",
        $"Option 3 - set the {Secrets.TimeZoneEnv} environment variable. It takes priority over options 1 and 2.",
        "",
        "Names: IANA names such as America/New_York, America/Chicago or Europe/Berlin,",
        "    or Windows names such as Eastern Standard Time (list them with: tzutil /l).",
        "Remove the setting to use this PC's time zone. Run the report again after changing it."
    ];
    int helpRow = settings.GetLength(0) + 3;
    ss.Cell(helpRow, 1).Value = "How to change the time zone";
    ss.Cell(helpRow, 1).Style.Font.Bold = true;
    for (int i = 0; i < help.Length; i++)
        ss.Cell(helpRow + 1 + i, 1).Value = help[i];

    ss.Column(1).Width = 18;
    ss.Column(2).Width = 60;
    if (zoneSetBy.Contains("not a known"))
        ss.Cell(5, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC7CE");

    string folder =Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Nuki Reports");
    Directory.CreateDirectory(folder);
    string path = Path.Combine(folder, $"Nuki Locks {DateTime.Now:yyyy-MM-dd HHmmss}.xlsx");
    wb.SaveAs(path);
    return path;
}

void SetDate(IXLCell cell, DateTime? utc)
{
    if (utc == null) return;
    cell.Value = TimeZoneInfo.ConvertTimeFromUtc(utc.Value, lockZone);
    cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm";
}

bool ApiOk(string what, ApiResult r)
{
    if (r.Ok) return true;
    string detail = $"{NukiApi.Explain(r.Status)} {Snippet(r.Body)}".Trim();
    Error($"{what}: {(r.Status == 0 ? "no response" : $"HTTP {r.Status}")}{(detail == "" ? "" : " - " + detail)}");
    return false;
}

void Error(string text, bool record = true)
{
    if (record) errors.Add(text);
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"  {text}");
    Console.ResetColor();
}

// ================================================================ helpers

static void WriteHeaders(IXLWorksheet ws, string[] headers)
{
    for (int i = 0; i < headers.Length; i++)
        ws.Cell(1, i + 1).Value = headers[i];
    var range = ws.Range(1, 1, 1, headers.Length);
    range.Style.Font.Bold = true;
    range.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9E1F2");
}

static void Finish(IXLWorksheet ws, int columns, int lastRow)
{
    ws.SheetView.FreezeRows(1);
    ws.Range(1, 1, Math.Max(lastRow, 1), columns).SetAutoFilter();
    ws.Columns(1, columns).AdjustToContents();
}

static void Header(string text)
{
    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine($"\n== {text}");
    Console.ResetColor();
}

static string LockName(JsonElement l) => Str(l, "name") ?? l.GetProperty("smartlockId").GetInt64().ToString();
static string Snippet(string body) => body.Length > 200 ? body[..200] + "..." : body;

static string? Str(JsonElement e, string name) =>
    e.TryGetProperty(name, out var v) && v.ValueKind != JsonValueKind.Null ? (v.ValueKind == JsonValueKind.String ? v.GetString() : v.GetRawText()) : null;
static int? Int(JsonElement e, string name) =>
    e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out int i) ? i : null;
static bool? Bool(JsonElement e, string name) =>
    e.TryGetProperty(name, out var v) && v.ValueKind is JsonValueKind.True or JsonValueKind.False ? v.GetBoolean() : null;
static DateTime? Date(JsonElement e, string name) =>
    Str(e, name) is string s && DateTime.TryParse(s, CultureInfo.InvariantCulture,
        DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var d) ? d : null;

// Bitmask: Mon=64 Tue=32 Wed=16 Thu=8 Fri=4 Sat=2 Sun=1; 0 or 127 means no restriction
static string WeekDays(int? mask)
{
    if (mask is null or 0 or 127) return "Every day";
    string[] names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
    return string.Join(" ", names.Where((_, i) => (mask.Value & (64 >> i)) != 0));
}

// Minutes after midnight; both 0 means all day
static string TimeWindow(int? fromMin, int? untilMin)
{
    if ((fromMin ?? 0) == 0 && (untilMin ?? 0) == 0) return "All day";
    static string Hm(int? m) => $"{(m ?? 0) / 60:00}:{(m ?? 0) % 60:00}";
    return $"{Hm(fromMin)}-{Hm(untilMin)}";
}

static string LockType(int t) => t switch
{
    0 => "Smart Lock 1/2", 1 => "Box", 2 => "Opener", 3 => "Smart Door", 4 => "Smart Lock 3/4", 5 => "Smart Lock 5", _ => t.ToString()
};
static string ServerState(int s) => s switch
{
    0 => "ok", 1 => "unregistered", 2 => "auth uuid invalid", 3 => "auth invalid", 4 => "offline", _ => s.ToString()
};
static string AuthType(int t) => t switch
{
    0 => "App", 1 => "Bridge", 2 => "Fob", 3 => "Keypad", 13 => "Keypad code", 14 => "Z-key", 15 => "Virtual", _ => $"Type {t}"
};
