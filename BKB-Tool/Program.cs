// Stefan Bäumer 2025
using System.Reflection;
using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

try{ Console.WindowHeight = 33;} catch { }

// Pfad in Programmen: yellow
// Pfad in Dateien: aqua
// Action in Menüs: springGreen2
// Einstellungen Rahmen: dodgerBlue1
// Hinweise: pink3
// Kopfzeile in CSV: deeppink1_1
// Hyperlink: lightskyblue3_1
// Zahlen: tan
// Überschrift: springGreen2


Global.User = Environment.UserName;
IConfiguration? configuration = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile($"{Global.User}.json", optional: true, reloadOnChange: true).Build();

Global.AppVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0.0"; // Major.Minor.Build.Revision

Global.SchulnummernGesperrt = new List<string> { "999999" }; // Diese Schulnummer können das Programm nicht verwenden.
Global.SchulnummernPrivilegiert = new List<string>{ "177659" }; // Diese Schulnummern bekommen alle Jedermann-Punkte plus weitere Menüpunkte angezeigt.
Global.Schulnummer177659 = new List<string> { "177659" }; // Diese Schulnummer bekommen alle Privilegierten plus weitere Menüpunkte angezeigt.
Global.SchulnummernDebug = new List<string>{ "000000" }; // alles

configuration["AppDescription"] = "[bold springGreen2]BKB-Tool[/] - Ein Werkzeug an der Schnittstelle zwischen SchILD und Untis.";

var dateien = new Dateien(configuration);

do
{
    dateien = new Dateien();
    if (!File.Exists(Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json")))
    {
        configuration = Global.EinstellungenDurchlaufen(Global.Modus.Create, configuration);
    }
    else
    {
        configuration = Global.EinstellungenDurchlaufen(Global.Modus.Read, configuration);
    }

    CheckForUpdate(configuration);
    var table = new Table().Centered();

    dateien.ExportAusSchildVerschieben(configuration);
    dateien.GetInteressierendeDateienMitAllenEigenschaften(configuration);
    dateien.GetZeilen(configuration);

    var menu = MenueHelper.Einlesen(dateien, configuration);
    if (menu == null) continue;
    var menuGefiltert = new Menue();
    menuGefiltert.AddRange(menu.Where(m =>
             m.NurBeiDiesenSchulnummern == Global.NurBeiDiesenSchulnummern.Alle ||
            (m.NurBeiDiesenSchulnummern == Global.NurBeiDiesenSchulnummern.Nur000000 && configuration["Schulnummer"] == "000000") ||
            (m.NurBeiDiesenSchulnummern == Global.NurBeiDiesenSchulnummern.Nur177659 && configuration["Schulnummer"] == "177659") ||
            (m.NurBeiDiesenSchulnummern == Global.NurBeiDiesenSchulnummern.NurPrivilegiert && Global.SchulnummernPrivilegiert.Contains(configuration["Schulnummer"])) ||
            (m.NurBeiDiesenSchulnummern == Global.NurBeiDiesenSchulnummern.AlleBisAufGesperrte && !Global.SchulnummernGesperrt.Contains(configuration["Schulnummer"]))));
    menuGefiltert.AuswahlGridRendern();

    configuration = menuGefiltert.GetAusgewaehlterMenueintrag(configuration, ["x", "y"]);
    var i = Convert.ToInt32(configuration["Auswahl"]);

    dateien.DisplayHeader(configuration);
    if (i >= 0)
    {
        menuGefiltert[i].RenderAuswahlÜberschrift(configuration);
        menuGefiltert[i].Quelldateien.FehlermeldungRendern(configuration);

        if (menuGefiltert[i].Quelldateien.Where(q => !string.IsNullOrEmpty(q.Fehlermeldung) && !q.IstOptional).Any())
        {
            AnsiConsole.MarkupLine($"[grey]Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren.[/]");
        }
        else
        {
            Datei zieldatei = menuGefiltert[i].Ausführen();
        }

        Global.WeiterMitAnykey(configuration, menuGefiltert[i]);
    }

} while (true);

IConfiguration CheckForUpdate(IConfiguration configuration)
{
    if (Global.RunningInCodeSpace())
    {
        AnsiConsole.MarkupLine("[bold yellow]Running in Codespace, skipping update check.[/]");
        return configuration;
    }

    try
    {
        // Lokale Version (z.B. aus AssemblyInfo)
        string lokaleVersion = Global.AppVersion ?? "0.1";

        // GitHub API abfragen
        string apiUrl = "https://api.github.com/repos/stbaeumer/BKB-Tool/releases";
        var client = new System.Net.WebClient();
        client.Headers.Add("User-Agent", "request"); // GitHub verlangt einen User-Agent
        string json = client.DownloadString(apiUrl);

        using var doc = System.Text.Json.JsonDocument.Parse(json);
        var releases = doc.RootElement.EnumerateArray();
        string githubVersion = null;
        bool allowPrerelease = configuration["Schulnummer"] == "177659";

        foreach (var release in releases)
        {
            if (!release.GetProperty("draft").GetBoolean())
            {
                bool isPrerelease = release.GetProperty("prerelease").GetBoolean();
                if (allowPrerelease || !isPrerelease)
                {
                    githubVersion = release.GetProperty("tag_name").GetString();
                    break;
                }
            }
        }

        string githubVersionClean = githubVersion?.TrimStart('v', 'V');
        string lokaleVersionClean = lokaleVersion?.TrimStart('v', 'V');

        // In Version-Objekte umwandeln
        if (Version.TryParse(githubVersionClean, out var githubVer) && Version.TryParse(lokaleVersionClean, out var lokalVer))
        {
            if (githubVer > lokalVer)
            {
                dateien.DisplayHeader(configuration,
                [
                    $"Ein Update auf Version [tan]{githubVersion}[/] ist verfügbar.",                    
                    $"Drücken Sie eine [springGreen2 bold]beliebige Taste[/], um das Update zu starten."
                ]
                );
                
                Console.ReadKey(); // Warten auf Benutzereingabe, bevor das Update gestartet wird

                // Lade die Datei nach bkb-neu.exe herunter und führe den Autoupdater aus.

                // Download-URL aus dem Release-Asset holen (wie zuvor empfohlen)
                string downloadUrl = null;
                foreach (var asset in releases
                    .First(r => !r.GetProperty("draft").GetBoolean() && (allowPrerelease || !r.GetProperty("prerelease").GetBoolean()))
                    .GetProperty("assets").EnumerateArray())
                {
                    var name = asset.GetProperty("name").GetString();
                    if (name != null && name.Equals("BKB-Tool.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        downloadUrl = asset.GetProperty("browser_download_url").GetString();
                        break;
                    }
                }

                // Zielpfad für die neue Datei
                string zielDatei = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu.exe");

                using (var webClient = new System.Net.WebClient())
                {
                    webClient.Headers.Add("User-Agent", "request");

                    AnsiConsole.Progress()
                        .Start(ctx =>
                        {
                            var task = ctx.AddTask("Lade Update herunter ...");
                            webClient.DownloadProgressChanged += (s, e) =>
                            {
                                task.Value = e.ProgressPercentage;
                            };
                            var downloadCompleted = new System.Threading.ManualResetEvent(false);
                            webClient.DownloadFileCompleted += (s, e) =>
                            {
                                task.Value = 100;
                                downloadCompleted.Set();
                            };
                            webClient.DownloadFileAsync(new Uri(downloadUrl), zielDatei);
                            downloadCompleted.WaitOne();
                        });
                }

                string updaterPath = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool-autoupdater.bat");

                // wenn die Datei schon existiert, dann löschen
                if (File.Exists(updaterPath))
                {
                    AnsiConsole.MarkupLine($"[bold yellow]{updaterPath}, wird gelöscht.[/]");
                    File.Delete(updaterPath);
                }

                File.WriteAllText(updaterPath,
                    "@echo off\n" +
                    "echo.\n" +
                    "echo BKB-Tool\n" +
                    "echo =========\n" +
                    "echo Warte auf Beenden von BKB-Tool.exe ...\n" +
                    ":waitforend\n" +
                    "tasklist | find /I \"BKB-Tool.exe\" >nul\n" +
                    "if not errorlevel 1 (\n" +
                    "    timeout /t 1 >nul\n" +
                    "    goto waitforend\n" +
                    ")\n" +
                    "echo Ersetze alte Version ...\n" +
                    "del /F /Q BKB-Tool.exe\n" +
                    "rename BKB-Tool_neu.exe BKB-Tool.exe\n" +
                    "echo Starte neue Version ...\n" +
                    "start \"\" BKB-Tool.exe\n" +
                    "exit\n"
                );

                dateien.DisplayHeader(configuration,
                [
                    $"Die neue Datei wurde heruntergeladen und gespeichert als [aqua]{zielDatei}[/].",                    
                    $"Mit [springGreen2 bold]Enter[/] wird jetzt in die Version {githubVer} neugestartet."
                ]);
                Console.ReadKey();

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = updaterPath,
                    UseShellExecute = true,
                    CreateNoWindow = true
                });

                // Bei einem Update wird die Auswahl auf "x" gesetzt, damit der Anwender bei einer möglicherweise veränderten Reihenfolge der Menüpunkte nicht durcheinander kommt.
                configuration = Global.Konfig("Auswahl", Global.Modus.Update, configuration, "", "", Global.Datentyp.String, "x");

                Environment.Exit(0); // Beendet das aktuelle Programm sofort, damit das Update funktioniert                
            }
        }
        else
        {
            //AnsiConsole.MarkupLine("[bold springGreen2]Keine Updates verfügbar.[/]");
            // Auto-Updater-Batch-Datei löschen, wenn sie existiert
            string updaterPath = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool-autoupdater.bat");
            if (File.Exists(updaterPath))
            {
                File.Delete(updaterPath);
            }
        }
    }
    catch (Exception ex)
    {
        AnsiConsole.MarkupLine($"[bold red]Error starting updater: {ex.Message}[/]");
    }

    // Sicherstellen, dass die Konfiguration zurückgegeben wird, auch wenn ein Fehler auftritt
    return configuration;    
}
