// Stefan Bäumer 2025
using System.Reflection;
using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

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

configuration["AppName"] = "BKB-Tool";
Global.AppVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0.0"; // Major.Minor.Build.Revision
Global.SchulnummernPrivilegiert = new List<string>{"177659"};
Global.SchulnummernGesperrt = new List<string>{"999999"};

configuration["AppDescription"] = "BKB-Tool - Ein Werkzeug an der Schnittstelle zwischen SchILD und Untis.";

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

    //dateien.DisplayHeader(configuration, dateien.Meldung);

    var menuGefiltert = new Menue();
    menuGefiltert.AddRange(menu.Where(m => m.BeiDiesenSchulnummernAnzeigen == Global.NurBeiDiesenSchulnummern.Alle || Global.SchulnummernPrivilegiert.Contains(configuration["Schulnummer"])));
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

void CheckForUpdate(IConfiguration configuration)
{
    if (Global.RunningInCodeSpace())
    {
        AnsiConsole.MarkupLine("[bold yellow]Running in Codespace, skipping update check.[/]");
        return;
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
                    $"Drücken Sie eine [green bold]beliebige Taste[/], um das Update zu starten."
                ]
                );
                
                Console.ReadKey(); // Warten auf Benutzereingabe, bevor das Update gestartet wird

                // Die Autoupdater-Batch-Datei wird erzeugt und neben die exe gespeichert. Danach wird sie ausgeführt.
                string updaterPath = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool-autoupdater.bat");

                // wenn die Datei schon existiert, dann löschen
                if (File.Exists(updaterPath))
                {
                    AnsiConsole.MarkupLine($"[bold yellow]{updaterPath}, wird gelöscht.[/]");
                    File.Delete(updaterPath);
                }

                var downloadUrl = "https://github.com/stbaeumer/BKB-Tool/releases/latest/download/BKB-Tool.exe";

                File.WriteAllText(updaterPath,
        "@echo off\n" +
        "echo.\n" +
        "echo BKB-Tool\n" +
        "echo =========\n" +
        "echo Update wird heruntergeladen...\n" +
        $"curl -L -# -o BKB-Tool_neu.exe {downloadUrl}\n" +
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

                // Autoupdater ausführen
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = updaterPath,
                    UseShellExecute = true,
                    CreateNoWindow = true
                });
                Environment.Exit(0); // Beendet das aktuelle Programm sofort, damit das Update funktioniert

                // Autoupdater ausführen
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = updaterPath,
                    UseShellExecute = true,
                    CreateNoWindow = true
                });

                Console.ReadKey();
            }
        }
        else
        {
            //AnsiConsole.MarkupLine("[bold green]Keine Updates verfügbar.[/]");
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
}
