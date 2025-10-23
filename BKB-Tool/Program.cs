// Stefan Bäumer 2025
using System;
using System.Diagnostics;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Linq;

try { Console.WindowHeight = 35;} catch { }

Global.User = Environment.UserName;
IConfiguration? configuration = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile($"{Global.User}.json", optional: true, reloadOnChange: true).Build();

Global.ColorÜberschrift = Color.Aqua; // Überschrift
Global.ColorUnterschrift = Color.Aqua; // 2.Überschrift
Global.ColorBeschreibung = Color.SpringGreen2; // Beschreibung
Global.ColorPfadInProgrammen = Color.Yellow; 
Global.ColorPfadInDateien = Color.SpringGreen2; // Pfad in Dateien
Global.ColorActionInMenüs = Color.SpringGreen2; // Action in Menüs
Global.ColorEinstellungenRahmen = Color.DodgerBlue1; // Einstellungen Rahmen
Global.ColorHinweise = Color.Pink3; // Hinweise
Global.ColorKopfzeileInCSV = Color.DeepPink1_1; // Kopfzeile in CSV
Global.ColorHyperlink = Color.LightSkyBlue1; // Hyperlink
Global.ColorZahlen = Color.Tan; // Zahlen
Global.ColorTextHervorheben = Color.SpringGreen2; // Zahlen
Global.ColorFehler = Color.Red; // Zahlen
Global.ColorInfoBox = Color.Orange1; // Zahlen
Global.HilfeUrl = "https://github.com/stbaeumer/BKB-Tool";

var version = Assembly.GetExecutingAssembly().GetName().Version;
Global.AppVersion = version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "0.0.0";

Global.SchulnummernGesperrt = new List<string> { "999999" }; // Diese Schulnummern können das Programm nicht verwenden.
Global.SchulnummernPrivilegiert = new List<string>{ "177659" }; // Diese Schulnummern bekommen alle Jedermann-Punkte plus weitere Menüpunkte angezeigt.
Global.Schulnummer177659 = new List<string> { "177659" }; // Diese Schulnummer bekommt alle Privilegierten plus weitere Menüpunkte angezeigt.
Global.SchulnummernDebug = new List<string>{ "000000" }; // alles

configuration["AppDescription"] = "";// $"[bold {Global.GetColor(Global.ColorBeschreibung)}]BKB-Tool[/] - Ein Werkzeug an der Schnittstelle zwischen SchILD und Webuntis.";
//Global.DisplayHeader(configuration);
var dateien = new Dateien(configuration);

do
{
    try
    {
        dateien = new Dateien();
        configuration = Global.EinstellungenDurchlaufen(configuration, Global.Modus.ReadSilent);

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

        configuration = menuGefiltert.GetAusgewaehlterMenueintrag(configuration, ["e", "h"]);

        var i = Convert.ToInt32(configuration["Auswahl"]);
        
        if (i >= 0)
        {
            Global.DisplayHeader(configuration);
            menuGefiltert[i].RenderAuswahlÜberschrift(configuration);        
            if (menuGefiltert[i].Quelldateien.Any(q => !string.IsNullOrEmpty(q.Fehlermeldung)))
                throw new Exception($"[bold {Global.GetColor(Global.ColorHinweise)}]{menuGefiltert[i].Quelldateien.FirstOrDefault(q => !string.IsNullOrEmpty(q.Fehlermeldung)).Fehlermeldung}[/]\n[gray]{string.Join("\n", menuGefiltert[i].Quelldateien.FirstOrDefault(q => !string.IsNullOrEmpty(q.Fehlermeldung)).Hinweise)}[/]");
            //menuGefiltert[i].Quelldateien.FehlermeldungRendern(configuration);

            if (menuGefiltert[i].Quelldateien.Where(q => !string.IsNullOrEmpty(q.Fehlermeldung) && !q.IstOptional && !q.Nur177659).Any())
            {
                //AnsiConsole.MarkupLine($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren.[/]");
                throw new Exception($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren![/]");
            }
            else
            {
                Datei zieldatei = menuGefiltert[i].Ausführen();
            }

            Global.WeiterMitAnykey(configuration, menuGefiltert[i]);
        }    
    }
    catch (Exception ex)
    {
        var panel3 = new Panel(ex.Message)
            //.Header($"[bold {Global.GetColor(Global.ColorFehler)}] Es ist zu einem Fehler gekommen [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Global.ColorFehler);
        
            AnsiConsole.Write(panel3);
            Global.WeiterMitAnykey(configuration);
            
        continue; // Fehler behandeln und zum nächsten Durchlauf springen
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
        configuration = Global.Konfig("Schulnummer", Global.Modus.ReadSilent, configuration);
        bool allowPrerelease = configuration["Schulnummer"] == "000000";

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
        string os = OperatingSystem.IsWindows() ? "Windows" : OperatingSystem.IsLinux() ? "Linux" : OperatingSystem.IsMacOS() ? "macOS" : "Unknown";

        // In Version-Objekte umwandeln
        if (Version.TryParse(githubVersionClean, out var githubVer) && Version.TryParse(lokaleVersionClean, out var lokalVer))
        {
            if (githubVer > lokalVer)
            {
                Global.DisplayHeader(configuration);

                // Der Update-Hinweis wird in einem Panel angezeigt
                var panelUpdate = new Panel(new Markup($"Ein Update auf {os}-Version [tan]{githubVersion}[/] ist verfügbar. Drücken Sie eine [{Global.GetColor(Global.ColorActionInMenüs)} bold]beliebige Taste[/], um das Update zu starten."))
                    .BorderStyle(new Style(Color.Red))
                    .Expand();
                AnsiConsole.Write(panelUpdate);

                Console.ReadKey(); // Warten auf Benutzereingabe, bevor das Update gestartet wird

                // Lade die Datei nach bkb-neu.exe herunter und führe den Autoupdater aus.
                
                string downloadUrl = null;
                foreach (var asset in releases
                    .First(r => !r.GetProperty("draft").GetBoolean() && (allowPrerelease || !r.GetProperty("prerelease").GetBoolean()))
                    .GetProperty("assets").EnumerateArray())
                {
                    var name = asset.GetProperty("name").GetString();
                    
                    // Je nach Betriebssystem die richtige Datei herunterladen
                    string expectedPattern = os switch
                    {
                        "Windows" => "BKB-Tool.exe",
                        "Linux" => ".AppImage",
                        "macOS" => "BKB-Tool-macos",
                        _ => "BKB-Tool.exe"
                    };
                    
                    bool matches = os == "Linux" 
                        ? (name != null && name.EndsWith(expectedPattern, StringComparison.OrdinalIgnoreCase))
                        : (name != null && name.Equals(expectedPattern, StringComparison.OrdinalIgnoreCase));
                    
                    if (matches)
                    {
                        downloadUrl = asset.GetProperty("browser_download_url").GetString();
                        break;
                    }
                }

                if (string.IsNullOrEmpty(downloadUrl))
                {
                    AnsiConsole.MarkupLine($"[red]Fehler: Keine passende Datei für {os} im Release gefunden.[/]");
                    return configuration;
                }

                // Zielpfad für die neue Datei
                string zielDatei = os switch
                {
                    "Windows" => Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu.exe"),
                    "Linux" => Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu.AppImage"),
                    "macOS" => Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu"),
                    _ => Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu.exe")
                };

                if(os != "Windows")
                    zielDatei = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu");

                using (var webClient = new System.Net.WebClient())
                {
                    webClient.Headers.Add("User-Agent", "request");

                    AnsiConsole.Progress()
                        .Start(ctx =>
                        {
                            var task = ctx.AddTask("Lade Update herunter ");
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

                if(os != "Windows")
                    updaterPath = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool-autoupdater.sh");
                
                if (File.Exists(updaterPath))
                    File.Delete(updaterPath);

                if (os == "Windows")
                    File.WriteAllText(updaterPath,
                    "@echo off\n" +
                    "echo.\n" +
                    "echo BKB-Tool\n" +
                    "echo =========\n" +
                    "echo Warte auf Beenden von BKB-Tool.exe ...\n" +
                    "set /a counter=0\n" +
                    ":waitforend\n" +
                    "tasklist | find /I \"BKB-Tool.exe\" >nul\n" +
                    "if not errorlevel 1 (\n" +
                    "    timeout /t 1 >nul\n" +
                    "    set /a counter+=1\n" +
                    "    if %counter%==7 (\n" +
                    "        echo Das Beenden dauert zu lange. Versuchen Sie BKB-Tool im Taskmanager zu beenden oder starten Sie den Rechner neu.\n" +
                    "    )\n" +
                    "    goto waitforend\n" +
                    ")\n" +
                    "echo Ersetze alte Version ...\n" +
                    "del /F /Q BKB-Tool.exe\n" +
                    "if exist BKB-Tool.exe (\n" +
                    "    echo Fehler: Die alte BKB-Tool.exe konnte nicht gelöscht werden. Bitte schließen Sie alle Instanzen und versuchen Sie es erneut.\n" +
                    "    pause\n" +
                    "    exit /b 1\n" +
                    ")\n" +
                    "rename BKB-Tool_neu.exe BKB-Tool.exe\n" +
                    "echo Starte neue Version ...\n" +
                    "start \"\" BKB-Tool.exe\n" +
                    "exit\n"
                    );
                
                                                if (os == "Linux" || os == "macOS")
                                {
                                    // Linux / macOS: shell updater
                                                                        // ...existing code...
                                    // Linux / macOS: shell updater
                                    var shContent = @"#!/usr/bin/env bash
                                    set -euo pipefail
                                    cd ""$(dirname ""$0"")""
                                    
                                    LOG=""BKB-Tool-update.log""
                                    
                                    echo ""[$(date)] Updater gestartet in $(pwd)"" >>""$LOG""
                                    echo ""Warte auf Beenden von BKB-Tool ..."" | tee -a ""$LOG""
                                    # Wichtig: exakten Prozessnamen matchen, nicht -f (verhindert Match auf Autoupdater)
                                    while pgrep -x ""BKB-Tool"" >/dev/null 2>&1; do
                                      echo ""[$(date)] BKB-Tool läuft noch ..."" >>""$LOG""
                                      sleep 1
                                    done
                                    
                                    echo ""Ersetze alte Version ..."" | tee -a ""$LOG""
                                    if [[ -f ""BKB-Tool_neu"" ]]; then
                                      rm -f ""BKB-Tool"" || true
                                      mv -f ""BKB-Tool_neu"" ""BKB-Tool""
                                      chmod +x ""BKB-Tool""
                                      echo ""Starte neue Version ..."" | tee -a ""$LOG""
                                      nohup ./BKB-Tool >/dev/null 2>&1 &
                                      echo ""Update abgeschlossen."" | tee -a ""$LOG""
                                    else
                                      echo ""Fehler: BKB-Tool_neu wurde nicht gefunden."" | tee -a ""$LOG""
                                      exit 1
                                    fi
                                    ";
                                    shContent = shContent.Replace("\r\n", "\n");
                                    File.WriteAllText(updaterPath, shContent, new System.Text.UTF8Encoding(false));
                                    
                                    try
                                    {
                                        var chmod = new ProcessStartInfo
                                        {
                                            FileName = "bash",
                                            ArgumentList = { "-c", $"chmod +x \"{updaterPath}\"" },
                                            UseShellExecute = false
                                        };
                                        using var proc = Process.Start(chmod);
                                        proc?.WaitForExit();
                                    }
                                    catch { }
                                    
                                    // Updater in neuem Terminalfenster starten
                                    try
                                    {
                                        string? terminal = null;
                                        if (File.Exists("/usr/bin/gnome-terminal")) terminal = "/usr/bin/gnome-terminal";
                                        else if (File.Exists("/usr/bin/konsole")) terminal = "/usr/bin/konsole";
                                        else if (File.Exists("/usr/bin/xterm")) terminal = "/usr/bin/xterm";
                                    
                                        if (terminal != null)
                                        {
                                            var args = terminal.Contains("gnome-terminal")
                                                ? $"-- bash -lc '\"{updaterPath}\"; read -n1 -s -p \"Fertig. Fenster schließen...\"'"
                                                : terminal.Contains("konsole")
                                                    ? $"--noclose -e bash -lc '\"{updaterPath}\"'"
                                                    : $"-hold -e bash -lc '\"{updaterPath}\"'";
                                    
                                            Process.Start(new ProcessStartInfo
                                            {
                                                FileName = terminal,
                                                Arguments = args,
                                                UseShellExecute = false,
                                                CreateNoWindow = false
                                            });
                                        }
                                        else
                                        {
                                            // Fallback: Hintergrund
                                            Process.Start(new ProcessStartInfo
                                            {
                                                FileName = "bash",
                                                ArgumentList = { "-c", $"nohup \"{updaterPath}\" >/dev/null 2>&1 &" },
                                                UseShellExecute = false,
                                                CreateNoWindow = true
                                            });
                                        }
                                    
                                        AnsiConsole.MarkupLine("[yellow]Update wird in einem neuen Terminalfenster gestartet...[/]");
                                        System.Threading.Thread.Sleep(500);
                                        Environment.Exit(0);
                                    }
                                    catch (Exception ex)
                                    {
                                        AnsiConsole.MarkupLine($"[red]Fehler beim Starten des Updaters: {ex.Message}[/]");
                                    }
                                    // ...existing code...
                                }

                var panel = new Panel(
                    //$"Die neue Datei wurde heruntergeladen und gespeichert als [{Global.GetColor(Global.ColorPfadInDateien)}]{zielDatei}[/].\n" +
                    $"Mit [{Global.GetColor(Global.ColorActionInMenüs)} bold]ENTER[/] wird jetzt in die Version v{githubVer} neugestartet.")
                    .Header("[bold green]  Update erfolgreich  [/]")
                    .HeaderAlignment(Justify.Left)
                    .SquareBorder()
                    .Expand()
                    .BorderColor(Global.ColorActionInMenüs);
                
                AnsiConsole.Write(panel);

                while (Console.KeyAvailable) Console.ReadKey(true);
                Console.ReadKey();

                Process.Start(new ProcessStartInfo
                {
                    FileName = updaterPath,
                    UseShellExecute = true,
                    CreateNoWindow = true
                });

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