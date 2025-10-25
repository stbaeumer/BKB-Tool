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

configuration["AppDescription"] = "";
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
        string lokaleVersion = Global.AppVersion ?? "0.1";
        string apiUrl = "https://api.github.com/repos/stbaeumer/BKB-Tool/releases";
        using var client = new System.Net.WebClient();
        client.Headers.Add("User-Agent", "request");
        string json = client.DownloadString(apiUrl);

        using var doc = System.Text.Json.JsonDocument.Parse(json);
        var releases = doc.RootElement.EnumerateArray();
        string githubVersion = null;

        configuration = Global.Konfig("Schulnummer", Global.Modus.ReadSilent, configuration);
        bool allowPrerelease = configuration["Schulnummer"] == "000000";

        System.Text.Json.JsonElement? selectedRelease = null;
        foreach (var release in releases)
        {
            if (!release.GetProperty("draft").GetBoolean())
            {
                bool isPrerelease = release.GetProperty("prerelease").GetBoolean();
                if (allowPrerelease || !isPrerelease)
                {
                    githubVersion = release.GetProperty("tag_name").GetString();
                    selectedRelease = release;
                    break;
                }
            }
        }

        string githubVersionClean = githubVersion?.TrimStart('v', 'V');
        string lokaleVersionClean = lokaleVersion?.TrimStart('v', 'V');
        string os = OperatingSystem.IsWindows() ? "Windows" : OperatingSystem.IsLinux() ? "Linux" : OperatingSystem.IsMacOS() ? "macOS" : "Unknown";

        if (Version.TryParse(githubVersionClean, out var githubVer) && Version.TryParse(lokaleVersionClean, out var lokalVer))
        {
            if (githubVer > lokalVer)
            {
                Global.DisplayHeader(configuration);

                var updatePanel = new Panel(new Markup($"Ein Update auf {os}-Version [tan]{githubVersionClean}[/] ist verfügbar. Drücken Sie eine [{Global.GetColor(Global.ColorActionInMenüs)} bold]beliebige Taste[/], um das Update zu starten."))
                    .BorderStyle(new Style(Color.Red))
                    .Expand();
                AnsiConsole.Write(updatePanel);
                Console.ReadKey(true);

                if (os == "Linux")
                {
                    if (selectedRelease is null)
                    {
                        AnsiConsole.MarkupLine("[red]Fehler: Release nicht gefunden.[/]");
                        return configuration;
                    }

                    // 1) Download-URL für AppImage finden (Name exakt: BKB-Tool.AppImage)
                    string downloadUrl = null;
                    foreach (var asset in selectedRelease.Value.GetProperty("assets").EnumerateArray())
                    {
                        var name = asset.GetProperty("name").GetString();
                        if (string.Equals(name, "BKB-Tool.AppImage", StringComparison.OrdinalIgnoreCase))
                        {
                            downloadUrl = asset.GetProperty("browser_download_url").GetString();
                            break;
                        }
                    }

                    if (string.IsNullOrEmpty(downloadUrl))
                    {
                        AnsiConsole.MarkupLine("[red]Fehler: BKB-Tool.AppImage im Release nicht gefunden.[/]");
                        return configuration;
                    }

                    // 2) Zielpfade bestimmen
                    string appImagePath = Environment.GetEnvironmentVariable("APPIMAGE");
                    if (string.IsNullOrEmpty(appImagePath))
                    {
                        // Fallback: im aktuellen Ordner nach BKB-Tool.AppImage oder BKB-Tool*.AppImage suchen
                        var guess = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool.AppImage");
                        if (File.Exists(guess))
                            appImagePath = guess;
                        else
                            appImagePath = Directory.GetFiles(Directory.GetCurrentDirectory(), "BKB-Tool*.AppImage").FirstOrDefault() ?? "";
                    }
                    if (string.IsNullOrEmpty(appImagePath) || !File.Exists(appImagePath))
                    {
                        AnsiConsole.MarkupLine("[red]Fehler: Laufendes AppImage nicht ermittelt. Starten Sie BKB-Tool als AppImage und versuchen Sie es erneut.[/]");
                        return configuration;
                    }

                    string appDir = Path.GetDirectoryName(appImagePath) ?? Directory.GetCurrentDirectory();
                    string newPath = Path.Combine(appDir, "BKB-Tool_neu.AppImage");

                    // 3) Neue Version herunterladen (mit Fortschritt)
                    using (var wc = new System.Net.WebClient())
                    {
                        wc.Headers.Add("User-Agent", "request");
                        AnsiConsole.Progress().Start(ctx =>
                        {
                            var task = ctx.AddTask("Lade Update herunter");
                            var mre = new System.Threading.ManualResetEvent(false);
                            wc.DownloadProgressChanged += (_, e) => task.Value = e.ProgressPercentage;
                            wc.DownloadFileCompleted += (_, __) => { task.Value = 100; mre.Set(); };
                            wc.DownloadFileAsync(new Uri(downloadUrl), newPath);
                            mre.WaitOne();
                        });
                    }

                    // 4) Updater-Skript schreiben und starten
                    string updaterScript = Path.Combine(Path.GetTempPath(), "bkb-tool-selfupdate.sh");
                    string script = $@"#!/usr/bin/env bash
set -euo pipefail

APP=""{appImagePath}""
NEW=""{newPath}""
DIR=""$(dirname ""$APP"")""
LOG=""$DIR/BKB-Tool-update.log""

echo ""[$(date)] Selfupdate gestartet. APP=$APP NEW=$NEW"" >>""$LOG""

# Kurz warten, bis die App beendet ist
sleep 1

# Warten bis Prozess beendet
n=0
while pgrep -f ""BKB-Tool.*AppImage"" >/dev/null 2>&1; do
  sleep 1
  n=$((n+1))
  if [ $n -gt 120 ]; then
    echo ""Timeout: BKB-Tool beendet sich nicht."" | tee -a ""$LOG""
    exit 1
  fi
done

if [ ! -f ""$NEW"" ]; then
  echo ""Fehler: Update-Datei fehlt: $NEW"" | tee -a ""$LOG""
  exit 1
fi

# Ersetzen
mv -f ""$NEW"" ""$APP""
chmod +x ""$APP""

echo ""Starte neue Version..."" | tee -a ""$LOG""
nohup ""$APP"" >/dev/null 2>&1 &

# Skript entfernt sich selbst
rm -- ""$0""
";

                    // LF + UTF8 ohne BOM schreiben
                    File.WriteAllText(updaterScript, script.Replace("\r\n", "\n"), new System.Text.UTF8Encoding(false));

                    // Ausführbar machen
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "chmod",
                        Arguments = $"+x \"{updaterScript}\"",
                        UseShellExecute = false
                    })?.WaitForExit();

                    // In Terminal starten (wenn vorhanden), sonst direkt
                    string? terminal = null;
                    if (File.Exists("/usr/bin/gnome-terminal")) terminal = "/usr/bin/gnome-terminal";
                    else if (File.Exists("/usr/bin/konsole")) terminal = "/usr/bin/konsole";
                    else if (File.Exists("/usr/bin/xterm")) terminal = "/usr/bin/xterm";

                    if (terminal != null)
                    {
                        var args = terminal.Contains("gnome-terminal")
                            ? $"-- bash -c '\"{updaterScript}\"; read -n1 -s -p \"Update abgeschlossen. Taste drücken...\"'"
                            : terminal.Contains("konsole")
                                ? $"--noclose -e bash -c '\"{updaterScript}\"'"
                                : $"-hold -e bash -c '\"{updaterScript}\"'";

                        Process.Start(new ProcessStartInfo
                        {
                            FileName = terminal,
                            Arguments = args,
                            UseShellExecute = false
                        });
                    }
                    else
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "bash",
                            Arguments = $"\"{updaterScript}\"",
                            UseShellExecute = false
                        });
                    }

                    // Hauptprozess beenden, damit das Skript ersetzen kann
                    Environment.Exit(0);
                    return configuration; // unreachable
                }

                // Windows: unverändert (Download + .bat Updater)
                string downloadUrlWin = null;
                foreach (var asset in selectedRelease.Value.GetProperty("assets").EnumerateArray())
                {
                    var name = asset.GetProperty("name").GetString();
                    if (name != null && name.Equals("BKB-Tool.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        downloadUrlWin = asset.GetProperty("browser_download_url").GetString();
                        break;
                    }
                }

                if (string.IsNullOrEmpty(downloadUrlWin))
                {
                    AnsiConsole.MarkupLine("[red]Fehler: BKB-Tool.exe im Release nicht gefunden.[/]");
                    return configuration;
                }

                string zielDatei = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu.exe");

                using (var webClient = new System.Net.WebClient())
                {
                    webClient.Headers.Add("User-Agent", "request");
                    AnsiConsole.Progress().Start(ctx =>
                    {
                        var task = ctx.AddTask("Lade Update herunter");
                        var mre = new System.Threading.ManualResetEvent(false);
                        webClient.DownloadProgressChanged += (_, e) => task.Value = e.ProgressPercentage;
                        webClient.DownloadFileCompleted += (_, __) => { task.Value = 100; mre.Set(); };
                        webClient.DownloadFileAsync(new Uri(downloadUrlWin), zielDatei);
                        mre.WaitOne();
                    });
                }

                string updaterPath = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool-autoupdater.bat");
                if (File.Exists(updaterPath)) File.Delete(updaterPath);

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
                    "exit\n");

                var panelWin = new Panel($"Mit [{Global.GetColor(Global.ColorActionInMenüs)} bold]ENTER[/] wird jetzt in die Version {githubVer} neugestartet.")
                    .Header("[bold green]  Update erfolgreich  [/]")
                    .BorderColor(Global.ColorActionInMenüs)
                    .Expand();
                AnsiConsole.Write(panelWin);

                while (Console.KeyAvailable) Console.ReadKey(true);
                Console.ReadKey();

                Process.Start(new ProcessStartInfo { FileName = updaterPath, UseShellExecute = true, CreateNoWindow = true });
                Environment.Exit(0);
            }
        }
        else
        {
            string updaterPath = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool-autoupdater.bat");
            if (File.Exists(updaterPath)) File.Delete(updaterPath);
        }
    }
    catch (Exception ex)
    {
        AnsiConsole.MarkupLine($"[bold red]Error checking update: {ex.Message}[/]");
    }

    return configuration;
}