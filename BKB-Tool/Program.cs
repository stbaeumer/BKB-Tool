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
        var client = new System.Net.WebClient();
        client.Headers.Add("User-Agent", "request");
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

        if (Version.TryParse(githubVersionClean, out var githubVer) && Version.TryParse(lokaleVersionClean, out var lokalVer))
        {
            if (true)
            //if (githubVer > lokalVer)
            {
                Global.DisplayHeader(configuration);

                var panelUpdate = new Panel(new Markup($"Ein Update auf {os}-Version [tan]{githubVersion}[/] ist verfügbar. Drücken Sie eine [{Global.GetColor(Global.ColorActionInMenüs)} bold]beliebige Taste[/], um das Update zu starten."))
                    .BorderStyle(new Style(Color.Red))
                    .Expand();
                AnsiConsole.Write(panelUpdate);

                Console.ReadKey();

                // Linux: AppImageUpdate verwenden
                if (os == "Linux")
                {
                    string appImagePath = Environment.GetEnvironmentVariable("APPIMAGE");
                    
                    if (string.IsNullOrEmpty(appImagePath))
                    {
                        AnsiConsole.MarkupLine("[red]Fehler: Keine AppImage-Umgebung erkannt.[/]");
                        return configuration;
                    }

                    // Prüfen ob AppImageUpdate verfügbar ist
                    bool hasAppImageUpdate = false;
                    try
                    {
                        var checkProcess = Process.Start(new ProcessStartInfo
                        {
                            FileName = "which",
                            Arguments = "AppImageUpdate",
                            UseShellExecute = false,
                            RedirectStandardOutput = true
                        });
                        checkProcess?.WaitForExit();
                        hasAppImageUpdate = checkProcess?.ExitCode == 0;
                    }
                    catch { }

                    if (!hasAppImageUpdate)
                    {
                        var panelAppImageUpdate = new Panel(new Markup(
                            $"[red]AppImageUpdate ist nicht installiert.[/]\n\n" +
                            $"Installation:\n" +
                            $"[{Global.GetColor(Global.ColorPfadInDateien)}]sudo apt install appimageupdatetool[/]\n\n" +
                            $"oder über [link=https://flathub.org/apps/de.toolgear.GearLever]Gear Lever[/]"))
                            .Header("[red]AppImageUpdate fehlt[/]")
                            .BorderColor(Color.Red)
                            .Expand();
                        AnsiConsole.Write(panelAppImageUpdate);
                        Console.ReadKey();
                        return configuration;
                    }

                    // Updater-Skript erstellen
                    string updaterScript = Path.Combine(Path.GetTempPath(), "bkb-tool-update.sh");
                    string script = $@"#!/bin/bash
set -e
LOG=""{Path.Combine(Path.GetDirectoryName(appImagePath) ?? "/tmp", "BKB-Tool-update.log")}""

echo ""[$(date)] Update gestartet"" >> ""$LOG""
echo ""Warte auf Beenden von BKB-Tool...""

# Warte bis Prozess beendet ist
while pgrep -f ""BKB-Tool.*AppImage"" >/dev/null 2>&1; do
  sleep 1
done

echo ""Starte AppImageUpdate..."" | tee -a ""$LOG""
cd ""$(dirname ""{appImagePath}"")""

# AppImageUpdate ausführen
if AppImageUpdate ""{appImagePath}"" 2>&1 | tee -a ""$LOG""; then
  # Update erfolgreich: .new → alte Datei ersetzen
  if [[ -f ""{appImagePath}.new"" ]]; then
    mv -f ""{appImagePath}.new"" ""{appImagePath}""
    chmod +x ""{appImagePath}""
    echo ""Update erfolgreich installiert."" | tee -a ""$LOG""
    
    echo ""Starte neue Version...""
    nohup ""{appImagePath}"" >/dev/null 2>&1 &
  else
    echo ""Fehler: {appImagePath}.new nicht gefunden."" | tee -a ""$LOG""
  fi
else
  echo ""AppImageUpdate fehlgeschlagen."" | tee -a ""$LOG""
fi

# Skript löscht sich selbst
rm -- ""$0""
";

                    File.WriteAllText(updaterScript, script.Replace("\r\n", "\n"), new System.Text.UTF8Encoding(false));
                    
                    // Ausführbar machen
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "chmod",
                        Arguments = $"+x \"{updaterScript}\"",
                        UseShellExecute = false
                    })?.WaitForExit();

                    // Updater in neuem Terminal starten
                    string? terminal = null;
                    if (File.Exists("/usr/bin/gnome-terminal")) terminal = "/usr/bin/gnome-terminal";
                    else if (File.Exists("/usr/bin/konsole")) terminal = "/usr/bin/konsole";
                    else if (File.Exists("/usr/bin/xterm")) terminal = "/usr/bin/xterm";

                    if (terminal != null)
                    {
                        var args = terminal.Contains("gnome-terminal")
                            ? $"-- bash -c '\"{updaterScript}\"; read -n1 -s -p \"Update abgeschlossen. Drücken Sie eine Taste...\"'"
                            : terminal.Contains("konsole")
                                ? $"--noclose -e bash -c '\"{updaterScript}\"'"
                                : $"-hold -e bash -c '\"{updaterScript}\"'";

                        Process.Start(new ProcessStartInfo
                        {
                            FileName = terminal,
                            Arguments = args,
                            UseShellExecute = false
                        });

                        AnsiConsole.MarkupLine("[yellow]Update läuft in neuem Terminal...[/]");
                        System.Threading.Thread.Sleep(1000);
                        Environment.Exit(0);
                    }
                    else
                    {
                        // Fallback: Hintergrund
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "bash",
                            Arguments = updaterScript,
                            UseShellExecute = false
                        });
                        Environment.Exit(0);
                    }

                    return configuration;
                }

                // Windows: Bisheriger Code
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

                if (string.IsNullOrEmpty(downloadUrl))
                {
                    AnsiConsole.MarkupLine($"[red]Fehler: BKB-Tool.exe im Release nicht gefunden.[/]");
                    return configuration;
                }

                string zielDatei = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool_neu.exe");

                using (var webClient = new System.Net.WebClient())
                {
                    webClient.Headers.Add("User-Agent", "request");
                    AnsiConsole.Progress()
                        .Start(ctx =>
                        {
                            var task = ctx.AddTask("Lade Update herunter ");
                            webClient.DownloadProgressChanged += (s, e) => task.Value = e.ProgressPercentage;
                            var downloadCompleted = new System.Threading.ManualResetEvent(false);
                            webClient.DownloadFileCompleted += (s, e) => { task.Value = 100; downloadCompleted.Set(); };
                            webClient.DownloadFileAsync(new Uri(downloadUrl), zielDatei);
                            downloadCompleted.WaitOne();
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

                var panel = new Panel($"Mit [{Global.GetColor(Global.ColorActionInMenüs)} bold]ENTER[/] wird jetzt in die Version v{githubVer} neugestartet.")
                    .Header("[bold green]  Update erfolgreich  [/]")
                    .BorderColor(Global.ColorActionInMenüs)
                    .Expand();
                AnsiConsole.Write(panel);

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