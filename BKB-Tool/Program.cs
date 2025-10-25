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
                    // Eindeutiger Temp-Dateiname
                    string updaterScript = Path.Combine(Path.GetTempPath(), $"bkb-tool-selfupdate-{Guid.NewGuid():N}.sh");
                    string script = $@"#!/usr/bin/env bash
                    set -euo pipefail
                    # Bei Fehler: sichtbar pausieren
                    trap 'echo; echo ""Es ist ein Fehler aufgetreten.""; read -n1 -s -r -p ""Taste drücken, um dieses Fenster zu schließen ...""; echo' ERR

                    APP=""{appImagePath}""
                    NEW=""{newPath}""
                    DIR=""$(dirname ""$APP"")""
                    LOG=""$DIR/BKB-Tool-update.log""

                    echo ""[$(date)] Selfupdate gestartet. APP=$APP NEW=$NEW"" | tee -a ""$LOG""

                    # Kurz warten, bis die App beendet ist
                    sleep 1

                    # Warten bis Prozess beendet
                    n=0
                    while pgrep -f ""BKB-Tool.*AppImage"" >/dev/null 2>&1; do
                    echo ""[$(date)] BKB-Tool läuft noch..."" | tee -a ""$LOG""
                    sleep 1
                    n=$((n+1))
                    if [ $n -gt 120 ]; then
                        echo ""Timeout: BKB-Tool beendet sich nicht."" | tee -a ""$LOG""
                        echo
                        read -n1 -s -r -p ""Taste drücken, um dieses Fenster zu schließen ...""; echo
                        exit 1
                    fi
                    done

                    if [ ! -f ""$NEW"" ]; then
                    echo ""Fehler: Update-Datei fehlt: $NEW"" | tee -a ""$LOG""
                    echo
                    read -n1 -s -r -p ""Taste drücken, um dieses Fenster zu schließen ...""; echo
                    exit 1
                    fi

                    echo ""Ersetze alte Version..."" | tee -a ""$LOG""
                    mv -f ""$NEW"" ""$APP""
                    chmod +x ""$APP""

                    echo ""Starte neue Version..."" | tee -a ""$LOG""
                    nohup ""$APP"" >/dev/null 2>&1 &

                    echo
                    echo ""Update abgeschlossen.""
                    echo ""Log: $LOG""
                    echo
                    read -n1 -s -r -p ""Taste drücken, um dieses Fenster zu schließen ...""; echo

                    # Skript entfernt sich selbst
                    rm -- ""$0""
                    ";

                    // Skript speichern (UTF-8 ohne BOM) und ausführbar machen
                    File.WriteAllText(updaterScript, script.Replace("\r\n", "\n"), new System.Text.UTF8Encoding(false));
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "chmod",
                        Arguments = $"+x \"{updaterScript}\"",
                        UseShellExecute = false
                    })?.WaitForExit();

                    // In Terminal starten (sichtbar halten)
                    string? terminal = null;
                    string args;
                    if (File.Exists("/usr/bin/gnome-terminal"))
                    {
                        terminal = "/usr/bin/gnome-terminal";
                        args = $"--title=\"BKB-Tool Update\" -- bash -lc '\"{updaterScript}\"'";
                    }
                    else if (File.Exists("/usr/bin/kgx")) // GNOME Console
                    {
                        terminal = "/usr/bin/kgx";
                        args = $"-- bash -lc '\"{updaterScript}\"'";
                    }
                    else if (File.Exists("/usr/bin/konsole"))
                    {
                        terminal = "/usr/bin/konsole";
                        args = $"--new-tab --noclose -e bash -lc '\"{updaterScript}\"'";
                    }
                    else if (File.Exists("/usr/bin/xfce4-terminal"))
                    {
                        terminal = "/usr/bin/xfce4-terminal";
                        args = $"--title=\"BKB-Tool Update\" --hold -e bash -lc '\"{updaterScript}\"'";
                    }
                    else if (File.Exists("/usr/bin/mate-terminal"))
                    {
                        terminal = "/usr/bin/mate-terminal";
                        // FIX: korrekt escapen, keine doppelten "" im C#-String
                        args = $"--title=\"BKB-Tool Update\" -- bash -lc '\"{updaterScript}\"'";
                    }
                    else if (File.Exists("/usr/bin/alacritty"))
                    {
                        terminal = "/usr/bin/alacritty";
                        args = $"-t \"BKB-Tool Update\" -e bash -lc '\"{updaterScript}\"'";
                    }
                    else if (File.Exists("/usr/bin/xterm"))
                    {
                        terminal = "/usr/bin/xterm";
                        args = $"-T \"BKB-Tool Update\" -hold -e bash -lc '\"{updaterScript}\"'";
                    }
                    else if (File.Exists("/usr/bin/x-terminal-emulator"))
                    {
                        terminal = "/usr/bin/x-terminal-emulator";
                        args = $"-e bash -lc '\"{updaterScript}\"'";
                    }
                    else
                    {
                        terminal = null;
                        args = $"\"{updaterScript}\"";
                    }

                    if (terminal != null)
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = terminal,
                            Arguments = args,
                            UseShellExecute = false
                        });
                    }
                    else
                    {
                        // Fallback: ohne eigenes Terminal (unsichtbar)
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "bash",
                            Arguments = args,
                            UseShellExecute = false
                        });
                    }

                    // Hauptprozess beenden, damit das Skript ersetzen kann
                    Environment.Exit(0);
                    return configuration; // unreachable
                }
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