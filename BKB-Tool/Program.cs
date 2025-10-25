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

// ...existing code...
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

                    // Updater-Skript schreiben (sichtbar, bleibt erhalten)
                    string updaterScript = Path.Combine(appDir, "BKB-Tool-update.sh");
                    string script = "#!/usr/bin/env bash\n" +
                    "set -euo pipefail\n" +
                    "trap 'echo; echo \"Es ist ein Fehler aufgetreten.\"; read -n1 -s -r -p \"Taste drücken, um dieses Fenster zu schließen ...\"; echo' ERR\n" +
                    $"APP=\"{appImagePath}\"\n" +
                    $"NEW=\"{newPath}\"\n" +
                    "DIR=\"$(dirname \"$APP\")\"\n" +
                    "LOG=\"$DIR/BKB-Tool-update.log\"\n" +
                    "echo \"[$(date)] Selfupdate gestartet. APP=$APP NEW=$NEW\" | tee -a \"$LOG\"\n" +
                    "sleep 1\n" +
                    "n=0\n" +
                    "while pgrep -f \"BKB-Tool.*AppImage\" >/dev/null 2>&1; do\n" +
                    "  echo \"[$(date)] BKB-Tool läuft noch...\" | tee -a \"$LOG\"\n" +
                    "  sleep 1\n" +
                    "  n=$((n+1))\n" +
                    "  if [ $n -gt 120 ]; then\n" +
                    "    echo \"Timeout: BKB-Tool beendet sich nicht.\" | tee -a \"$LOG\"\n" +
                    "    echo\n" +
                    "    read -n1 -s -r -p \"Taste drücken, um dieses Fenster zu schließen ...\"; echo\n" +
                    "    exit 1\n" +
                    "  fi\n" +
                    "done\n" +
                    "if [ ! -f \"$NEW\" ]; then\n" +
                    "  echo \"Fehler: Update-Datei fehlt: $NEW\" | tee -a \"$LOG\"\n" +
                    "  echo\n" +
                    "  read -n1 -s -r -p \"Taste drücken, um dieses Fenster zu schließen ...\"; echo\n" +
                    "  exit 1\n" +
                    "fi\n" +
                    "echo \"Ersetze alte Version...\" | tee -a \"$LOG\"\n" +
                    "mv -f \"$NEW\" \"$APP\"\n" +
                    "chmod +x \"$APP\"\n" +
                    "echo \"Starte neue Version...\" | tee -a \"$LOG\"\n" +
                    "nohup \"$APP\" >/dev/null 2>&1 &\n" +
                    "echo\n" +
                    "echo \"Update abgeschlossen.\"\n" +
                    "echo \"Log: $LOG\"\n" +
                    "echo\n" +
                    "read -n1 -s -r -p \"Taste drücken, um dieses Fenster zu schließen ...\"; echo\n";

                    File.WriteAllText(updaterScript, script, new System.Text.UTF8Encoding(false));
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "chmod",
                        Arguments = $"+x \"{updaterScript}\"",
                        UseShellExecute = false
                    })?.WaitForExit();

                    AnsiConsole.MarkupLine($"[gray]Updater-Skript:[/] [bold {Global.GetColor(Global.ColorPfadInDateien)}]{updaterScript}[/]");

                    // Terminal-Erkennung: alacritty > gnome-terminal
                    bool hasAlacritty = false;
                    bool hasGnomeTerminal = false;
                    try
                    {
                        var checkAlacritty = Process.Start(new ProcessStartInfo
                        {
                            FileName = "which",
                            Arguments = "alacritty",
                            UseShellExecute = false,
                            RedirectStandardOutput = true,
                            CreateNoWindow = true
                        });
                        checkAlacritty?.WaitForExit();
                        hasAlacritty = checkAlacritty?.ExitCode == 0;

                        if (!hasAlacritty)
                        {
                            var checkGnome = Process.Start(new ProcessStartInfo
                            {
                                FileName = "which",
                                Arguments = "gnome-terminal",
                                UseShellExecute = false,
                                RedirectStandardOutput = true,
                                CreateNoWindow = true
                            });
                            checkGnome?.WaitForExit();
                            hasGnomeTerminal = checkGnome?.ExitCode == 0;
                        }
                    }
                    catch { }

                    if (!hasAlacritty && !hasGnomeTerminal)
                    {
                        var terminalWarning = new Panel(new Markup(
                            $"[red]Weder alacritty noch gnome-terminal ist installiert.[/]\n\n" +
                            $"Das Update-Skript wurde erstellt:\n" +
                            $"[bold {Global.GetColor(Global.ColorPfadInDateien)}]{updaterScript}[/]\n\n" +
                            $"Installation:\n" +
                            $"[{Global.GetColor(Global.ColorPfadInDateien)}]sudo dnf install alacritty[/] (Fedora) oder\n" +
                            $"[{Global.GetColor(Global.ColorPfadInDateien)}]sudo apt install alacritty[/] (Debian/Ubuntu)\n\n" +
                            $"Oder führen Sie das Skript manuell aus:\n" +
                            $"[{Global.GetColor(Global.ColorPfadInDateien)}]bash {updaterScript}[/]"))
                            .Header("[red]Terminal fehlt[/]")
                            .BorderColor(Color.Red)
                            .Expand();
                        AnsiConsole.Write(terminalWarning);
                        Console.ReadKey();
                        return configuration;
                    }

                    try
                    {
                        if (hasAlacritty)
                        {
                            Process.Start("/usr/bin/gnome-terminal");
                            Process.Start("bash", new [] { "-c", "gnome-terminal" });


                            var psi = new ProcessStartInfo
                            {
                                FileName = "/usr/bin/alacritty",
                                UseShellExecute = false,
                                WorkingDirectory = "/home/stefan/"
                            };
                            /*psi.ArgumentList.Add("-t");
                            psi.ArgumentList.Add("BKB-Tool Update");
                            psi.ArgumentList.Add("-e");
                            psi.ArgumentList.Add("bash");
                            psi.ArgumentList.Add(updaterScript);*/
                            Process.Start(psi);
                            psi = new ProcessStartInfo
                            {
                                FileName = "/usr/bin/alacritty",
                                UseShellExecute = true,
                                WorkingDirectory = "/home/stefan/"
                            };
                            /*psi.ArgumentList.Add("-t");
                            psi.ArgumentList.Add("BKB-Tool Update");
                            psi.ArgumentList.Add("-e");
                            psi.ArgumentList.Add("bash");
                            psi.ArgumentList.Add(updaterScript);*/
                            Process.Start(psi);
                        }
                        else
                        {
                            Process.Start("/usr/bin/gnome-terminal");
                            Process.Start("bash", new [] { "-c", "gnome-terminal" });


                            var psi = new ProcessStartInfo
                            {
                                FileName = "gnome-terminal",
                                UseShellExecute = false,
                                WorkingDirectory = "/home/stefan/"
                            };
                            //psi.ArgumentList.Add("--wait");
                            //psi.ArgumentList.Add("--");
                            //psi.ArgumentList.Add("bash");
                            //psi.ArgumentList.Add(updaterScript);
                            Process.Start(psi);
                            psi = new ProcessStartInfo
                            {
                                FileName = "/usr/bin/gnome-terminal",
                                UseShellExecute = true,
                                WorkingDirectory = "/home/stefan/"
                            };
                            /*psi.ArgumentList.Add("-t");
                            psi.ArgumentList.Add("BKB-Tool Update");
                            psi.ArgumentList.Add("-e");
                            psi.ArgumentList.Add("bash");
                            psi.ArgumentList.Add(updaterScript);*/
                            Process.Start(psi);
                        }
                    }
                    catch (Exception e)
                    {
                        AnsiConsole.MarkupLine($"[red]Start des Terminals fehlgeschlagen:[/] {e.Message}");
                        AnsiConsole.MarkupLine($"[gray]Tipp:[/] Versuchen Sie manuell:\n  gnome-terminal --wait -- bash \"{updaterScript}\"");
                    }

                    Environment.Exit(0);
                    return configuration; // unreachable
                }
                else if (os == "Windows")
                {
                    if (selectedRelease is null)
                    {
                        AnsiConsole.MarkupLine("[red]Fehler: Release nicht gefunden.[/]");
                        return configuration;
                    }

                    // Download-URL für Windows finden
                    string downloadUrlWin = null;
                    foreach (var asset in selectedRelease.Value.GetProperty("assets").EnumerateArray())
                    {
                        var name = asset.GetProperty("name").GetString();
                        if (!string.IsNullOrEmpty(name) && name.Equals("BKB-Tool.exe", StringComparison.OrdinalIgnoreCase))
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

                    // Ablageordner: neben der laufenden EXE
                    var exePath = Process.GetCurrentProcess().MainModule?.FileName ?? Assembly.GetExecutingAssembly().Location;
                    var exeDir = Path.GetDirectoryName(exePath) ?? Directory.GetCurrentDirectory();
                    var zielDatei = Path.Combine(exeDir, "BKB-Tool_neu.exe");

                    // Download mit Fortschritt
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

                    // Updater-Batch neben der EXE erstellen
                    string updaterPath = Path.Combine(exeDir, "BKB-Tool-autoupdater.bat");
                    var bat = string.Join("\r\n", new[]
                    {
                        "@echo off",
                        "title BKB-Tool Updater",
                        "echo.",
                        "echo BKB-Tool Updater",
                        "echo =================",
                        "echo Warte auf Beenden von BKB-Tool.exe ...",
                        ":waitforend",
                        "tasklist | find /I \"BKB-Tool.exe\" >nul",
                        "if not errorlevel 1 (",
                        "  timeout /t 1 >nul",
                        "  goto waitforend",
                        ")",
                        "echo Ersetze alte Version ...",
                        "del /F /Q \"BKB-Tool.exe\"",
                        "if exist \"BKB-Tool.exe\" (",
                        "  echo Fehler: Die alte BKB-Tool.exe konnte nicht geloescht werden.",
                        "  echo Bitte alle Instanzen schliessen und erneut versuchen.",
                        "  pause",
                        "  exit /b 1",
                        ")",
                        "rename \"BKB-Tool_neu.exe\" \"BKB-Tool.exe\"",
                        "echo Starte neue Version ...",
                        "start \"\" \"BKB-Tool.exe\"",
                        "echo.",
                        "echo Fertig.",
                        // entfernt: "pause",
                        "exit /b 0"
                    });
                    File.WriteAllText(updaterPath, bat, System.Text.Encoding.ASCII);

                    // Batch in neuem sichtbaren Fenster starten
                    var psiWin = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c start \"BKB-Tool Updater\" \"{updaterPath}\"",
                        UseShellExecute = false,
                        WorkingDirectory = exeDir
                    };
                    Process.Start(psiWin);

                    // Aktuelle App beenden, damit die EXE ersetzt werden kann
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