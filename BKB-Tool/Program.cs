// Stefan Bäumer 2025
using System.Reflection;
using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

Global.User = Environment.UserName;
IConfiguration? configuration = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile($"{Global.User}.json", optional: true, reloadOnChange: true).Build();

configuration["AppName"] = "BKB-Tool";
configuration["AppVersion"] = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "0.0.0.0"; // Major.Minor.Build.Revision
configuration["AppDescription"] = "BKB-Tool - Ein Werkzeug an der Schnittstelle zwischen SchILD und Untis.";

Global.DisplayHeader(configuration);

CheckForUpdate(configuration);

Global.PrivilegierteSchulnummern = new List<string>
{
    "177659" // BK Borken
};

do
{
    if (!File.Exists(Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json")))
    {
        configuration = Global.EinstellungenDurchlaufen(Global.Modus.Create);
        Global.DisplayHeader(configuration);
    }
    else
    {
        configuration = Global.EinstellungenDurchlaufen(Global.Modus.Read);
    }

    var table = new Table().Centered();

    var dateien = new Dateien();
    dateien.ExportAusSchildVerschieben(configuration);
    dateien.GetInteressierendeDateienMitAllenEigenschaften(configuration);
    dateien.GetZeilen(configuration);

    var menu = MenueHelper.Einlesen(dateien, configuration);
    if (menu == null) continue;

    var menuGefiltert = new Menue();
    menuGefiltert.AddRange(menu.Where(m => m.BeiDiesenSchulnummernAnzeigen == Global.NurBeiDiesenSchulnummern.Alle || Global.PrivilegierteSchulnummern.Contains(configuration["Schulnummer"])));
    menuGefiltert.AuswahlGridRendern();

    configuration = menuGefiltert.GetAusgewaehlterMenueintrag(configuration, ["x", "y"]);
    var i = Convert.ToInt32(configuration["Auswahl"]);

    Global.DisplayHeader(configuration);
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
        string lokaleVersion = configuration["AppVersion"] ?? "0.1";

        // GitHub API abfragen
        string apiUrl = "https://api.github.com/repos/stbaeumer/BKB-Tool/releases/latest";
        var client = new System.Net.WebClient();
        client.Headers.Add("User-Agent", "request"); // GitHub verlangt einen User-Agent
        string json = client.DownloadString(apiUrl);

        // Version aus JSON extrahieren
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        string githubVersion = doc.RootElement.GetProperty("tag_name").GetString();

        // Vergleich
        if (githubVersion != lokaleVersion)
        {
            var panel2 = new Panel($"Das Update wird jetzt heruntergeladen und der Autoupdater wird gestartet.")
                .Header("[bold lightPink1]  Update verfügbar  [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.LightPink1);

            AnsiConsole.Write(panel2);

            // Die Autoupdater-Batch-Datei wird erzeugt und neben die exe gespeichert. Danach wird sie ausgeführt.
            string updaterPath = Path.Combine(Directory.GetCurrentDirectory(), "BKB-Tool-autoupdater.bat");

            // wenn die Datei schon existiert, dann löschen
            if (File.Exists(updaterPath))
            {
                File.Delete(updaterPath);
            }

            File.WriteAllText(updaterPath,
                "@echo off\n" +
                "echo Update wird heruntergeladen...\n" +
                "curl -L -o BKB-Tool.zip https://github.com/stbaeumer/BKB-Tool/releases/latest/download/BKB-Tool.zip\n" +
                "echo Warte auf Beenden von BKB-Tool.exe ...\n" +
                "taskkill /IM \"BKB-Tool.exe\" /F >nul 2>&1\n" +
                "timeout /t 2 >nul\n" +
                "echo Entpacke Update ...\n" +
                "powershell -Command \"Expand-Archive -Path 'BKB-Tool.zip' -DestinationPath '.' -Force\"\n" +
                "echo Ersetze alte Version ...\n" +
                "move /Y BKB-Tool.exe .\\BKB-Tool_neu.exe >nul\n" +
                "del /F /Q BKB-Tool.exe\n" +
                "rename BKB-Tool_neu.exe BKB-Tool.exe\n" +
                "del /F /Q BKB-Tool.zip\n" +
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

            Console.ReadKey();
        }
    }
    catch (Exception ex)
    {
        AnsiConsole.MarkupLine($"[bold red]Error starting updater: {ex.Message}[/]");
    }
}
