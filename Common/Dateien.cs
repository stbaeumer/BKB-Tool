using System.Configuration;
using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

#pragma warning disable CS8603 // Mögliche Null-Verweis-Rückgabe
#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise null-Objekts.
#pragma warning disable CS8604 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8620 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8600 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8618 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8619 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0219 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8625 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8601 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0168 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0618 // Möglicher Null-Verweis-Argument
#pragma warning disable NU1903 // Möglicher Null-Verweis-Argument
#pragma warning disable NU1902 // Möglicher Null-Verweis-Argument

public class Dateien : List<Datei>
{
    public Dateien()
    {
    }

    public void GetInteressierendeDateienMitAllenEigenschaften(IConfiguration configuration)
    {
        var schildhinweise = new string[]
        {
            "Exportieren Sie alle *.dat-Dateien aus SchILD, indem Sie den Pfad gehen:",
            "1. [bold green]Datenaustausch > Schnittstelle SchILD NRW > Export[/]",
            $"2. [bold green]Arbeitsverzeichnis[/]: [bold blue]{configuration["PfadSchilddatenaustausch"]}[/]",
            "3. [bold green]Export starten[/]"            
        };

        var untishinweise = new string[]
        {
            "Exportieren Sie die Datei aus Untis, indem Sie den Pfad gehen:",
            "1. [bold green]Datei > Import/Export > Export TXT Datei[/]",
            "2. Als Delimiter muss '|' ausgewählt werden.",
            "3. Die Datei auswählen.",
            $"4. Die Datei in [bold blue]{configuration["PfadDownloads"]}[/] speichern."
        };

        Add(new Datei(
            "SchuelerBasisdaten",
            "Beschreibung",
            schildhinweise,
            [""],
            true,           
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchildSchuelerExport",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus SchILD, indem Sie den Pfad gehen:",
                "[bold green]Datenaustausch > Export in Text-/Exceldateien > Exportieren[/]",
                "Die Vorlage 'SchildSchuelerExport' laden.",
                "Export starten.",
                "GGfs. muss die Vorlage erst erstellt werden. Es müssen folgende Felder enthalten sein:",
                "Geburtsdatum, Interne ID-Nummer, Nachname, Vorname, Klasse, Externe ID-Nummer, Status",
                "Delimiter: '|'; Dateiendung: *.dat"
            ],
            [""],
            true,
            d => d.FilternSchildSchuelerExport(),
            "*.txt"
        ));
        Add(new Datei(
            "OpenPeriod",
            "Beschreibung",
            ["Exportieren Sie die Datei aus Webuntis, indem Sie den Pfad gehen:", 
            "[bold green]Klassenbuch > Offene Stunden > Bericht[/]", 
            $"Die PDF-Datei in [bold blue]{configuration["PfadDownloads"]}[/] speichern."],
            [""],
            true,
            d => d.FilterOpenPeriod(),
            "*.pdf"
        ));        
        Add(new Datei(
            "Student_",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie als Administrator den Pfad gehen:",
                " [bold green]Stammdaten > Schüler*innen > Berichte > Schüler > CSV-Ausgabe[/]",
                $" Die Datei in [bold blue]{configuration["PfadDownloads"]}[/] speichern."
            ],
            [""],
            true,
            d => d.FilternWebuntisStudent(),
            "*.csv",
            "\t"
        ));        
        Add(new Datei(
            "GPU003",
            "Beschreibung",
            untishinweise,
            [""],
            false,
            d => d.FilternKlassenGPU003(),
            "*.TXT",
            "|"
        ));
        Add(new Datei(
            "GPU004",
            "Beschreibung: Lehrkraefte",
            untishinweise,
            [""],
            false,
            d => d.FilternLehrkraefteGPU004(),
            "*.TXT",
            "|"
        ));
        Add(new Datei(
            "GPU020",
            "Beschreibung",
            untishinweise,
            [""],
            false,
            d => d.FilternGPU020(),
            "*.TXT",
            "|",
            true
        ));
        Add(new Datei(
            "Lehrkraefte.dat",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilterLehrkraefte()
        ));
        Add(new Datei(
            "LehrkraefteSonderzeiten.dat",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilterLehrkraefte(),
            "*.dat",
            "|",
            true
        ));
        Add(new Datei(
            @"Gost.csv",
            "Beschreibung",
            [""],
            [""],
            false,
            d => d.FilterGost(),
            "*.csv",
            ","
        ));
        Add(new Datei(
            "ExportLessons",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilterExportLessons(),
            "*.csv",
            "\t"
        ));        
        Add(new Datei(
            "Kurse.",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilterKurse()
        ));
        Add(new Datei(
            @"DatumsAusAtlantis.csv",
            "Beschreibung",
            [],
            [""],
            true,
            d => d.FilterDatumsAusAtlantis(),
            "*.csv",
            ","
        ));
        Add(new Datei(
            @"Adressen.csv",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Outlook, indem Sie:",
                "den Kalender in Listenansicht anzeigen",
                "Mit Strg+A alles markieren",
                "Mit Strg+C kopieren",
                "Die Datei " + Path.Combine(configuration["PfadDownloads"],"termine_kollegium.csv") + " überschreiben oder neu anlegen."
            ],
            [""],
            true,
            d => d.FilternAdressenAtlantis(),
            "*.csv",
            ";"
        ));
        Add(new Datei(
            @"termine_kollegium.csv",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Outlook, indem Sie:",
                "den Kalender in Listenansicht anzeigen",
                "Mit Strg+A alles markieren",
                "Mit Strg+C kopieren",
                "Die Datei " + Path.Combine(configuration["PfadDownloads"],"termine_kollegium.csv") + " überschreiben oder neu anlegen."
            ],
            [""],
            true,
            d => d.FilternTermineKollegium(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"termine_fhr.csv",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Outlook, indem Sie:",
                "den Kalender in Listenansicht anzeigen",
                "Mit Strg+A alles markieren",
                "Mit Strg+C kopieren",
                "Die Datei " + Path.Combine(configuration["PfadDownloads"],"termine_fhr.csv") + " überschreiben oder neu anlegen."
            ],
            [""],
            true,
            d => d.FilternTermineFhr(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"termine_verwaltung.csv",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Outlook, indem Sie:",
                "den Kalender in Listenansicht anzeigen",
                "Mit Strg+A alles markieren",
                "Mit Strg+C kopieren",
                "Die Datei " + Path.Combine(configuration["PfadDownloads"],"termine_verwaltung.csv") + " überschreiben oder neu anlegen."
            ],
            [""],
            true,
            d => d.FilternTermineVerwaltung(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"termine_berufliches_gymnasium.csv",
            "Beschreibung",
             [
                "Exportieren Sie die Datei aus Outlook, indem Sie:",
                "den Kalender in Listenansicht anzeigen",
                "Mit Strg+A alles markieren",
                "Mit Strg+C kopieren",
                "Die Datei " + Path.Combine(configuration["PfadDownloads"],"termine_berufliches_gymnasium.csv") + " überschreiben oder neu anlegen."
            ],
            [""],
            true,
            d => d.FilternTermineBeruflichesGymnasium(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            @"Atlantis-Zeugnisse-Noten.csv",
            "Beschreibung",
            [""],
            [""],
            false,
            d => d.FilternAtlantisZeugnisseNoten(),
            "*.csv",
            ","
        ));
        Add(new Datei(
            "SchuelerLeistungsdaten",
            "Beschreibung",
            schildhinweise,
            ["Mahnung", "Mahndatum", "Sortierung"],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerErzieher",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerTelefonnummern",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerAdressen",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "Klassen",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternSchildKlassen()
        ));        
        Add(new Datei(
            "SchuelerZusatzdaten",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "Faecher",
            "Beschreibung",
            schildhinweise,
            ["Unterrichtsprache", "Sortierung S1", "Sortierung S2", "Gewichtung"],
            true,
            d => d.FilternSchildFaecher()
        ));
        Add(new Datei(
            "SchuelerLernabschnittsdaten",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerTeilleistungen",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerFehlstunden",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "SchuelerVermerke",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternDatDatei()
        ));
        Add(new Datei(
            "MarksPerLesson",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternMarksPerLessons(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "StudentgroupStudents",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternStudentgroupStudents(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "AbsencePerStudent",
            "Beschreibung",
            schildhinweise,
            [""],
            true,
            d => d.FilternAbsencePerLEssons(),
            "*.csv",
            "\t"
        ));
    }

    public List<dynamic>? GetMatchingList(IConfiguration configuration, string pattern, Students students = null, Klassen klassen = null)
    {   
        var datei = this.FirstOrDefault(datei => !string.IsNullOrEmpty(datei.Dateiname) && datei.Dateiname.ToLower().Contains(pattern, StringComparison.CurrentCultureIgnoreCase));

        // Mögliche Meldungen werden ausgegeben, wenn die Datei nicht gefunden wurde oder veraltet ist.

        var meldung = "";
        // if (datei.IstVeraltet(configuration))

        if (string.IsNullOrEmpty(datei.AbsoluterPfad) && datei.Endung.ToLower().Contains("dat"))
            meldung += $"Die Datei [bold blue]SchuelerBasisdaten.dat[/] wurde weder in [bold blue]{configuration["pfadDownloads"]}[/] noch in [bold blue]{configuration["pfadSchilddatenaustausch"]}[/] gefunden. Am besten gehen Sie jetzt in SchILD zu [bold green]Datenaustausch > Schnittstelle SchILD NRW > Export[/] und klicken [bold green]Export starten[/], um alle Dateien nach [bold blue]{configuration["pfadSchilddatenaustausch"]}[/] zu exportieren. Anschließend kehren Sie hierher zurück.";
        else if (datei.AbsoluterPfad == null)
            meldung += $"Die Datei [bold blue]{pattern}[/] wurde nicht gefunden. Bitte prüfen Sie, ob sie im Ordner {configuration["PfadDownloads"]} vorhanden ist.";
        else if (datei.IstOptional && datei.Count == 0)
            meldung = $"Die Datei [bold blue]{pattern}[/] ist optional und wurde nicht gefunden. Sie wird nicht benötigt.";
        else if (datei.IstOptional && datei.Count > 0)
            return datei.ToList();
        else if (datei.AbsoluterPfad == null || datei.AbsoluterPfad.Length == 0 && !datei.IstOptional)
            meldung = $"Die Datei [bold red]{pattern}[/] existiert nicht. Bitte prüfen Sie, ob sie im Ordner {configuration["PfadDownloads"]} vorhanden ist.";        

        datei.FehlermeldungRendern(meldung, configuration);

        // Rückgabe der Datei, wenn sie gefunden wurde und die Bedingungen erfüllt sind.

        if (!datei.IstOptional && (students == null || students.Count == 0) && datei.IstVeraltet(configuration))
            return [];
        else if (datei.AbsoluterPfad.ToLower().Contains("schuelerbasisdaten") && (students == null || students.Count == 0))
            return datei.ToList();
        else if (datei.AbsoluterPfad.ToLower().Contains("schildschuelerexport") && (students == null || students.Count == 0))
            return datei.ToList();
        else if (datei.IstOptional)
            return datei.ToList();
        else if (students != null && students.Count > 0)
            return datei.Filtern(students, students.GetKlassen(klassen));
        else if (datei.AbsoluterPfad.ToLower().Contains("lehrkraefte"))
            return datei.ToList();
        
        return [];
    }
    
    public Dateien Notwendige(IConfiguration configuration, List<string> dateinamenNotwendigeDateien)
    {
        var pfadDownloads = configuration["PfadDownloads"];
        int maxDateiAlter = int.TryParse(configuration["MaxDateiAlter"], out var parsedMaxDateiAlter) ? parsedMaxDateiAlter : 0;

        Dateien notwendige = new Dateien();

        foreach (var dateinameNotwendig in dateinamenNotwendigeDateien)
        {
            try
            {
                var dateiendung = dateinameNotwendig.Split(',')[1].Trim().ToLower();
                var dateiname = dateinameNotwendig.Split(',')[0];

                var datei = this.First(datei => !string.IsNullOrEmpty(datei.Dateiname)
                        && !datei.Dateiname.ToLower().Contains("-kennwort")
                        && datei.Dateiname.ToLower().StartsWith(dateiname.ToLower(), StringComparison.CurrentCultureIgnoreCase)
                        && datei.Endung.ToLower().Contains("*." + dateiendung.ToLower(), StringComparison.CurrentCultureIgnoreCase));

                datei.IstOptional = dateinameNotwendig.Split(',').Length > 2 && dateinameNotwendig.Split(',')[2].ToLower().Contains("opt") ? true : false;
                
                var absoluterPfad = this.First(datei => !string.IsNullOrEmpty(datei.Dateiname)
                && !datei.Dateiname.ToLower().Contains("-kennwort")
                && datei.Dateiname.ToLower().StartsWith(dateiname.ToLower(), StringComparison.CurrentCultureIgnoreCase)
                && datei.Endung.ToLower().Contains("*." + dateiendung.ToLower(), StringComparison.CurrentCultureIgnoreCase)
                ).AbsoluterPfad;

                if (absoluterPfad.ToLower().Contains("-kennwort"))
                {
                    absoluterPfad = "";
                }

                if (absoluterPfad.Length > 0)
                {
                    if (datei.Count == 0)
                    {
                        if (datei.AbsoluterPfad.EndsWith(".pdf"))
                        {
                            if (datei.Erstelldatum.Date.AddDays(maxDateiAlter) < DateTime.Now)
                            {
                                datei.FehlermeldungRendern($"Die Datei {absoluterPfad} existiert, ist aber veraltet.", configuration);
                            }
                        }
                        else
                        {
                            if (!datei.DarfLeerSein)
                            {
                                datei.FehlermeldungRendern($"Die Datei {absoluterPfad} existiert, ist aber leer. Ist die Datei evtl. vorher in Excel o.ä. geöffnet worden? Oder stimmt der Delimiter nicht? Der korrekte Delimiter ist: '[bold blue]{datei.Delimiter}[/]'", configuration);
                            }
                        }
                    }
                    else
                    {
                        if (((IDictionary<string, object>)datei[0]).Count == 1)
                            datei.FehlermeldungRendern($"Die Datei {absoluterPfad} hat nur eine einzige Spalte. Das korrekte Trennzeichen ist: {datei.Delimiter}'.", configuration);

                        if (datei.Erstelldatum.Date.AddDays(maxDateiAlter) < DateTime.Now.Date)
                        {
                            var veraltet = $"darf aber nicht älter als [bold red]{configuration["MaxDateiAlter"]}[/] Tage sein.";
                            if (configuration["MaxDateiAlter"] == "1")
                                veraltet = "darf aber nicht älter als von gestern sein.";
                            if (configuration["MaxDateiAlter"] == "0")
                                veraltet = "muss aber von heute sein.";
                            datei.FehlermeldungRendern($"Die Datei [bold red]{datei.AbsoluterPfad}[/] ist veraltet. Sie wurde am [bold red]{datei.Erstelldatum:dd.MM.yyyy}[/] erstellt, {veraltet}", configuration);
                        }
                    }
                }
                else
                {
                    var opt = "";

                    if (datei.IstOptional)
                    {
                        opt = ", ist aber optional";
                    }
                    else
                    {
                        opt = " und ist nicht optional";
                    }
                    
                    datei.FehlermeldungRendern($"Die Datei [bold red]{Path.Combine(pfadDownloads, dateiname)}***[/] existiert nicht{opt}.",configuration);
                }
                notwendige.Add(datei);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        return notwendige;
    }

    public void GetZeilen(IConfiguration configuration, List<string>? dateienImPfad = null)
    {
        var maxDateiAlterString = configuration["maxDateiAlter"];
        int maxDateiAlter = 6; // Standardwert

        if (!string.IsNullOrEmpty(maxDateiAlterString) && int.TryParse(maxDateiAlterString, out int parsedValue))
        {
            maxDateiAlter = parsedValue;
        }

        if (dateienImPfad == null)
        {
            dateienImPfad = GetDateienImPfad(configuration);
        }

        int anzahlDateienMitZeilen = 0;

        foreach (var datei in this)
        {
            if (dateienImPfad.Any(d => Path.GetFileName(d).ToLower().StartsWith(datei.Dateiname.ToLower())))
            {
                var passendeDatei = dateienImPfad
                    .OrderBy(d => File.GetLastWriteTime(d))
                    .LastOrDefault(d => Path.GetFileName(d).StartsWith(datei.Dateiname));

                if (passendeDatei != null)
                {
                    datei.AbsoluterPfad = passendeDatei;
                    datei.Erstelldatum = File.GetLastWriteTime(datei.AbsoluterPfad);
                    datei.GetZeilen();

                    if (datei.Erstelldatum.AddDays(maxDateiAlter) <= DateTime.Now)
                    {
                        Global.Zeilen.Add(new ValueTuple<string, ConsoleColor>(datei.AbsoluterPfad + ": Datei veraltet", ConsoleColor.Red));
                    }
                    else
                    {
                        if (datei.Count > 0)
                        {
                            anzahlDateienMitZeilen++;
                        }
                    }
                }
            }
        }
        if (anzahlDateienMitZeilen == 0)
        {
            Global.ZeileSchreiben($"Keine Dateien mit Zeilen in [bold red]{configuration["pfadDownloads"]}[/] gefunden.", "0", ConsoleColor.Red, ConsoleColor.White);
        }
        else
        {
            AnsiConsole.Write(new Rule($"[bold fuchsia] {anzahlDateienMitZeilen} Dateien eingelesen aus {configuration["PfadDownloads"]} [/] ").RuleStyle("fuchsia").LeftJustified());
        }
    }

    public List<string> GetKlassen()
    {
        var klassenSet = new HashSet<string>(); // HashSet für einzigartige Werte

        foreach (var datei in this.Where(x => x.Count > 0))
        {
            foreach (var record in datei) // Direkt durch die bereits geladenen Records iterieren
            {
                var dict = (IDictionary<string, object>)record; // FastDynamicObject als Dictionary behandeln

                // Suche nach einer passenden Spalte
                var key = dict.Keys.FirstOrDefault(k =>
                    k.Equals("Klasse", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("klasse", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Klassen", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("klassen", StringComparison.OrdinalIgnoreCase) ||
                    k.Equals("Class", StringComparison.OrdinalIgnoreCase));

                if (key != null && dict[key] != null)
                {
                    var value = dict[key].ToString().Trim();

                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        // Falls mehrere Klassen mit "~" getrennt sind, aufsplitten
                        var klassenArray = value.Split('~', StringSplitOptions.RemoveEmptyEntries);

                        foreach (var klasse in klassenArray)
                        {
                            klassenSet.Add(klasse.Trim()); // Trim, um unnötige Leerzeichen zu entfernen
                        }
                    }
                }
            }
        }

        // Sortierte Liste erstellen und "alle" an den Anfang setzen
        var klassenListe = klassenSet.Where(x => !string.IsNullOrWhiteSpace(x))
            .OrderBy(x => x)
            .ToList();

        klassenListe.Insert(0, "Für alle Klassen oder "); // "alle" an erster Stelle einfügen

        return klassenListe;
    }

    /// <summary>
    /// Holt alle Dateien im Download-Pfad, die den Kriterien entsprechen.
    /// Die Kriterien sind: Dateiendung .csv, .txt, .dat oder Dateiname beginnt mit "openperiod".
    /// </summary>
    /// <param name="configuration"></param>
    /// <returns></returns>
    public List<string> GetDateienImPfad(IConfiguration configuration)
    {
        return Directory.GetFiles(configuration["PfadDownloads"], "*", SearchOption.TopDirectoryOnly) // Nur im aktuellen Verzeichnis suchen
        .Where(f => f.ToLower().EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ||
                    f.ToLower().EndsWith(".txt", StringComparison.OrdinalIgnoreCase) ||
                    f.ToLower().EndsWith(".dat", StringComparison.OrdinalIgnoreCase) ||
                    Path.GetFileNameWithoutExtension(f).ToLower().StartsWith("openperiod", StringComparison.OrdinalIgnoreCase))
        .OrderBy(f => File.GetLastWriteTime(f)) // Nach Erstellungsdatum sortieren
        .ToList();
    }

    /// <summary>
    /// Die Exportdateien aus SchILD werden zu den anderen heruntergeladenen Dateien verschoben, um Platz zu machen für neue
    /// Dateien, die dann nach SchILD importiert werden. 
    /// </summary>
    public void ExportAusSchildVerschieben(IConfiguration configuration)
    {
        var pfadSchilddatenaustausch = configuration["PfadSchilddatenaustausch"];
        var pfadDownloads = configuration["PfadDownloads"];

        // Stelle sicher, dass der Zielordner existiert
        if (!Directory.Exists(pfadSchilddatenaustausch))
        {
            try
            {
                Directory.CreateDirectory(pfadSchilddatenaustausch);
            }
            catch (Exception ex)
            {
                AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
                return;
            }
        }

        // Die SchildSchuelerExport wird immer kopiert.
        var datei = !string.IsNullOrEmpty(pfadSchilddatenaustausch)
            ? Directory.GetFiles(pfadSchilddatenaustausch, "*", SearchOption.TopDirectoryOnly).FirstOrDefault(f => Path.GetFileName(f).ToLower().Contains("schildschuelerexport"))
            : null;

        if (datei != null)
        {
            if (pfadDownloads == null)
                throw new ArgumentNullException(nameof(pfadDownloads), "Der Pfad zum Downloads-Ordner ist null.");

            var destinationPath = Path.Combine(pfadDownloads, Path.GetFileName(datei));
            File.Copy(datei, destinationPath, true);
        }

        var datFiles = !string.IsNullOrEmpty(pfadSchilddatenaustausch)
            ? Directory.GetFiles(pfadSchilddatenaustausch, "*.dat").ToList()
            : new List<string>();

        if (datFiles.Count > 5)
        {
            // Hole die Erstellungszeiten der Dateien
            var creationTimes = datFiles
                .Select(file => File.GetLastWriteTime(file))
                .OrderBy(time => time)
                .ToList();

            // Prüfe, ob alle Dateien innerhalb von vier Minuten erstellt wurden. Daran wird erkannt, ob es sich um einen Export aus SchILD handelt.
            var within4Minutes = creationTimes.Last() - creationTimes.First() <= TimeSpan.FromMinutes(4);

            if (within4Minutes)
            {
                var existingDatFiles = !string.IsNullOrEmpty(pfadDownloads)
                    ? Directory.GetFiles(pfadDownloads, "*.dat")
                    : Array.Empty<string>();

                // Panel: Wenn keine .dat-Dateien vorhanden sind, wird eine Warnung ausgegeben.
    
        var panel = new Panel(
            $"SchILD verwendet für den Export und den (Re-)Import denselben Ordner. Deswegen verschiebt {configuration["AppName"]} jetzt " +
            $"[bold fuchsia]{datFiles.Count} aus SchILD exportierte *.dat-Dateien[/] direkt von [bold fuchsia]{pfadSchilddatenaustausch}[/] nach [bold fuchsia]{pfadDownloads}[/]. " +
            $"Die aufbereiteten Dateien stellt {configuration["AppName"]} wiederum in [bold fuchsia]{pfadSchilddatenaustausch}[/] bereit. " +
            "So bleiben die Import-Dateien und Export-Dateien stets getrennt voneinander."
            )
            .Header("[bold fuchsia]  Hinweis  [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Color.Fuchsia);

        AnsiConsole.Write(panel);

                // Lösche die vorhandenen .dat-Dateien im Zielordner
                foreach (var file in existingDatFiles)
                {
                    File.Delete(file);
                }

                // Verschiebe die Dateien
                foreach (var file in datFiles)
                {
#pragma warning disable CS8604 // Possible null reference argument.
                    var destinationPath = Path.Combine(pfadDownloads, Path.GetFileName(file));
#pragma warning restore CS8604 // Possible null reference argument.
                    
                    File.Move(file, destinationPath);
                }

                AnsiConsole.Write(new Rule($"[bold fuchsia] {datFiles.Count} Dateien verschoben von {configuration["PfadSchilddatenaustausch"]} nach {configuration["PfadDownloads"]}[/] ").RuleStyle("fuchsia").LeftJustified());
            }
            else
            {
                Console.WriteLine("Die Dateien wurden nicht innerhalb von einer Minute erstellt.");
            }
        }
        else
        {
            foreach (var file in datFiles)
            {
                File.Delete(file);
            }
        }
    }

    public void FehlermeldungRendern(IConfiguration configuration)
    {
        foreach (var datei in this.Where(q => !string.IsNullOrEmpty(q.Fehlermeldung)))
        {
            datei.FehlermeldungRendern("",configuration);
        }
    }
}