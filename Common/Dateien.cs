using System.Configuration;
using System.Diagnostics;
using System.Text.Json;
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
#pragma warning disable CS0252 // Unbeabsichtigter Verweisvergleich. Wandeln Sie die linke Seite in den Typ "string" um, um einen Wertvergleich durchzuführen.
#pragma warning disable CA2200
#pragma warning disable CS8618
#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise nullverweisenden Objekts.
#pragma warning disable CS8604 // Möglicherweise wird ein nullverweisendes Argument an eine nicht-nullbare Parameterreferenz übergeben.
#pragma warning disable CS8073
#pragma warning disable CS0472
#pragma warning disable CS8625
#pragma warning disable CS8600
#pragma warning disable CS0162

public class Dateien : List<Datei>
{
    public List<string> Meldung { get; private set; }

    public Dateien()
    {
        Meldung = new List<string>();
    }

    public Dateien(IConfiguration configuration)
    {
        Meldung = new List<string>();        
    }

    public void GetInteressierendeDateienMitAllenEigenschaften(IConfiguration configuration)
    {
        configuration = Global.Konfig("PfadDownloads", Global.Modus.ReadSilent, configuration);
        configuration = Global.Konfig("PfadSchilddatenaustausch", Global.Modus.ReadSilent, configuration);
    
        var schildhinweise = new string[]
        {
            "Exportieren Sie alle *.dat-Dateien aus SchILD, indem Sie den Pfad gehen:",
            $"1. [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Datenaustausch > Schnittstelle SchILD NRW > Export[/]",
            $"2. [bold silver]Ausgabeverzeichnis[/]: [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadSchilddatenaustausch"]}[/]",
            $"3. [bold {Global.GetColor(Global.ColorActionInMenüs)}]Export starten[/]"
        };

        var untishinweise = new string[]
        {
            "Exportieren Sie die Datei aus Untis, indem Sie den Pfad gehen:",
            $"1. [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Datei > Import/Export > Export TXT Datei[/]",
            $"2. Als Delimiter muss '|' ausgewählt werden; UTF8; Textbegrenzung: '\"'",
            $"3. Die Datei auswählen.",
            $"4. Die Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."
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
            "OpenPeriod",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie den Pfad gehen:",
                $"[bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Klassenbuch > Offene Stunden > Bericht[/]",
                $"Die PDF-Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."
            ],
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
                $"1. [bold {Global.GetColor(Global.ColorKopfzeileInCSV)}]Stammdaten > Schüler*innen > Berichte > Schüler > CSV-Ausgabe[/]",
                $"2. Die Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."
            ],
            [""],
            true,
            d => d.FilternWebuntisStudent(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "ApprenticeRepresentative_",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie als Administrator den Pfad gehen:",
                $"1. [bold {Global.GetColor(Global.ColorKopfzeileInCSV)}]Stammdaten > Ausbildungsbeauftragte > Berichte > Ausbildungsbeauftragte > CSV-Ausgabe[/]",
                $"2. Die Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."
            ],
            [""],
            true,
            d => d.FilternWebuntisStudent(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "LegalGuardian_",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie als Administrator den Pfad gehen:",
                $"1. [bold {Global.GetColor(Global.ColorKopfzeileInCSV)}]Stammdaten > Erziehungsberechtigte > Berichte > Erziehungsberechtigte > CSV-Ausgabe[/]",
                $"2. Die Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."
            ],
            [""],
            true,
            d => d.FilternWebuntisStudent(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "GPU002",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Untis, indem Sie als Admin den Pfad gehen:",
                $"[bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Datei > Import/Export > Export TXT Datei > Unterricht[/]",
                $"Trennzeichen: [{Global.GetColor(Global.ColorPfadInDateien)}]|[/]",
                $"Textbegrenzung: [{Global.GetColor(Global.ColorPfadInDateien)}]\"[/]",
                $"Encoding: [{Global.GetColor(Global.ColorPfadInDateien)}]UTF-8[/]",
                $"Die Datei speichern: [{Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}/GPU002.TXT[/]"
            ],
            [""],
            false,
            d => d.FilternKlassenGPU002(),
            "*.TXT",
            "|"
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
            "GPU005",
            "Beschreibung: Lehrkraefte",
            untishinweise,
            [""],
            false,
            d => d.FilternLehrkraefteGPU004(),
            "*.TXT",
            "|"
        ));
        Add(new Datei(
            "GPU006",
            "Beschreibung: Fächer",
            untishinweise,
            [""],
            false,
            d => d.FilternFaecherGPU006(),
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
            untishinweise,
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
            d => d.FilterKurse(),
            "*.dat",
            "|",
            true
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
            "AbsencePerStudent",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie als Admin den Pfad gehen:",
                $"1. [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Klassenbuch > Berichte[/]",
                $"2. Evtl. Klasse auswählen. Evtl. den Zeitraum eingrenzen. Beides ist nicht zwingend notwendig.",
                $"3. [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Fehlzeiten[/], [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Verspätungen[/] und [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Nur zählende Abwesenheiten[/] anhaken.",
                $"4. Auf das CSV-Symbol hinter [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Fehlzeiten pro Schüler*in[/] [bold {Global.GetColor(Global.ColorFehler)}]pro Tag[/] klicken.",
                $"5. Die Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."

            ],
            [""],
            true,
            d => d.FilternAbsencePerLessons(),
            "*.csv",
            "\t"
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
            $"Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(configuration["PfadDownloads"],"termine_kollegium.csv") + "[/] überschreiben oder neu anlegen."
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
            $"Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(configuration["PfadDownloads"],"termine_fhr.csv") + "[/] überschreiben oder neu anlegen."
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
            $"Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(configuration["PfadDownloads"],"termine_verwaltung.csv") + "[/] überschreiben oder neu anlegen."
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
            $"Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(configuration["PfadDownloads"],"termine_berufliches_gymnasium.csv") + "[/] überschreiben oder neu anlegen."
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
            "Adressen",
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
            "SchuelerTelefonnummern",
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
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie als Admin den Pfad gehen:",
                $"1. [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Klassenbuch > Berichte[/]",
                $"2. Evtl. Klasse auswählen. Evtl. den Zeitraum eingrenzen.",
                $"3. [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Noten > Noten pro Schüler*in > CSV-Symbol[/] klicken.",
                $"4. Die Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."

            ],
            [""],
            true,
            d => d.FilternMarksPerLessons(),
            "*.csv",
            "\t"
        ));
        Add(new Datei(
            "StudentgroupStudents",
            "Beschreibung",
            [
                "Exportieren Sie die Datei aus Webuntis, indem Sie als Administrator den Pfad gehen:",
                $"1. [bold {Global.GetColor(Global.ColorKopfzeileInCSV)}]Administration > Export > StudentgroupStudents.csv[/]",
                $"2. Die Datei in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] speichern."
            ],
            [],
            true,
            d => d.FilternStudentgroupStudents(),
            "*.csv",
            "\t"
        ));                    
    }

    public List<dynamic>? GetMatchingList(IConfiguration configuration, string pattern, Students students = null, Klassen klassen = null, string[] spalten = null, DokuwikiZugriff dokuwikiZugriff = null)
    {
        Datei datei = this.FirstOrDefault(datei => !string.IsNullOrEmpty(datei.Dateiname) && datei.Dateiname.ToLower().StartsWith(pattern, StringComparison.CurrentCultureIgnoreCase));
        
        // Spezielle Behandlung für "struct" Dateiendung: Struct-Dateien müssen erst gefüllt werden.
        // Die bereits angelegte Datei wird gefüllt
        if(spalten != null && spalten.Length > 0)
        {                       
            datei.GetSchema(pattern, spalten, configuration, dokuwikiZugriff );
            datei.SchreibeZeilen("|");
            return null;
        }        
                                       
        // Mögliche Meldungen werden ausgegeben, wenn die Datei nicht gefunden wurde oder veraltet ist.

        if (string.IsNullOrEmpty(datei.AbsoluterPfad) && datei.Endung.ToLower().Contains("dat"))
            datei.Fehlermeldung += $"Die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]SchuelerBasisdaten.dat[/] wurde weder in [bold aqua]{configuration["pfadDownloads"]}[/] noch in [bold aqua]{configuration["pfadSchilddatenaustausch"]}[/] gefunden. Am besten gehen Sie jetzt in SchILD zu [bold springGreen2]Datenaustausch > Schnittstelle SchILD NRW > Export[/] und klicken [bold springGreen2]Export starten[/], um alle Dateien nach [bold aqua]{configuration["pfadSchilddatenaustausch"]}[/] zu exportieren. Anschließend kehren Sie hierher zurück.";
        else if (datei.AbsoluterPfad == null)
            datei.Fehlermeldung += $"Die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{pattern}[/] wurde nicht gefunden. Bitte prüfen Sie, ob sie im Ordner [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] vorhanden ist.";
        //else if (datei.IstOptional && datei.Count == 0)
        //    datei.Fehlermeldung = $"Die Datei [bold aqua]{pattern}[/] ist optional und wurde nicht gefunden. Sie wird nicht benötigt.";
        else if (datei.IstOptional && datei.Count > 0)
            return datei.ToList();
        else if (datei.AbsoluterPfad == null || datei.AbsoluterPfad.Length == 0 && !datei.IstOptional)
            datei.Fehlermeldung = $"Die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{pattern}[/] existiert nicht. Bitte prüfen Sie, ob sie im Ordner [{Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] vorhanden ist.";

        datei.FehlermeldungRendern(configuration);

        // Rückgabe der Datei, wenn sie gefunden wurde und die Bedingungen erfüllt sind.

        if (!datei.IstOptional && (students == null || students.Count == 0) && datei.IstVeraltet(configuration))
            return [];
        else if (datei.AbsoluterPfad.ToLower().Contains("schuelerbasisdaten") && (students == null || students.Count == 0))
            return datei.ToList();        
        else if (datei.IstOptional)
            return datei.ToList();
        else if (datei.AbsoluterPfad.ToLower().Contains("studentgroupstudents") && students != null && students.Count > 0)
            return datei.FilternStudentgroupStudents(students, students.GetKlassen());
        else if (students != null && students.Count > 0)
            return datei.Filtern(students, students.GetKlassen());
        else if (datei.AbsoluterPfad.ToLower().Contains("lehrkraefte"))
            return datei.ToList();

        return datei.ToList();
    }

    /// <summary>
    /// Aus allen möglchen Quelldateien (quelldateien) werden die notwendigen Quelldateien (m.Quelldateien) herausgefiltert.
    /// </summary>
    /// <param name="configuration"></param>
    /// <param name="dateinamenNotwendigeDateien"></param>
    /// <param name="meldungAnzeigen"></param>
    /// <returns></returns>
    public Dateien Notwendige(IConfiguration configuration, List<string> dateinamenNotwendigeDateien, bool meldungAnzeigen = false)
    {
        var pfadDownloads = configuration["PfadDownloads"];
        int maxDateiAlter = int.TryParse(configuration["MaxDateiAlter"], out var parsedMaxDateiAlter) ? parsedMaxDateiAlter : 0;

        Dateien notwendige = new Dateien();

        foreach (var dateinameNotwendig in dateinamenNotwendigeDateien)
        {
            var dateiendung = dateinameNotwendig.Split(',')[1].Trim().ToLower();
            var dateiname = dateinameNotwendig.Split(',')[0];

            var datei = new Datei();
            
            if(dateiendung == "struct")
            {
                // Wenn die Datei noch nicht existiert, wird sie erstellt, aber nicht gefüllt. Sie wird später gefüllt, wenn die Spalten bekannt sind.
                datei.Erstellen(Path.Combine(pfadDownloads, dateiname + "." + dateiendung));
                datei.AbsoluterPfad = Path.Combine(pfadDownloads, dateiname + "." + dateiendung);
                datei.Dateiname = dateiname;                
            }
            else
            {
                datei = this.First(datei => !string.IsNullOrEmpty(datei.Dateiname)
                    && !datei.Dateiname.ToLower().Contains("-kennwort")
                    && datei.Dateiname.ToLower().StartsWith(dateiname.ToLower(), StringComparison.CurrentCultureIgnoreCase)
                    && datei.Endung.ToLower().Contains("*." + dateiendung.ToLower(), StringComparison.CurrentCultureIgnoreCase));                

                datei.IstOptional = dateinameNotwendig.Split(',').Length > 2 && dateinameNotwendig.Split(',')[2].ToLower().Contains("opt") ? true : false;
                datei.Nur177659 = dateinameNotwendig.Split(',').Length > 2 && dateinameNotwendig.ToLower().Contains("177659") ? true : false;

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
                                datei.Fehlermeldung = $"Die Datei [bold aqua]{absoluterPfad}[/] existiert, ist aber veraltet.";
                                if (meldungAnzeigen && !datei.IstOptional)
                                    datei.FehlermeldungRendern(configuration);
                            }
                        }
                        else
                        {
                            if (!datei.DarfLeerSein)
                            {
                                //datei.Fehlermeldung = $"Die Datei [bold aqua]{absoluterPfad}[/] existiert, ist aber leer. Ist die Datei evtl. vorher in Excel o.ä. geöffnet worden? Oder stimmt der Delimiter nicht? Der korrekte Delimiter ist: '[bold aqua]{datei.Delimiter}[/]'";
                                if (meldungAnzeigen && !datei.IstOptional)
                                    datei.FehlermeldungRendern(configuration);
                            }
                        }
                    }
                    else
                    {
                        if (((IDictionary<string, object>)datei[0]).Count == 1)
                        {
                            datei.Fehlermeldung = $"Die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{absoluterPfad}[/] hat nur eine einzige Spalte. Das korrekte Trennzeichen ist: [{Global.GetColor(Global.ColorZahlen)}]'{datei.Delimiter}'[/].";
                            if (meldungAnzeigen && !datei.IstOptional)
                                datei.FehlermeldungRendern(configuration);
                            
                        }

                        if (datei.Erstelldatum.Date.AddDays(maxDateiAlter) < DateTime.Now.Date)
                        {
                            var veraltet = $"darf aber nicht älter als [bold {Global.GetColor(Global.ColorHinweise)}]{configuration["MaxDateiAlter"]}[/] Tage sein.";
                            if (configuration["MaxDateiAlter"] == "1")
                                veraltet = "darf aber nicht älter als von gestern sein.";
                            if (configuration["MaxDateiAlter"] == "0")
                                veraltet = "muss aber von heute sein.";
                            datei.Fehlermeldung = $"Die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{datei.AbsoluterPfad}[/] ist veraltet. Sie wurde am [bold {Global.GetColor(Global.ColorHinweise)}]{datei.Erstelldatum:dd.MM.yyyy}[/] erstellt, {veraltet}";
                            if (meldungAnzeigen && !datei.IstOptional)
                                datei.FehlermeldungRendern(configuration);
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

                    datei.Fehlermeldung = $"Die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads, dateiname)}[/] existiert nicht{opt}.";
                    if (meldungAnzeigen && !datei.IstOptional)
                        datei.FehlermeldungRendern(configuration);
                }
            }
            notwendige.Add(datei);            
        }

        return notwendige;
    }

    public void GetZeilen(IConfiguration configuration, List<string>? dateienImPfad = null)
    {
        try
        {
            configuration = Global.Konfig("MaxDateiAlter", Global.Modus.ReadSilent, configuration);
            var maxDateiAlterString = configuration["MaxDateiAlter"];
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

            AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("Menü laden ...", ctx =>
            {
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
                    Meldung.Add($"Keine Dateien mit auswertbaren Zeilen in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["pfadDownloads"]}[/] gefunden.");
                }
                else if (anzahlDateienMitZeilen == 1)
                {
                    Meldung.Add($"Nur eine Datei mit auswertbaren Zeilen in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["pfadDownloads"]}[/] gefunden.");
                }
                else
                {
                    //Meldung.Add($"[bold {Global.GetColor(Global.ColorZahlen)}]{anzahlDateienMitZeilen}[/] Dateien aus [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["pfadDownloads"]}[/] eingelesen.");
                }
                if (anzahlDateienMitZeilen == 0 && dateienImPfad.Count > 0)
                {
                    Meldung.Add($"[bold {Global.GetColor(Global.ColorZahlen)}]{dateienImPfad.Count}[/] Dateien im Download-Pfad gefunden, aber keine mit auswertbaren Zeilen.");
                }
            });
            Meldung.Add("");
            Global.DisplayHeader(configuration, Meldung);
        }
        catch (Exception e)
        {
            throw e;
        }
    }

    /// <summary>
    /// Holt alle Dateien im Download-Pfad, die den Kriterien entsprechen.
    /// Die Kriterien sind: Dateiendung .csv, .txt, .dat oder Dateiname beginnt mit "openperiod".
    /// </summary>
    /// <param name="configuration"></param>
    /// <returns></returns>
    public List<string> GetDateienImPfad(IConfiguration configuration)
    {
        string pfad = configuration["PfadDownloads"];
        
        // Sicherstellen, dass der Pfad existiert, um Exceptions zu vermeiden
        if (string.IsNullOrEmpty(pfad) || !Directory.Exists(pfad))
            return new List<string>();

        // Definition der erlaubten Endungen
        var erlaubteEndungen = new[] { ".csv", ".txt", ".dat" };

        return Directory.GetFiles(pfad, "*", SearchOption.AllDirectories) // "AllDirectories" sucht in Unterordnern
            .Where(f => 
            {
                string ext = Path.GetExtension(f);
                string name = Path.GetFileNameWithoutExtension(f);

                // Prüfung auf Endung ODER ob der Name mit "openperiod" beginnt
                return erlaubteEndungen.Contains(ext, StringComparer.OrdinalIgnoreCase) || 
                    name.StartsWith("openperiod", StringComparison.OrdinalIgnoreCase);
            })
            .OrderBy(f => File.GetLastWriteTime(f)) // Sortierung nach letztem Schreibzugriff
            .ToList();
    }

    /// <summary>
    /// Die Exportdateien aus SchILD werden zu den anderen heruntergeladenen Dateien verschoben, um Platz zu machen für neue
    /// Dateien, die dann nach SchILD importiert werden. 
    /// Die Funktion muss mit jedem Durchlauf und unmittelbar vor dem Erstellen neuer Dateien im Ausgabeverzeichnis
    /// aufgerufen werden, damit Dateien aus SchILD nicht mit den neuen Dateien gemixt werden.
    /// </summary>
    public void ExportAusSchildVerschieben(IConfiguration configuration)
    {
        configuration = Global.Konfig("PfadSchilddatenaustausch", Global.Modus.ReadSilent, configuration);
        var pfadDownloads = configuration["PfadDownloads"];

        // Stelle sicher, dass der Zielordner existiert
        if (!Directory.Exists(configuration["PfadSchilddatenaustausch"]))
        {
            try
            {
                Directory.CreateDirectory(configuration["PfadSchilddatenaustausch"]);
            }
            catch (Exception ex)
            {
                AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
                return;
            }
        }

        // Die SchildSchuelerExport wird immer kopiert.
        var datei = !string.IsNullOrEmpty(configuration["PfadSchilddatenaustausch"])
            ? Directory.GetFiles(configuration["PfadSchilddatenaustausch"], "*", SearchOption.TopDirectoryOnly).FirstOrDefault(f => Path.GetFileName(f).ToLower().Contains("schildschuelerexport"))
            : null;

        if (datei != null)
        {
            if (pfadDownloads == null)
                throw new ArgumentNullException(nameof(pfadDownloads), "Der Pfad zum Downloads-Ordner ist null.");

            var destinationPath = Path.Combine(pfadDownloads, Path.GetFileName(datei));
            File.Copy(datei, destinationPath, true);
        }

        var datFiles = !string.IsNullOrEmpty(configuration["PfadSchilddatenaustausch"])
            ? Directory.GetFiles(configuration["PfadSchilddatenaustausch"], "*.dat").ToList()
            : new List<string>();

        // Wenn mehr als 5 Dateien im Ausgangsordner sind, dann muss es sich um Exportdateien handeln,
        // da BKB-Tool niemals mehr als 5 Dateien gleichzeitig in das Ausgabeverzeichnis verschiebt.
        if (datFiles.Count > 7)
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

                var hinweis = $"SchILD verwendet für den Export und den (Re-)Import denselben Ordner. Deswegen verschiebt [bold {Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] jetzt " +
                    $"[bold {Global.GetColor(Global.ColorZahlen)}]{datFiles.Count} aus SchILD exportierte *.dat-Dateien[/] direkt von [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadSchilddatenaustausch"]}[/] nach [bold {Global.GetColor(Global.ColorPfadInDateien)}]{pfadDownloads}[/]. " +
                    $"\nDie aufbereiteten Dateien stellt [bold {Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] wiederum in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadSchilddatenaustausch"]}[/] bereit. " +
                    "So bleiben die Import-Dateien und Export-Dateien stets getrennt voneinander." +
                    $"\nWeiter mit [bold {Global.GetColor(Global.ColorActionInMenüs)}]ENTER[/].";

                Global.Konfig("FirstRun", Global.Modus.Read, configuration, null, -1, -1, hinweis);

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

                //AnsiConsole.Write(new Rule($"[bold fuchsia] {datFiles.Count} Dateien verschoben von [bold aqua]{configuration["PfadSchilddatenaustausch"]}[/] nach [bold aqua]{configuration["PfadDownloads"]}[/][/] ").RuleStyle("fuchsia").LeftJustified());
                this.Meldung.Add($"[bold {Global.GetColor(Global.ColorZahlen)}]{datFiles.Count}[/] Dateien verschoben von [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadSchilddatenaustausch"]}[/] nach [bold {Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/]");
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
            datei.FehlermeldungRendern(configuration);
        }
    }

    private bool IsDirectoryWritable(string dirPath)
    {
        try
        {
            string testFile = Path.Combine(dirPath, Path.GetRandomFileName());
            using (FileStream fs = File.Create(testFile, 1, FileOptions.DeleteOnClose)) { }
            return true;
        }
        catch
        {
            return false;
        }
    }

    internal void VergleichenFilternErstellen(Dateien quelldateien)
    {
        var bereitsGeöffnet = false;

        // Erstelle alle Dateien, die in der Liste enthalten sind.
        foreach (var datei in this)
        {
            try
            {
                datei.VergleichenUndFiltern(quelldateien);

                // OrdnerOeffnen wird nur einmal aufgerufen, sobald die erste Datei, die count > 0 hat, erstellt wird.
                if (datei.Count > 0 && !bereitsGeöffnet)
                {
                    datei.OrdnerOeffnen();
                    bereitsGeöffnet = true;
                }
                datei.Erstellen();
            }
            catch (Exception ex)
            {
                AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
                AnsiConsole.MarkupLine($"[bold {Global.GetColor(Global.ColorFehler)}]Fehler beim Erstellen der Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{datei.Dateiname}[/]: {ex.Message}[/]");
            }
        }
    }

    public void OrdnerOeffnen()
    {
        var verschiedeneOrdner = this.Select(datei => Path.GetDirectoryName(datei.AbsoluterPfad))
            .Distinct()
            .Where(dir => !string.IsNullOrEmpty(dir) && Directory.Exists(dir))
            .ToList();

        // Öffne die Ordner, in denen die Dateien gespeichert werden sollen.
        foreach (var ordner in verschiedeneOrdner)
        {
            if (ordner != null && Directory.Exists(ordner))
            {
                try
                {
                    if (IsDirectoryWritable(ordner))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = ordner,
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        AnsiConsole.MarkupLine($"[bold {Global.GetColor(Global.ColorFehler)}]Der Ordner [bold {Global.GetColor(Global.ColorPfadInDateien)}]{ordner}[/] ist nicht beschreibbar.[/]");
                    }
                }
                catch (Exception ex)
                {
                    AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
                }
            }
        }
    }

    internal void Erstellen()
    {
        foreach (var datei in this)
        {
            datei.Erstellen();
        }
    }

    internal void OrdnerÖffnen()
    {
        var bereitsGeöffnet = false;
        
        foreach (var datei in this)
        {
            try
            {   
                // OrdnerOeffnen wird nur einmal aufgerufen, sobald die erste Datei, die count > 0 hat, erstellt wird.
                if (datei.Count > 0 && !bereitsGeöffnet)
                {
                    datei.OrdnerOeffnen();
                    bereitsGeöffnet = true;
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
                AnsiConsole.MarkupLine($"[bold {Global.GetColor(Global.ColorFehler)}]Fehler beim Erstellen der Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{datei.Dateiname}[/]: {ex.Message}[/]");
            }
        }
    }

    internal void Verschieben(IConfiguration configuration)
    {
        foreach (var datei in this)
        {
            if(datei.AbsoluterPfad.ToLower().Contains("littera"))
                datei.Verschieben(configuration["PfadLitteraImport"]);
        }
    }

    internal void ZippenMitKennwort(IConfiguration configuration)
    {
        configuration = Global.Konfig("ZipKennwort", Global.Modus.Update, configuration);

        foreach (var datei in this)
        {
            if(datei.AbsoluterPfad.ToLower().Contains("netman"))
                datei.ZippenMitKennwort(configuration);
        }
    }
}