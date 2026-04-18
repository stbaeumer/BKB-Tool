using System.Configuration;
using System.Dynamic;
using System.Globalization;
using Microsoft.Extensions.Configuration;
using System.Text.RegularExpressions;
using CookComputing.XmlRpc;
using Spectre.Console;
using Color = Spectre.Console.Color;
using System.Text;
using System.Diagnostics;
using Spectre.Console.Rendering;

namespace Common;

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
#pragma warning disable CS0162 // Unbenutzter Code

public class Menüeintrag
{
    public Gruppen Gruppen { get; set; }
    public List<string> Beschreibung { get; set; }
    public Dateien Quelldateien { get; set; }
    public bool DateienFehlenOderSindLeer { get; set; }
    public string Titel { get; set; }
    public Action<Menüeintrag> Funktion { get; } // Funktion mit Menüeintrag als Parameter
    public Students Students { get; set; }
    public Klassen Klassen { get; set; }

    /// <summary>
    /// Interessierende Klassen
    /// </summary>
    public List<string> IKlassen { get; set; }

    /// <summary>
    /// Interessierende Students
    /// </summary>
    public Students IStudents { get; set; }
    public Global.Rubrik Rubrik { get; private set; }
    public Global.NurBeiDiesenSchulnummern NurBeiDiesenSchulnummern { get; }
    public Datei Zieldatei { get; set; }
    public List<string> AlleSchildKursBez { get; private set; }
    public Unterrichte Unterrichte { get; set; }
    public Dateien Zieldateien { get; set; }
    public Relationsgruppen Relationsgruppen { get; internal set; }

    public Menüeintrag(string titel, Dateien quelldateien, Students students, Klassen klassen, List<string> beschreibung, Action<Menüeintrag> funktion, Global.Rubrik rubrik = Global.Rubrik.Allgemein, Global.NurBeiDiesenSchulnummern nurbeiDiesenSchulnummern = Global.NurBeiDiesenSchulnummern.Alle)
    {
        try
        {
            if (titel == null) throw new ArgumentNullException(nameof(titel));
            if (quelldateien == null) throw new ArgumentNullException(nameof(quelldateien));
            if (students == null) throw new ArgumentNullException(nameof(students));
            if (klassen == null) throw new ArgumentNullException(nameof(klassen));
            if (beschreibung == null) throw new ArgumentNullException(nameof(beschreibung));
            {
                Titel = titel;
                Quelldateien = quelldateien;
                Students = students;
                Klassen = klassen;
                DateienFehlenOderSindLeer = false;
                Beschreibung = beschreibung;
                Funktion = funktion;
                Gruppen = new Gruppen();
                IKlassen = new List<string>();
                IStudents = new Students();
                Rubrik = rubrik;
                NurBeiDiesenSchulnummern = nurbeiDiesenSchulnummern;
                Unterrichte = new Unterrichte();
                Zieldateien = new Dateien();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Fehler beim Erstellen des Menüeintrags '{titel}': {ex.Message}", ex);
        }
    }

    public Datei Ausführen()
    {
        Funktion?.Invoke(this);
        return Zieldatei;
    }

    public Datei LernabschnittsdatenAlt(string zielDateiname, IConfiguration configuration)
    {
        var zieldatei = new Datei(zielDateiname);

        FilterInteressierendeStudentsUndKlassen(configuration);

        var basisDa = Quelldateien.GetMatchingList(configuration, "basisdaten", IStudents, Klassen);
        if (basisDa == null || !basisDa.Any()) return new Datei(zielDateiname);

        var atlzeug = Quelldateien.GetMatchingList(configuration, "atlantis-zeugnisse", IStudents, Klassen);
        if (atlzeug == null || !atlzeug.Any()) return new Datei(zielDateiname);

        var klassen = Quelldateien.GetMatchingList(configuration, "klassen", IStudents, Klassen);
        if (klassen == null || !klassen.Any()) return new Datei(zielDateiname);

        var quit = Global.Konfig("MaxDateiAlter", Global.Modus.Update, configuration);

        throw new Exception("Fehler bei der Konfiguration des maximalen Alters der eingelesenen Dateien. Bitte überprüfen Sie die Eingabe.");
        /*
        if (quit) return [];

        foreach (var student in IStudents)
        {
            foreach (var zeugnisdatum in student.GetZeugnisDatums(atlzeug))
            {
                var zeile = student.Nachname + ", " + student.Vorname + ", " + student.Klasse + ", Zeugnisdatum: " +
                            zeugnisdatum.ToShortDateString();

                var recBasis = basisDa
                    .Where(record =>
                    {
                        var dict = (IDictionary<string, object>)record;
                        return dict["Vorname"].ToString() == student.Vorname &&
                               dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                    })
                    .FirstOrDefault();

                if(recBasis == null) continue;

                var recKla = klassen
                    .Where(record =>
                    {
                        var dict = (IDictionary<string, object>)record;
                        return dict["InternBez"].ToString() == student.Klasse;
                    })
                    .FirstOrDefault();

                if (recKla == null) continue;
                
                var dictKla = (IDictionary<string, object>)recKla;                

                var recAtl = atlzeug
                .Where(record =>
                {
                    var dict = record as IDictionary<string, object>;
                    return dict != null &&
                        dict["Field1"]?.ToString()?.Replace("'", "") == student.Nachname &&
                        dict["Field3"]?.ToString()?.Replace("'", "") == student.Geburtsdatum &&
                        dict["Field4"]?.ToString()?.Replace("'", "") == zeugnisdatum.ToShortDateString();
                })
                .FirstOrDefault();

                if (recAtl == null) continue;
                    
                var dictAtl = (IDictionary<string, object>)recAtl;                
                                
                var dictBasis = (IDictionary<string, object>)recBasis;

                var jahr = student.GetJahr(zeugnisdatum);
                zeile += ", Jahr: " + jahr;
                var abschnitt = student.GetAbschnitt(zeugnisdatum);
                zeile += ", Abschnitt: " + abschnitt;
                var jahrgang = student.GetJahrgang(klassen, jahr, zeugnisdatum, recBasis);
                zeile += ", Jg: " + jahrgang;

                dynamic record = new ExpandoObject();
                record.Nachname = student.Nachname;
                record.Vorname = student.Vorname;
                record.Geburtsdatum = student.Geburtsdatum;
                record.Jahr = jahr;
                record.Abschnitt = abschnitt;
                record.Jahrgang = jahrgang;
                record.Klasse = student.Klasse;
                record.Schulgliederung = dictKla["Gliederung"];
                record.OrgForm = dictKla["OrgForm"];
                record.Klassenart = dictKla["Klassenart"];
                record.Fachklasse = "";
                record.Förderschwerpunkt = "";
                record.ZWEIPUNKTLEERZEICHENFörderschwerpunkt = "";
                record.Schwerstbehinderung = student.Schwerstbehinderung;
                record.Wertung = "J";
                record.Wiederholung = "N";
                record.Klassenlehrer = "";
                record.Versetzung = "";
                record.Abschluss = "";
                record.Schwerpunkt = "";
                if (dictAtl != null)
                {
                    record.Konferenzdatum = dictAtl["Field5"].ToString().Replace("'", "");
                }
                else
                {
                    record.Konferenzdatum = string.Empty;
                }
                record.Zeugnisdatum = zeugnisdatum.ToShortDateString();
                record.SummeFehlstd = "";
                record.SummeFehlstdLEERZEICHENunentschuldigt = "";
                record.allgPUNKTLEERZEICHENbildenderLEERZEICHENAbschluss = "";
                record.berufsbezPUNKTLEERZEICHENAbschluss = "";
                record.Zeugnisart = "";
                record.FehlstundenMINUSGrenzwert = "";
                // Ein Jahrgang kleiner als 01 deutet auf eine Laufbahn vor der aktuellen Klasse hin und wird ignoriert.
                if (jahrgang.EndsWith("00")) continue;
                // -- deutet auf einen noch älteren Abschnitt hin
                if (jahrgang.Contains("--") || jahrgang.Contains("00")) continue;
                zieldatei.Add(record);
                Global.ZeileSchreiben(zeile, "ok", ConsoleColor.White, ConsoleColor.Gray);
            }
        }*/

        return zieldatei;
    }

    /// <summary>
    /// Die IStundets und die IKlassen (List<string>) werden als Eigenschaft des Menüeintrags initialisiert.
    /// </summary>
    public void FilterInteressierendeStudentsUndKlassen(IConfiguration configuration, string aufforderung = "", string hinweise = "")
    {
        var interessierendeStudents = new Students();

        Global.Konfig("Klassen", Global.Modus.Update, configuration, aufforderung, -1, -1, hinweise, "", this.Students);

        var interessierendeKlassen = configuration["Klassen"].ToString().Split(",").ToList();

        IStudents = new Students();
        IStudents.AddRange(from s in Students
                           where interessierendeKlassen.Contains(s.Klasse)
                        where s.Status == "2" || s.Status == "6"
                           select s);
        IKlassen = interessierendeKlassen;

        var linkeSeite = (string.Join(",", interessierendeKlassen.Order()));
        linkeSeite = linkeSeite.Substring(0, Math.Min(Console.WindowWidth / 2, linkeSeite.Length));

        if (interessierendeKlassen.Count > 3)
        {
            linkeSeite = linkeSeite + " (" + interessierendeKlassen.Count().ToString() + " Klassen)";
        }

        if (linkeSeite != "")
        {
            Global.ZeileSchreiben(linkeSeite, IStudents.Count().ToString(), ConsoleColor.Gray, ConsoleColor.White);
        }
    }

    public Datei? LeistungsdatenAlt(IConfiguration configuration, string unterordnerUndDateiname)
    {
        var fehlendeFaecher = new List<string>();

        var basisdaten = Quelldateien.GetMatchingList(configuration, "basisdaten", IStudents, Klassen);
        //var gostdaten = GetMatchingList("gost");
        var atlantiszeugnisse = Quelldateien.GetMatchingList(configuration, "atlantis-zeugnisse", IStudents, Klassen);
        var klassen = Quelldateien.GetMatchingList(configuration, "klassen", IStudents, Klassen);
        var faecher = Quelldateien.GetMatchingList(configuration, "faecher", IStudents, Klassen);

        if (atlantiszeugnisse == null || atlantiszeugnisse.Count == 0 || klassen == null || klassen.Count == 0 || faecher == null || faecher.Count == 0) return [];

        var zieldatei = new Datei(unterordnerUndDateiname);

        foreach (var student in IStudents.Where(x => !x.Klasse.StartsWith("G")))
        {
            foreach (var zeugnisdatum in student.GetZeugnisDatums(atlantiszeugnisse))
            {
                var writeline = student.Nachname + ", " + student.Vorname + ", " + student.Klasse + ", ";

                List<dynamic> recAtl = atlantiszeugnisse
                .Where(record =>
                {
                    var dict = record as IDictionary<string, object>;
                    return dict != null &&
                        dict["Field1"]?.ToString()?.Replace("'", "") == student.Nachname &&
                        dict["Field3"]?.ToString()?.Replace("'", "") == student.Geburtsdatum &&
                        dict["Field4"]?.ToString()?.Replace("'", "") == zeugnisdatum.ToShortDateString() &&
                        !string.IsNullOrEmpty(dict["Field8"].ToString()); // Nur benotete Fächer
                })
                .GroupBy(record => // keine doppelten Einträge 
                {
                    var dict = record as IDictionary<string, object>;
                    return dict != null ? new
                    {
                        Nachname = dict["Field1"]?.ToString()?.Replace("'", ""),
                        Vorname = dict["Field2"]?.ToString()?.Replace("'", ""),
                        Geburtsdatum = dict["Field3"]?.ToString()?.Replace("'", ""),
                        Zeugniskonferenzdatum = dict["Field4"]?.ToString()?.Replace("'", ""),
                        Kurztext = dict["Field9"]?.ToString()?.Replace("'", "")
                    } : null;
                })
                .Select(group => group.First())
                .ToList();

                var recBasis = basisdaten
                    .Where(record =>
                    {
                        var dict = (IDictionary<string, object>)record;
                        return dict["Vorname"].ToString() == student.Vorname &&
                               dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                    })
                    .FirstOrDefault();
                var dictBasis = (IDictionary<string, object>)recBasis;

                var jahr = student.GetJahr(zeugnisdatum);
                writeline += ", " + jahr;
                var abschnitt = student.GetAbschnitt(zeugnisdatum);
                writeline += ", " + abschnitt;
                var jahrgang = student.GetJahrgang(klassen, jahr, zeugnisdatum, dictBasis);
                writeline += ", " + jahrgang;

                foreach (var zeile in recAtl)
                {
                    var dict = (IDictionary<string, object>)zeile;
                    var fach = dict["Field9"].ToString();

                    fach = fach.Replace("  ", " ").Replace("B1", "").Replace("C1", "").Replace("A1", "")
                        .Replace("  ", " ")
                        .Replace("B2", "").Replace("C2", "").Replace("A2", "")
                        .Replace(" GD", " G1").Replace(" GE", " G1").Replace(" GB", " G1").Replace("  ", " ");

                    var kursart = GetKursart(configuration, jahrgang, fach);
                    var note = dict["Field8"].ToString();
                    var tendenz = dict["Field10"].ToString();
                    var punkte = dict["Field11"].ToString();
                    var noteOderPunkte = dict["Field13"].ToString();

                    dynamic record = new ExpandoObject();
                    record.Nachname = $"{student.Nachname}#{student.Klasse}";
                    record.Vorname = student.Vorname;
                    record.Geburtsdatum = student.Geburtsdatum;
                    record.Jahr = jahr;
                    record.Abschnitt = abschnitt;
                    record.Fach = fach.Replace("'", "").Replace("**", "");
                    record.Fachlehrer = "";
                    record.Kursart = kursart;
                    record.Kurs = "";
                    record.Note = student.GetNote(note, noteOderPunkte, punkte, fach, tendenz);
                    record.Abiturfach = "";
                    record.WochenstdPUNKT = "";
                    record.ExterneSchulnrPUNKT = "";
                    record.Zusatzkraft = "";
                    record.WochenstdPUNKTLEERZEICHENZK = "";
                    record.Jahrgang = jahrgang;
                    record.Jahrgänge = "";
                    record.FehlstdPUNKT = "";
                    record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";

                    // Doppelminus deutet auf einen noch ältern Abschnitt hin
                    if (!jahrgang.Contains("--") && !jahrgang.Contains("00"))
                    {
                        zieldatei.Add(record);
                        var writeline2 = writeline + ", Fach: " + dict["Field6"].ToString() + ", Note:" +
                                         dict["Field8"].ToString();
                        Global.ZeileSchreiben(writeline2, "ok", ConsoleColor.Green, ConsoleColor.White);
                    }
                }
            }
        }

        if (fehlendeFaecher.Count() <= 0) return zieldatei;
        Global.ZeileSchreiben("Es fehlen Fächer müssen in SchILD angelegt werden:",
            fehlendeFaecher.Count().ToString(), ConsoleColor.DarkYellow, ConsoleColor.White);
        Global.ZeileSchreiben("Fächer: ", string.Join(", ", fehlendeFaecher), ConsoleColor.Green, ConsoleColor.White);

        return zieldatei;
    }

    private static string GetKursart(IConfiguration configuration, string jahrgang, string fach)
    {
        //ToDo: Hier müssen weitere Konfigs kommen.
        
        if (!jahrgang.StartsWith("GY")) return "PUK";
        if (!jahrgang.EndsWith("02") && !jahrgang.EndsWith("03") && !jahrgang.EndsWith("12") && !jahrgang.EndsWith("13")) return "PUK";
        if (!fach.Contains(" L")) return fach.Contains(" G") ? "GKS" : "PUK";
        var linkerTeil = fach.Split(' ')[0].TrimEnd();
        
        configuration = Global.Konfig("Lk1faecher", Global.Modus.Read, configuration);
        var lk1faecher = configuration["Lk1faecher"].ToString().Split(',').Select(x => x.Trim()).ToList();
        return lk1faecher.Contains(linkerTeil) ? "LK1" : "LK2";
    }

    public void LernabschnittsdatenAnlegen(
        Lehrers lehrers,
        IConfiguration configuration,
        Global.Zweck art,
        string zieldateiname,        
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        var schuelerLernab = Quelldateien.GetMatchingList(configuration, "schuelerlernabschnittsdaten", IStudents, Klassen);
        if (schuelerLernab == null || !schuelerLernab.Any()) return;

        var schuelerBasisd = Quelldateien.GetMatchingList(configuration, "schuelerbasisdate", IStudents, Klassen);
        if (schuelerBasisd == null || !schuelerBasisd.Any()) return;

        var absencePerStud = new List<dynamic>();

        var konferenzdatum = DateTime.Now;
        var zeugnisdatum = DateTime.Now;

        configuration = Global.Konfig("Abschnitt", Global.Modus.Read, configuration);
        configuration = Global.Konfig("Abschnittswechsel", Global.Modus.Read, configuration);

        if (art != Global.Zweck.Statistik)
        {
            var konferenzart = "";
            switch (configuration["Abschnitt"])
            {
                case "1":
                    konferenzart = "Halbjahres";
                    break;
                case "2":
                    konferenzart = "Jahres";
                    break;
                default:
                    throw new Exception("Ungültiger Abschnitt. Bitte geben Sie 1 oder 2 ein.");
            }

            configuration = Global.Konfig($"Konferenzdatum", Global.Modus.Update, configuration);
            konferenzdatum = DateTime.Parse(configuration["Konferenzdatum"]);
            configuration = Global.Konfig($"ZeugnisDatum", Global.Modus.Update, configuration);
            zeugnisdatum = DateTime.Parse(configuration["ZeugnisDatum"]);
            configuration = Global.Konfig("MaximaleAnzahlFehlstundenProTag", Global.Modus.Read, configuration);
            configuration = Global.Konfig("FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt", Global.Modus.Read, configuration);
        }

        try
        {
            AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("Lernabschnittsdaten.dat verarbeiten ...", ctx =>
            {
                foreach (var student in IStudents)
                {
                    var abschnittswechsel = DateTime.Parse(configuration["Abschnittswechsel"]);
                    var abschnittswechelInderZukunft = abschnittswechsel > DateTime.Now;

                    var dictBasisdaten = schuelerBasisd
                        .Where(recBasis =>
                        {
                            var dictBasis = (IDictionary<string, object>)recBasis;
                            return dictBasis["Nachname"].ToString() == student.Nachname &&
                                dictBasis["Vorname"].ToString() == student.Vorname &&
                                dictBasis["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                dictBasis["Jahr"].ToString() == Global.AktSj[0] &&
                                ((art == Global.Zweck.Statistik || abschnittswechelInderZukunft) ? true : dictBasis["Abschnitt"].ToString() == configuration["Abschnitt"]);
                        }).FirstOrDefault() as IDictionary<string, object>;

                    if (dictBasisdaten != null)
                    {
                        //if (dictBasisdaten["Status"] == "2")
                        if (true)
                        {
                            var versetzung = "";
                            var abschluss = "";
                            var klassenlehrer = Klassen.Where(rec => rec.Name == student.Klasse)
                                .Select(rec => rec.Klassenlehrer).FirstOrDefault();
                            var jahrgang = string.IsNullOrEmpty(dictBasisdaten!["Jahrgang"].ToString())
                                ? ""
                                : dictBasisdaten["Jahrgang"].ToString();
                            var schulgliederung = string.IsNullOrEmpty(dictBasisdaten["Schulgliederung"].ToString())
                                ? ""
                                : dictBasisdaten["Schulgliederung"].ToString();
                            var orgForm = string.IsNullOrEmpty(dictBasisdaten["OrgForm"].ToString())
                                ? ""
                                : dictBasisdaten["OrgForm"].ToString();
                            var klassenart = string.IsNullOrEmpty(dictBasisdaten["Klassenart"].ToString())
                                ? ""
                                : dictBasisdaten["Klassenart"].ToString();
                            var fachklasse = string.IsNullOrEmpty(dictBasisdaten["Fachklasse"].ToString())
                                ? ""
                                : dictBasisdaten["Fachklasse"].ToString();
                            var zeugnisart = "";
                            var schwerstbehinderung = student.Schwerstbehinderung;
                            var wiederholung = "";

                            var dictLernabschnitt = schuelerLernab
                                .Where(recLern =>
                                {
                                    var dictLern = (IDictionary<string, object>)recLern;
                                    return dictLern["Nachname"].ToString() == student.Nachname &&
                                        dictLern["Vorname"].ToString() == student.Vorname &&
                                        dictLern["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                        dictLern["Jahr"].ToString() == Global.AktSj[0] &&
                                        dictLern["Abschnitt"].ToString() == configuration["Abschnitt"];
                                }).FirstOrDefault() as IDictionary<string, object>;

                            var fehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt = Global.FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt;

                            // Wenn bereits Lernabschnittsdaten existieren, werden die Daten dort entnommen.
                            if (dictLernabschnitt != null)
                            {
                                konferenzdatum = konferenzdatum.Year == 1
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Konferenzdatum"].ToString())
                                        ? konferenzdatum
                                        : Convert.ToDateTime(dictLernabschnitt["Konferenzdatum"].ToString())
                                    : konferenzdatum;
                                zeugnisdatum = zeugnisdatum.Year == 1
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Zeugnisdatum"].ToString())
                                        ? zeugnisdatum
                                        : Convert.ToDateTime(dictLernabschnitt["Zeugnisdatum"].ToString())
                                    : zeugnisdatum;
                                jahrgang = string.IsNullOrEmpty(jahrgang)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Jahrgang"].ToString())
                                        ? jahrgang
                                        : dictLernabschnitt["Jahrgang"].ToString()
                                    : jahrgang;
                                orgForm = string.IsNullOrEmpty(orgForm)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["OrgForm"].ToString())
                                        ? orgForm
                                        : dictLernabschnitt["OrgForm"].ToString()
                                    : orgForm;
                                klassenart = string.IsNullOrEmpty(klassenart)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Klassenart"].ToString())
                                        ? klassenart
                                        : dictLernabschnitt["Klassenart"].ToString()
                                    : klassenart;
                                schulgliederung = string.IsNullOrEmpty(schulgliederung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Schulgliederung"].ToString())
                                        ? schulgliederung
                                        : dictLernabschnitt["Schulgliederung"].ToString()
                                    : schulgliederung;
                                klassenlehrer = string.IsNullOrEmpty(klassenlehrer)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Klassenlehrer"].ToString())
                                        ? klassenlehrer
                                        : dictLernabschnitt["Klassenlehrer"].ToString()
                                    : klassenlehrer;
                                versetzung = string.IsNullOrEmpty(versetzung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Versetzung"].ToString())
                                        ? versetzung
                                        : dictLernabschnitt["Versetzung"].ToString()
                                    : versetzung;
                                abschluss = string.IsNullOrEmpty(abschluss)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Abschluss"].ToString())
                                        ? abschluss
                                        : dictLernabschnitt["Abschluss"].ToString()
                                    : abschluss;
                                fachklasse = string.IsNullOrEmpty(fachklasse)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Fachklasse"].ToString())
                                        ? fachklasse
                                        : dictLernabschnitt["Fachklasse"].ToString()
                                    : fachklasse;
                                zeugnisart = string.IsNullOrEmpty(zeugnisart)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Zeugnisart"].ToString())
                                        ? zeugnisart
                                        : dictLernabschnitt["Zeugnisart"].ToString()
                                    : zeugnisart;
                                schwerstbehinderung = string.IsNullOrEmpty(schwerstbehinderung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Schwerstbehinderung"].ToString())
                                        ? schwerstbehinderung
                                        : dictLernabschnitt["Schwerstbehinderung"].ToString()
                                    : schwerstbehinderung;
                                wiederholung = string.IsNullOrEmpty(wiederholung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Wiederholung"].ToString())
                                        ? wiederholung
                                        : dictLernabschnitt["Wiederholung"].ToString()
                                    : wiederholung;
                            }

                            dynamic record = new ExpandoObject();
                            record.Nachname = $"{student.Nachname}#{student.Klasse}";
                            record.Vorname = student.Vorname;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Jahr = Global.AktSj[0];
                            record.Abschnitt = art == Global.Zweck.Statistik ? "1" : configuration["Abschnitt"];
                            record.Jahrgang = jahrgang;
                            record.Klasse = student.Klasse;
                            record.Schulgliederung = schulgliederung;
                            record.OrgForm = orgForm;
                            record.Klassenart = klassenart;
                            record.Fachklasse = fachklasse;
                            record.Förderschwerpunkt = "";
                            record.ZWEIPUNKTLEERZEICHENFörderschwerpunkt = "";
                            record.Schwerstbehinderung = schwerstbehinderung;
                            record.Wertung = "J";
                            record.Wiederholung = wiederholung;
                            record.Klassenlehrer = klassenlehrer;
                            record.Versetzung = versetzung;
                            record.Abschluss = abschluss;
                            record.Schwerpunkt = "";
                            record.Konferenzdatum = art == Global.Zweck.Statistik ? "" : konferenzdatum.ToShortDateString();
                            record.Zeugnisdatum = art == Global.Zweck.Statistik ? "" : zeugnisdatum.ToShortDateString();
                            record.SummeFehlstd = art == Global.Zweck.Statistik ? "" : student.GetFehlstd(absencePerStud, configuration);
                            record.SummeFehlstdUNTERSTRICHunentschuldigt = art == Global.Zweck.Statistik ? "" : student.GetUnentFehlstd(absencePerStud, configuration);
                            record.allgPUNKTMINUSbildenderLEERZEICHENAbschluss = "";
                            record.berufsbezPUNKTLEERZEICHENAbschluss = "";
                            record.Zeugnisart = zeugnisart;
                            record.FehlstundenMINUSGrenzwert = "";
                            record.DatumLEERZEICHENvon = "";
                            record.DatumLEERZEICHENbis = "";
                            zieldatei.Add(record);
                        }
                    }
                }
                Global.ZeileSchreiben("Lernabschnittsdaten.dat", zieldatei.Count().ToString());
            });
            
            foreach (var aktion in zieldatei.Funktionen)
                aktion(zieldatei);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            while (Console.KeyAvailable) Console.ReadKey(true);

            Console.ReadKey();
        }
    }

    public void Lernabschnittsdaten(
        Lehrers lehrers,
        IConfiguration configuration,
        Global.Zweck art,
        string zieldateiname,        
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        var schuelerLernab = Quelldateien.GetMatchingList(configuration, "schuelerlernabschnittsdaten", IStudents, Klassen);
        if (schuelerLernab == null || !schuelerLernab.Any()) return;

        var schuelerBasisd = Quelldateien.GetMatchingList(configuration, "schuelerbasisdate", IStudents, Klassen);
        if (schuelerBasisd == null || !schuelerBasisd.Any()) return;

        var absencePerStud = new List<dynamic>();

        var konferenzdatum = DateTime.Now;
        var zeugnisdatum = DateTime.Now;

        configuration = Global.Konfig("Abschnitt", Global.Modus.Read, configuration);
        configuration = Global.Konfig("Abschnittswechsel", Global.Modus.Read, configuration);

        if (art != Global.Zweck.Statistik)
        {
            absencePerStud = Quelldateien.GetMatchingList(configuration, "absenceperstudent", IStudents, Klassen);
            if (absencePerStud == null || !absencePerStud.Any()) return;

            // Wenn der Abschnittswechsel hinter uns liegt und gleichzeitig Fehlzeiten von vorher vorliegen, dann wird gefragt, ob die vorherigen Fehlzeiten berücksichtigt werden sollen. 
            var abschnittswechsel = DateTime.Parse(configuration["Abschnittswechsel"]);
            var abschnittswechelInderZukunft = abschnittswechsel > DateTime.Now;
            if (!abschnittswechelInderZukunft)
            {
                var alteFehlzeiten = absencePerStud
                    .Where(rec =>
                    {
                        if (rec == null) return false;
                        var dict = (IDictionary<string, object>)rec;
                        return dict != null && DateTime.Parse(dict["Datum"].ToString()) < abschnittswechsel;
                    }).Count();

                if (alteFehlzeiten > 0)
                    configuration = Global.Konfig("FehlzeitenVorDemAbschnittswechselBeruecksichtigen", Global.Modus.Update, configuration);

                // Wenn die Fehlzeiten vor dem Abschnittswechsel nicht berücksichtigt werden sollen, dann werden sie aus der Liste entfernt.
                if (!configuration["FehlzeitenVorDemAbschnittswechselBeruecksichtigen"].ToLower().StartsWith("j"))
                {
                    absencePerStud = absencePerStud
                        .Where(rec =>
                        {
                            if (rec == null) return false;
                            var dict = (IDictionary<string, object>)rec;
                            return dict != null && DateTime.Parse(dict["Datum"].ToString()) >= abschnittswechsel;
                        }).ToList();
                }
            }

            // Falls es offene Stunden gibt, kommt ein Hinweis, den man auch bestätigen muss.
            var ab = absencePerStud
                    .Where(rec =>
                    {
                        if (rec == null) return false;
                        var dict = (IDictionary<string, object>)rec;
                        return dict != null && dict["Status"] != null && dict["Status"].ToString() == "offen";
                    }).Count();

            if (ab > 0)
            {
                configuration = Global.Konfig("OffeneFehlstunden", Global.Modus.Update, configuration);
                if(!configuration["OffeneFehlstunden"].ToLower().StartsWith("j"))
                {                    
                    zieldatei.Auswählen(configuration, this, lehrers, Global.Modus.NurEineKlasse);
                    throw new Exception("Sie haben wegen offener Fehlstunden in Webuntis abgebrochen.");
                }
            }

            var konferenzart = "";
            switch (configuration["Abschnitt"])
            {
                case "1":
                    konferenzart = "Halbjahres";
                    break;
                case "2":
                    konferenzart = "Jahres";
                    break;
                default:
                    throw new Exception("Ungültiger Abschnitt. Bitte geben Sie 1 oder 2 ein.");
            }

            configuration = Global.Konfig($"Konferenzdatum", Global.Modus.Update, configuration);
            konferenzdatum = DateTime.Parse(configuration["Konferenzdatum"]);
            configuration = Global.Konfig($"ZeugnisDatum", Global.Modus.Update, configuration);
            zeugnisdatum = DateTime.Parse(configuration["ZeugnisDatum"]);
            configuration = Global.Konfig("MaximaleAnzahlFehlstundenProTag", Global.Modus.Read, configuration);
            configuration = Global.Konfig("FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt", Global.Modus.Read, configuration);
        }

        try
        {
            AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("Lernabschnittsdaten.dat verarbeiten ...", ctx =>
            {
                foreach (var student in IStudents)
                {
                    var abschnittswechsel = DateTime.Parse(configuration["Abschnittswechsel"]);
                    var abschnittswechelInderZukunft = abschnittswechsel > DateTime.Now;

                    var dictBasisdaten = schuelerBasisd
                        .Where(recBasis =>
                        {
                            var dictBasis = (IDictionary<string, object>)recBasis;
                            return dictBasis["Nachname"].ToString() == student.Nachname &&
                                dictBasis["Vorname"].ToString() == student.Vorname &&
                                dictBasis["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                dictBasis["Jahr"].ToString() == Global.AktSj[0] &&
                                ((art == Global.Zweck.Statistik || abschnittswechelInderZukunft) ? true : dictBasis["Abschnitt"].ToString() == configuration["Abschnitt"]);
                        }).FirstOrDefault() as IDictionary<string, object>;

                    if (dictBasisdaten != null)
                    {
                        //if (dictBasisdaten["Status"] == "2")
                        if (true)
                        {
                            var versetzung = "";
                            var abschluss = "";
                            var klassenlehrer = Klassen.Where(rec => rec.Name == student.Klasse)
                                .Select(rec => rec.Klassenlehrer).FirstOrDefault();
                            var jahrgang = string.IsNullOrEmpty(dictBasisdaten!["Jahrgang"].ToString())
                                ? ""
                                : dictBasisdaten["Jahrgang"].ToString();
                            var schulgliederung = string.IsNullOrEmpty(dictBasisdaten["Schulgliederung"].ToString())
                                ? ""
                                : dictBasisdaten["Schulgliederung"].ToString();
                            var orgForm = string.IsNullOrEmpty(dictBasisdaten["OrgForm"].ToString())
                                ? ""
                                : dictBasisdaten["OrgForm"].ToString();
                            var klassenart = string.IsNullOrEmpty(dictBasisdaten["Klassenart"].ToString())
                                ? ""
                                : dictBasisdaten["Klassenart"].ToString();
                            var fachklasse = string.IsNullOrEmpty(dictBasisdaten["Fachklasse"].ToString())
                                ? ""
                                : dictBasisdaten["Fachklasse"].ToString();
                            var zeugnisart = "";
                            var schwerstbehinderung = student.Schwerstbehinderung;
                            var wiederholung = "";

                            var dictLernabschnitt = schuelerLernab
                                .Where(recLern =>
                                {
                                    var dictLern = (IDictionary<string, object>)recLern;
                                    return dictLern["Nachname"].ToString() == student.Nachname &&
                                        dictLern["Vorname"].ToString() == student.Vorname &&
                                        dictLern["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                        dictLern["Jahr"].ToString() == Global.AktSj[0] &&
                                        dictLern["Abschnitt"].ToString() == configuration["Abschnitt"];
                                }).FirstOrDefault() as IDictionary<string, object>;

                            var fehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt = Global.FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt;

                            // Wenn bereits Lernabschnittsdaten existieren, werden die Daten dort entnommen.
                            if (dictLernabschnitt != null)
                            {
                                konferenzdatum = konferenzdatum.Year == 1
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Konferenzdatum"].ToString())
                                        ? konferenzdatum
                                        : Convert.ToDateTime(dictLernabschnitt["Konferenzdatum"].ToString())
                                    : konferenzdatum;
                                zeugnisdatum = zeugnisdatum.Year == 1
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Zeugnisdatum"].ToString())
                                        ? zeugnisdatum
                                        : Convert.ToDateTime(dictLernabschnitt["Zeugnisdatum"].ToString())
                                    : zeugnisdatum;
                                jahrgang = string.IsNullOrEmpty(jahrgang)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Jahrgang"].ToString())
                                        ? jahrgang
                                        : dictLernabschnitt["Jahrgang"].ToString()
                                    : jahrgang;
                                orgForm = string.IsNullOrEmpty(orgForm)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["OrgForm"].ToString())
                                        ? orgForm
                                        : dictLernabschnitt["OrgForm"].ToString()
                                    : orgForm;
                                klassenart = string.IsNullOrEmpty(klassenart)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Klassenart"].ToString())
                                        ? klassenart
                                        : dictLernabschnitt["Klassenart"].ToString()
                                    : klassenart;
                                schulgliederung = string.IsNullOrEmpty(schulgliederung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Schulgliederung"].ToString())
                                        ? schulgliederung
                                        : dictLernabschnitt["Schulgliederung"].ToString()
                                    : schulgliederung;
                                klassenlehrer = string.IsNullOrEmpty(klassenlehrer)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Klassenlehrer"].ToString())
                                        ? klassenlehrer
                                        : dictLernabschnitt["Klassenlehrer"].ToString()
                                    : klassenlehrer;
                                versetzung = string.IsNullOrEmpty(versetzung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Versetzung"].ToString())
                                        ? versetzung
                                        : dictLernabschnitt["Versetzung"].ToString()
                                    : versetzung;
                                abschluss = string.IsNullOrEmpty(abschluss)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Abschluss"].ToString())
                                        ? abschluss
                                        : dictLernabschnitt["Abschluss"].ToString()
                                    : abschluss;
                                fachklasse = string.IsNullOrEmpty(fachklasse)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Fachklasse"].ToString())
                                        ? fachklasse
                                        : dictLernabschnitt["Fachklasse"].ToString()
                                    : fachklasse;
                                zeugnisart = string.IsNullOrEmpty(zeugnisart)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Zeugnisart"].ToString())
                                        ? zeugnisart
                                        : dictLernabschnitt["Zeugnisart"].ToString()
                                    : zeugnisart;
                                schwerstbehinderung = string.IsNullOrEmpty(schwerstbehinderung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Schwerstbehinderung"].ToString())
                                        ? schwerstbehinderung
                                        : dictLernabschnitt["Schwerstbehinderung"].ToString()
                                    : schwerstbehinderung;
                                wiederholung = string.IsNullOrEmpty(wiederholung)
                                    ? string.IsNullOrEmpty(dictLernabschnitt["Wiederholung"].ToString())
                                        ? wiederholung
                                        : dictLernabschnitt["Wiederholung"].ToString()
                                    : wiederholung;
                            }

                            dynamic record = new ExpandoObject();
                            record.Nachname = $"{student.Nachname}#{student.Klasse}";
                            record.Vorname = student.Vorname;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Jahr = Global.AktSj[0];
                            record.Abschnitt = art == Global.Zweck.Statistik ? "1" : configuration["Abschnitt"];
                            record.Jahrgang = jahrgang;
                            record.Klasse = student.Klasse;
                            record.Schulgliederung = schulgliederung;
                            record.OrgForm = orgForm;
                            record.Klassenart = klassenart;
                            record.Fachklasse = fachklasse;
                            record.Förderschwerpunkt = "";
                            record.ZWEIPUNKTLEERZEICHENFörderschwerpunkt = "";
                            record.Schwerstbehinderung = schwerstbehinderung;
                            record.Wertung = "J";
                            record.Wiederholung = wiederholung;
                            record.Klassenlehrer = klassenlehrer;
                            record.Versetzung = versetzung;
                            record.Abschluss = abschluss;
                            record.Schwerpunkt = "";
                            record.Konferenzdatum = art == Global.Zweck.Statistik ? "" : konferenzdatum.ToShortDateString();
                            record.Zeugnisdatum = art == Global.Zweck.Statistik ? "" : zeugnisdatum.ToShortDateString();
                            record.SummeFehlstd = art == Global.Zweck.Statistik ? "" : student.GetFehlstd(absencePerStud, configuration);
                            record.SummeFehlstdUNTERSTRICHunentschuldigt = art == Global.Zweck.Statistik ? "" : student.GetUnentFehlstd(absencePerStud, configuration);
                            record.allgPUNKTMINUSbildenderLEERZEICHENAbschluss = "";
                            record.berufsbezPUNKTLEERZEICHENAbschluss = "";
                            record.Zeugnisart = zeugnisart;
                            record.FehlstundenMINUSGrenzwert = "";
                            record.DatumLEERZEICHENvon = "";
                            record.DatumLEERZEICHENbis = "";
                            zieldatei.Add(record);
                        }
                    }
                }
                Global.ZeileSchreiben("Lernabschnittsdaten.dat", zieldatei.Count().ToString());
            });
            
            foreach (var aktion in zieldatei.Funktionen)
                aktion(zieldatei);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            while (Console.KeyAvailable) Console.ReadKey(true);

            Console.ReadKey();
        }
    }

    public void LeistungsdatenStatistik(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null, Global.Zweck art = Global.Zweck.Statistik)
    {
        var unterrichte = this.Unterrichte;

        List<dynamic>? schuelerLeistungsdaten = null;
        List<dynamic>? marksPerLs = null;

        List<dynamic> schBasisds = Quelldateien.GetMatchingList(configuration, "schuelerbasisdaten", IStudents, Klassen);
        if (art == null && (schBasisds == null || schBasisds.Count == 0)) return;

        List<dynamic> schLeistus = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
        if (art == null && (schLeistus == null || schLeistus.Count == 0)) return;

        if(art == Global.Zweck.Zeugnis)
        {
            marksPerLs = Quelldateien.GetMatchingList(configuration, "marksperlesson", IStudents, Klassen);
            if (marksPerLs == null || marksPerLs.Count == 0) return;
        }

        if(art == Global.Zweck.Zeugnis)
            schuelerLeistungsdaten = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);        
        
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"{zieldateiname} ...", ctx =>
        {
            foreach (var unt in this.Unterrichte)
            {
                foreach (var student in unt.Students.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname))
                {
                    var kursart = unt.Kursart;
                    var abiturfach = "";

                    // Kursart und Abiturfach werden in SchILD direkt gepflegt und dürfen von BKB-Tool nicht überschrieben werden.
                    // Sie müssen aus den bisherigen schuelerleistungsdaten entnommen werden.
                    var leistungsdaten = schuelerLeistungsdaten
                            .Where(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    // Vergleiche nur den ersten Teil des Fachnamens, da es z.B. "E L1" und "E L2" geben kann.
                                    dict["Fach"].ToString().Split(' ')[0] == unt.Fach.Split(' ')[0] &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                            }).LastOrDefault() as IDictionary<string, object>;

                    if (leistungsdaten != null && !string.IsNullOrEmpty(leistungsdaten["Kursart"].ToString()))                    
                        kursart = leistungsdaten["Kursart"].ToString();
                    if (leistungsdaten != null && !string.IsNullOrEmpty(leistungsdaten["Abiturfach"].ToString()))
                        abiturfach = leistungsdaten["Abiturfach"].ToString();
                    
                    string jahrgang = student.GetJahrgang(schBasisds);
                    string note = student.GetNote(jahrgang, marksPerLs, unt.Fach, art);

                    dynamic record = new ExpandoObject();
                    record.Nachname = $"{student.Nachname}#{student.Klasse}";
                    record.Vorname = student.Vorname;
                    record.Geburtsdatum = student.Geburtsdatum;
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = configuration["Abschnitt"];
                    record.Fach = unt.Fach;
                    if (unt.Fach == "DPF")
                    {
                        var saw = "";
                    }
                    record.Fachlehrer = unt.Kursleiter;
                    record.Kursart = kursart;
                    record.Kurs = unt.KursBez;
                    record.Note = art == Global.Zweck.Statistik ? "" : note;
                    record.Abiturfach = abiturfach;
                    record.WochenstdPUNKT = unt.Wochenstunden.ToString();
                    record.ExterneLEERZEICHENSchulnrPUNKT = "";
                    record.Zusatzkraft = "";
                    record.WochenstdPUNKTLEERZEICHENZK = "";
                    record.Jahrgang = "";
                    record.Jahrgänge = "";
                    record.FehlstdPUNKT = "";
                    record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                    zieldatei.Add(record);
                }
            }
            Global.ZeileSchreiben("SchuelerLeistungsdaten.dat", zieldatei.Count().ToString());
        });

        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);
    }
    public void NotenInLeistungsdatenErgaenzen(
        IConfiguration configuration,
        Lehrers lehrers,
        string zieldateiname,        
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null, Global.Zweck art = Global.Zweck.Statistik)
    {   
        List<dynamic> schuelerLeistungsdaten = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
        if (art == null && (schuelerLeistungsdaten == null || schuelerLeistungsdaten.Count == 0)) return;
  
        List<dynamic>? marksPerLs = Quelldateien.GetMatchingList(configuration, "marksperlesson", IStudents, Klassen);
        if (marksPerLs == null || marksPerLs.Count == 0) return;
  
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        zieldatei.Lehrers = lehrers;

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"{zieldateiname} ...", ctx =>
        {
            for (int i = 0; i < schuelerLeistungsdaten.Count; i++)
            {
                var dictLeist = (IDictionary<string, object>)schuelerLeistungsdaten[i];
                var student = Students.Where(s => s.Nachname == dictLeist["Nachname"].ToString() && s.Vorname == dictLeist["Vorname"].ToString() && s.Geburtsdatum == dictLeist["Geburtsdatum"].ToString()).FirstOrDefault();    

                // Nur Zeilen des interessierenden Abschnitts berücksichtigen
                if(dictLeist["Abschnitt"].ToString() != configuration["Abschnitt"]) continue;
                
                // Wenn der Schüler nicht aktiv oder extern ist, überspringe diese Zeile
                if (!(student.Status.ToString() == "2" || student.Status.ToString() == "6")) continue;

                dynamic record = new ExpandoObject();

                foreach (var prop in dictLeist)
                {
                    var name = prop.Key;
                    var value = prop.Value;

                    if (name == "Nachname")
                    {
                        var klasse = student.Klasse;
                        if(string.IsNullOrEmpty(klasse)){student.Warnung("hat keine Klasse in Schild."); continue;}
                        
                        ((IDictionary<string, object>)record)[name] = $"{value}#{klasse}";
                    }
                    else if (name == "Note")
                    {
                        var jahrgang = student.Jahrgang;
                        if(string.IsNullOrEmpty(jahrgang)){student.Warnung("hat keinen Jahrgang in Schild."); continue;}
                        
                        var fach = dictLeist["Fach"].ToString();

                        string note = student.GetNote(jahrgang, marksPerLs, fach, art);
                        if(string.IsNullOrEmpty(note))
                        {
                            var lol = dictLeist["Fachlehrer"].ToString();

                            if(art != Global.Zweck.Mahnung)
                                student.Warnung($"{lol} hat keine Note erteilt in {fach}.");
                        }
                        
                        ((IDictionary<string, object>)record)[name] = note;
                    }
                    else
                    {
                        ((IDictionary<string, object>)record)[name] = value;
                    }
                }
                zieldatei.Add(record);
            }
            Global.ZeileSchreiben("SchuelerLeistungsdaten.dat", zieldatei.Count().ToString());
        });
        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);
    }
    public void UnterrichteAnlegen(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null, Global.Zweck art = Global.Zweck.Statistik)
    {
        var unterrichte = this.Unterrichte;

        List<dynamic>? schuelerLeistungsdaten = null;
        List<dynamic>? marksPerLs = null;

        List<dynamic> schBasisds = Quelldateien.GetMatchingList(configuration, "schuelerbasisdaten", IStudents, Klassen);
        if (art == null && (schBasisds == null || schBasisds.Count == 0)) return;

        List<dynamic> schLeistus = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
        if (art == null && (schLeistus == null || schLeistus.Count == 0)) return;

        if(art == Global.Zweck.Zeugnis)
            schuelerLeistungsdaten = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);        
        
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"{zieldateiname} ...", ctx =>
        {
            foreach (var unt in this.Unterrichte)
            {
                foreach (var student in unt.Students.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname))
                {
                    var kursart = unt.Kursart;
                    var abiturfach = "";

                    // Kursart und Abiturfach werden in SchILD direkt gepflegt und dürfen von BKB-Tool nicht überschrieben werden.
                    // Sie müssen aus den bisherigen schuelerleistungsdaten entnommen werden.
                    var leistungsdaten = schuelerLeistungsdaten
                            .Where(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    // Vergleiche nur den ersten Teil des Fachnamens, da es z.B. "E L1" und "E L2" geben kann.
                                    dict["Fach"].ToString().Split(' ')[0] == unt.Fach.Split(' ')[0] &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                            }).LastOrDefault() as IDictionary<string, object>;

                    if (leistungsdaten != null && !string.IsNullOrEmpty(leistungsdaten["Kursart"].ToString()))                    
                        kursart = leistungsdaten["Kursart"].ToString();
                    if (leistungsdaten != null && !string.IsNullOrEmpty(leistungsdaten["Abiturfach"].ToString()))
                        abiturfach = leistungsdaten["Abiturfach"].ToString();
                    
                    string jahrgang = student.GetJahrgang(schBasisds);
                    string note = "";

                    dynamic record = new ExpandoObject();
                    record.Nachname = $"{student.Nachname}#{student.Klasse}";
                    record.Vorname = student.Vorname;
                    record.Geburtsdatum = student.Geburtsdatum;
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = configuration["Abschnitt"];
                    record.Fach = unt.Fach;
                    record.Fachlehrer = unt.Kursleiter;
                    record.Kursart = kursart;
                    record.Kurs = unt.KursBez;
                    record.Note = "";
                    record.Abiturfach = abiturfach;
                    record.WochenstdPUNKT = unt.Wochenstunden.ToString();
                    record.ExterneLEERZEICHENSchulnrPUNKT = "";
                    record.Zusatzkraft = "";
                    record.WochenstdPUNKTLEERZEICHENZK = "";
                    record.Jahrgang = "";
                    record.Jahrgänge = "";
                    record.FehlstdPUNKT = "";
                    record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                    zieldatei.Add(record);
                }
            }
            Global.ZeileSchreiben("SchuelerLeistungsdaten.dat", zieldatei.Count().ToString());
        });

        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);
    }

    /*public Datei Leistungsdaten(IConfiguration configuration, string zieldateiname, Global.Zweck art)
    {
        var unterrichte = this.Unterrichte;
        var zieldatei = new Datei(zieldateiname);

        List<dynamic> marksPerLs = new List<dynamic>();
        
        var stdgroupSs = Quelldateien.GetMatchingList(configuration, "studentgroupstudents", IStudents, Klassen);
        if (art == null && (stdgroupSs == null || stdgroupSs.Count == 0)) return [];

        List<dynamic> schBasisds = Quelldateien.GetMatchingList(configuration, "schuelerbasisdaten", IStudents, Klassen);
        if (art == null && (schBasisds == null || schBasisds.Count == 0)) return [];

        List<dynamic> schLeistus = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
        if (art == null && (schLeistus == null || schLeistus.Count == 0)) return [];

        if (Global.Zweck.Statistik != art)
        {
            marksPerLs = Quelldateien.GetMatchingList(configuration, "marksperlesson", IStudents, Klassen);
            if (marksPerLs == null || marksPerLs.Count == 0) return [];
        }
        else
        {
            configuration["Abschnitt"] = "1";
        }

        var records = new List<dynamic>();

        if (art == Global.Zweck.Mahnung)
        {
            marksPerLs = marksPerLs.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Prüfungsart"].ToString().Contains("Mahnung");
            }).ToList();

            // Reduziere die IStudents-Liste basierend auf den gefilterten marksPerLs
            var x = IStudents.Where(x => x.Status == "2").Where(student =>
                marksPerLs.Any(mark =>
                {
                    var dict = (IDictionary<string, object>)mark;
                    return dict["Name"].ToString().Contains(student.Vorname) &&
                        dict["Name"].ToString().Contains(student.Nachname) &&
                        dict["Klasse"].ToString() == student.Klasse;
                })
            ).ToList();

            IStudents.Clear();
            IStudents.AddRange(x);
        }

        AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("SchuelerLeistungsdaten.dat verarbeiten ...", ctx =>
        {
            foreach (var klasse in IStudents.OrderBy(x => x.Klasse).Select(x => x.Klasse).Distinct())
            {
                var isFirstRun = true;

                var verschiedeneFaecherDerKlasse = VerschiedeneFaecher(klasse, expLessons);

                var religionWurdeUnterrichtet = verschiedeneFaecherDerKlasse
                    .Any(fach => new List<string>() { "rel", "kr", "er", "reli" }.Contains(fach.ToLower()));



                foreach (var student in IStudents.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname).Where(x => x.Klasse == klasse))
                {
                    var istReliabmelder = schBasisds.Any(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["Nachname"].ToString() == student.Nachname
                            && dict["Vorname"].ToString() == student.Vorname
                            && dict["Geburtsdatum"].ToString() == student.Geburtsdatum
                            && !string.IsNullOrEmpty(dict["Abmeldedatum Religionsunterricht"].ToString());
                    });

                    foreach (var fach in verschiedeneFaecherDerKlasse)
                    {
                        // Normalerweise gibt es nur einen Unterricht. 
                        var unterrichteMitDiesemFach = GetUnterrichteMitDiesemFach(fach, klasse, expLessons);

                        var dictExp = (IDictionary<string, object>)unterrichteMitDiesemFach[0];

                        var zusatzlehrkraft = "";
                        var zusatzlehrkraftWochenstunden = "";

                        // In der Statistikzählen allen Fächer mit, auch wenn sie nicht relevant sind.
                        if (art != Global.Zweck.Statistik)
                        {
                            if (!student.UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(dictExp, configuration)) continue;
                        }

                        // Wenn dieses Fach mit diesem Lehrer bereits in den records existiert,
                        // dann wird es nicht erneut hinzugefügt.

                        var gibtDasFachMitDemLehrerSchon = records.Any(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["Fach"].ToString() == dictExp["subject"].ToString() &&
                                dict["Fachlehrer"].ToString() == dictExp["teacher"].ToString() &&
                                dict["Vorname"].ToString() == student.Vorname &&
                                dict["Nachname"].ToString() == student.Nachname &&
                                dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                        });

                        if (!gibtDasFachMitDemLehrerSchon)
                        {
                            string jahrgang = student.GetJahrgang(schBasisds);
                            string note = student.GetNote(jahrgang, marksPerLs, dictExp["subject"].ToString()!, art);

                            // Wenn Reli unterricht wurde und der Schüler abgemeldet ist, dann wird NT eingesetzt.
                            if (
                                new List<string>() { "rel", "kr", "er", "reli" }.Contains(fach.ToLower())
                                && religionWurdeUnterrichtet)
                            {
                                if (istReliabmelder)
                                {
                                    note = "NT";
                                }
                            }

                            string kursart = GetKursart(configuration, jahrgang, fach);
                            bool mahnung = student.GetMahnung(marksPerLs, dictExp["subject"].ToString()!);

                            // Die Kursart 
                            var kursartBisher = schLeistus
                                .Where(record =>
                                {
                                    var dict = (IDictionary<string, object>)record;
                                    return dict["Vorname"].ToString() == student.Vorname &&
                                        dict["Nachname"].ToString() == student.Nachname &&
                                        dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                        dictExp["subject"] != null &&
                                        dict["Fach"].ToString() == dictExp["subject"].ToString();
                                })
                                .Select(record =>
                                {
                                    var dict = (IDictionary<string, object>)record;
                                    return dict["Kursart"].ToString();
                                })
                                .FirstOrDefault()
                                ?.ToString();

                            if (!string.IsNullOrEmpty(kursartBisher))
                                kursart = kursartBisher;

                            // Klassenunterrichte und Religion wird immer hinzugefügt
                            if (dictExp["studentgroup"].ToString() == "" || new List<string>() { "rel", "kr", "er", "reli", "religion", "rel1" }.Contains(fach.ToLower()))
                            {
                                dynamic record = new ExpandoObject();
                                record.Nachname = $"{student.Nachname}#{klasse}";
                                record.Vorname = student.Vorname;
                                record.Geburtsdatum = student.Geburtsdatum;
                                record.Jahr = Global.AktSj[0];
                                record.Abschnitt = configuration["Abschnitt"];
                                record.Fach = dictExp["subject"].ToString();
                                record.Fachlehrer = dictExp["teacher"].ToString();
                                record.Kursart = kursart;
                                record.Kurs = "";
                                record.Note = art == Global.Zweck.Statistik ? "" : note;
                                record.Abiturfach = "";
                                record.WochenstdPUNKT = dictExp["periods"];
                                record.ExterneLEERZEICHENSchulnrPUNKT = "";
                                record.Zusatzkraft = zusatzlehrkraft;
                                record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                                record.Jahrgang = "";
                                record.Jahrgänge = "";
                                record.FehlstdPUNKT = ""; // Fehlzeiten werden über die Abschnittsdaten importiert.
                                record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                                if (art == Global.Zweck.Mahnung)
                                {
                                    record.Mahnung = "J";
                                    record.Sortierung = "";
                                    record.Mahndatum = "";//DateTime.Now.ToShortDateString();
                                }
                                if ((mahnung && art == Global.Zweck.Mahnung) || art != Global.Zweck.Mahnung)
                                {
                                    records.Add(record);
                                }
                            }
                            else // Bei Kursunterrichten wird geschaut, ob der Schüler den Kurs belegt hat. 
                            {
                                var id = student.Id;
                                var studentZeile = stdgroupSs
                                    .Where(record =>
                                    {
                                        var dict = (IDictionary<string, object>)record;
                                        return dict["studentId"].ToString() == id &&
                                            dict["studentgroup.name"].ToString() ==
                                            dictExp["studentgroup"].ToString();
                                    })
                                    .FirstOrDefault();
                                var dictStudentgroup = (IDictionary<string, object>)studentZeile!;

                                if (dictStudentgroup != null)
                                {
                                    if (!student.UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(dictStudentgroup, configuration))
                                        continue;
                                    dynamic record = new ExpandoObject();
                                    record.Nachname = $"{student.Nachname}#{klasse}";
                                    record.Vorname = student.Vorname;
                                    record.Geburtsdatum = student.Geburtsdatum;
                                    record.Jahr = Global.AktSj[0];
                                    record.Abschnitt = configuration["Abschnitt"];
                                    record.Fach = dictStudentgroup["subject"].ToString();
                                    record.Fachlehrer = dictExp["teacher"].ToString();
                                    record.Kursart = kursart;
                                    record.Kurs = dictStudentgroup["studentgroup.name"].ToString()!.Substring(0,
                                        Math.Min(dictStudentgroup["studentgroup.name"].ToString()!.Length, 20));
                                    record.Note = note;
                                    record.Abiturfach = "";
                                    record.WochenstdPUNKT = dictExp["periods"];
                                    record.ExterneLEERZEICHENSchulnrPUNKT = "";
                                    record.Zusatzkraft = zusatzlehrkraft;
                                    record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                                    record.Jahrgang = student.Jahrgang;
                                    record.Jahrgänge = "";
                                    record.FehlstdPUNKT = "";
                                    record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                                    if (art == Global.Zweck.Mahnung)
                                    {
                                        record.Mahnung = "";
                                        record.Sortierung = "";
                                        record.Mahndatum = "";//DateTime.Now.ToShortDateString();
                                    }
                                    if ((mahnung && art == Global.Zweck.Mahnung) || art != Global.Zweck.Mahnung)
                                    {
                                        records.Add(record);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Global.ZeileSchreiben("SchuelerLeistungsdaten.dat", records.Count().ToString());
        });

        zieldatei.AddRange(records);
        return zieldatei;
    }*/

    public void Kurse(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);
        var kurse = this.Unterrichte.Where(x => !string.IsNullOrEmpty(x.KursBez)).ToList();        

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Kurse bilden: ...", ctx =>
        {
            foreach (var kurs in kurse)
            {
                dynamic record = new ExpandoObject();
                record.KursBez = kurs.KursBez.Substring(0, Math.Min(kurs.KursBez.Length, 20));

                // record.Klasse muss leer bleiben, wenn Schülergruppe verwendet werden.
                // Anderenfalls werden alle SuS aller Klassen zugewiesen.
                record.Klasse = ""; //string.IsNullOrEmpty(kurs.Schülergruppe) ? string.Join(",", kurs.Klassen) : "";

                record.Jahr = Global.AktSj[0];
                record.Abschnitt = configuration["Abschnitt"].ToString();
                record.Jahrgang = ""; //kann leer bleiben
                record.Fach = kurs.Fach;
                record.Kursart = kurs.Kursart;
                record.WochenstdPUNKT = kurs.Wochenstunden;
                record.WochenstdPUNKTLEERZEICHENKL = kurs.KursleiterWochenstunden;
                record.Kursleiter = kurs.Kursleiter;
                record.Epochenunterricht = "";
                record.Schulnr = "";
                // Falls es eine Zusatzkraft gibt
                if (kurs.Lehrkraefte.Count > 0)
                {
                    record.WochenstdPUNKTLEERZEICHENZK = kurs.LehrkraefteWochenstunden[0];
                    record.Zusatzkraft = kurs.Lehrkraefte[0];
                }
                if (kurs.Lehrkraefte.Count > 1)
                {
                    for (int i = 1; i < kurs.Lehrkraefte.Count; i++)
                    {
                        ((IDictionary<string, object>)record)[$"WochenstdLEERZEICHENZK{i}"] = kurs.LehrkraefteWochenstunden[i];
                        ((IDictionary<string, object>)record)[$"WeitereLEERZEICHENZusatzkraft{i}"] = kurs.Lehrkraefte[i];
                    }
                }
                zieldatei.Add(record);
            }

            Global.ZeileSchreiben("Kurse", zieldatei.Count().ToString());
        });

        foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
    }

    private List<string> GetDistincteKombinationenUnrPlusFachOhneSchuelergruppen(List<dynamic> gpu002)
    {
        var mehrfachVorkommendeField1 = gpu002
            .Where(record =>
            {
                var dict = (IDictionary<string, object>)record;
                // Field42 muss null oder leer sein
                return string.IsNullOrEmpty(dict.ContainsKey("Field42") ? dict["Field42"]?.ToString() : "");
            })
            .Select(record =>
            {
                var dict = (IDictionary<string, object>)record;
                return dict.ContainsKey("Field1") ? dict["Field1"]?.ToString() : null;
            })
            .Where(val => !string.IsNullOrEmpty(val))
            .GroupBy(val => val)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .Distinct()
            .ToList();

        return gpu002
            .Where(record =>
            {
                var dict = (IDictionary<string, object>)record;
                var field1 = dict.ContainsKey("Field1") ? dict["Field1"]?.ToString() : "";
                // Field42 muss null oder leer sein UND Field1 muss mehrfach vorkommen
                return string.IsNullOrEmpty(dict.ContainsKey("Field42") ? dict["Field42"]?.ToString() : "")
                    && !string.IsNullOrEmpty(field1)
                    && mehrfachVorkommendeField1.Contains(field1);
            })
            .Select(record =>
            {
                var dict = (IDictionary<string, object>)record;
                var field1 = dict.ContainsKey("Field1") ? dict["Field1"]?.ToString() : "";
                var field7 = dict.ContainsKey("Field7") ? dict["Field7"]?.ToString() : "";
                return $"{field7}-{field1}";
            })
            .Where(val => !string.IsNullOrEmpty(val) && val != "-")
            .Distinct()
            .ToList();
    }

    private List<string> GetDistincteKombinationenUnrPlusFachBeiSchuelergruppen(List<dynamic> gpu002)
    {
        



        return gpu002
                .Where(record =>
                {
                    var dict = (IDictionary<string, object>)record;
                    return dict.ContainsKey("Field42") && !string.IsNullOrEmpty(dict["Field42"]?.ToString());
                })
                .Select(record =>
                {
                    var dict = (IDictionary<string, object>)record;
                    var field1 = dict.ContainsKey("Field1") ? dict["Field1"]?.ToString() : "";
                    var field7 = dict.ContainsKey("Field7") ? dict["Field7"]?.ToString() : "";
                    return $"{field7}-{field1}";
                })
                .Distinct()
                .ToList();        
    }

    public Datei Leistungsdaten(IConfiguration configuration, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        
        var stdgroupSs = Quelldateien.GetMatchingList(configuration, "studentgroupstudents", IStudents, Klassen);
        if (stdgroupSs == null || stdgroupSs.Count == 0) return [];

        List<dynamic> gpu002 = Quelldateien.GetMatchingList(configuration, "gpu002", IStudents, Klassen);
        if (gpu002 == null || gpu002.Count == 0) return [];

        List<dynamic> schLeistus = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
        if (schLeistus == null || schLeistus.Count == 0) return [];
        
        configuration["Abschnitt"] = "1";

        var records = new List<dynamic>();

        // Wenn in der GPU002 zwei Zeilen mit der gleichen ID in Column1 stehen, dann ist das in Schild zwingend ein Kurs.
        // Wenn eine Schülergruppe in der GPU002 existiert, wird die Untis-Schülergruppe zum Schild-Kursnamen.
        // Wenn keine Schülergruppe existiert, dann heißt der Kurs U1234, wobei 1234 die Untis-Kurs-ID ist und 2025 das Schuljahr. 
        // Wenn ID und Raum bei mehreren KuK identisch sind, dann Teamteaching. 

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("SchuelerLeistungsdaten.dat verarbeiten ...", ctx =>
        {
            foreach (var klasse in IStudents.OrderBy(x => x.Klasse).Select(x => x.Klasse).Distinct())
            {
                var isFirstRun = true;
                var verschiedeneFaecherDerKlasse = VerschiedeneFaecher(klasse, gpu002);
                var religionWurdeUnterrichtet = verschiedeneFaecherDerKlasse.Any(fach => new List<string>() { "rel", "kr", "er", "reli" }.Contains(fach.ToLower()));

                foreach (var student in IStudents.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname).Where(x => x.Klasse == klasse))
                {
                    // Prüfen, ob der Schüler in der GPU002 als Reliabmelder eingetragen ist.






                    var istReliabmelder = gpu002.Any(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["Nachname"].ToString() == student.Nachname
                            && dict["Vorname"].ToString() == student.Vorname
                            && dict["Geburtsdatum"].ToString() == student.Geburtsdatum
                            && !string.IsNullOrEmpty(dict["Abmeldedatum Religionsunterricht"].ToString());
                    });

                    foreach (var fach in verschiedeneFaecherDerKlasse)
                    {
                        // Normalerweise gibt es nur einen Unterricht. 
                        var unterrichteMitDiesemFach = GetUnterrichteMitDiesemFach(fach, klasse, gpu002);

                        var dictExp = (IDictionary<string, object>)unterrichteMitDiesemFach[0];

                        var zusatzlehrkraft = "";
                        var zusatzlehrkraftWochenstunden = "";

                        // Wenn dieses Fach mit diesem Lehrer bereits in den records existiert,
                        // dann wird es nicht erneut hinzugefügt.

                        var gibtDasFachMitDemLehrerSchon = records.Any(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["Fach"].ToString() == dictExp["subject"].ToString() &&
                                dict["Fachlehrer"].ToString() == dictExp["teacher"].ToString() &&
                                dict["Vorname"].ToString() == student.Vorname &&
                                dict["Nachname"].ToString() == student.Nachname &&
                                dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                        });

                        if (!gibtDasFachMitDemLehrerSchon)
                        {
                            string jahrgang = student.GetJahrgang(gpu002);                            

                            string kursart = GetKursart(configuration, jahrgang, fach);

                            // Die Kursart 
                            var kursartBisher = schLeistus
                                .Where(record =>
                                {
                                    var dict = (IDictionary<string, object>)record;
                                    return dict["Vorname"].ToString() == student.Vorname &&
                                        dict["Nachname"].ToString() == student.Nachname &&
                                        dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                        dictExp["subject"] != null &&
                                        dict["Fach"].ToString() == dictExp["subject"].ToString();
                                })
                                .Select(record =>
                                {
                                    var dict = (IDictionary<string, object>)record;
                                    return dict["Kursart"].ToString();
                                })
                                .FirstOrDefault()
                                ?.ToString();

                            if (!string.IsNullOrEmpty(kursartBisher))
                                kursart = kursartBisher;

                            // Klassenunterrichte und Religion wird immer hinzugefügt
                            if (dictExp["studentgroup"].ToString() == "" || new List<string>() { "rel", "kr", "er", "reli", "religion", "rel1" }.Contains(fach.ToLower()))
                            {
                                dynamic record = new ExpandoObject();
                                record.Nachname = $"{student.Nachname}#{klasse}";
                                record.Vorname = student.Vorname;
                                record.Geburtsdatum = student.Geburtsdatum;
                                record.Jahr = Global.AktSj[0];
                                record.Abschnitt = configuration["Abschnitt"];
                                record.Fach = dictExp["subject"].ToString();
                                record.Fachlehrer = dictExp["teacher"].ToString();
                                record.Kursart = kursart;
                                record.Kurs = "";
                                record.Note = "";
                                record.Abiturfach = "";
                                record.WochenstdPUNKT = dictExp["periods"];
                                record.ExterneLEERZEICHENSchulnrPUNKT = "";
                                record.Zusatzkraft = zusatzlehrkraft;
                                record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                                record.Jahrgang = "";
                                record.Jahrgänge = "";
                                record.FehlstdPUNKT = ""; // Fehlzeiten werden über die Abschnittsdaten importiert.
                                record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                                record.Mahnung = "J";
                                record.Sortierung = "";
                                record.Mahndatum = "";//DateTime.Now.ToShortDateString();
                                records.Add(record);
                            }
                            else // Bei Kursunterrichten wird geschaut, ob der Schüler den Kurs belegt hat. 
                            {
                                var id = student.Id;
                                var studentZeile = stdgroupSs
                                    .Where(record =>
                                    {
                                        var dict = (IDictionary<string, object>)record;
                                        return dict["studentId"].ToString() == id &&
                                            dict["studentgroup.name"].ToString() ==
                                            dictExp["studentgroup"].ToString();
                                    })
                                    .FirstOrDefault();
                                var dictStudentgroup = (IDictionary<string, object>)studentZeile!;

                                if (dictStudentgroup != null)
                                {
                                    if (!student.UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(dictStudentgroup, configuration))
                                        continue;
                                    dynamic record = new ExpandoObject();
                                    record.Nachname = $"{student.Nachname}#{klasse}";
                                    record.Vorname = student.Vorname;
                                    record.Geburtsdatum = student.Geburtsdatum;
                                    record.Jahr = Global.AktSj[0];
                                    record.Abschnitt = configuration["Abschnitt"];
                                    record.Fach = dictStudentgroup["subject"].ToString();
                                    record.Fachlehrer = dictExp["teacher"].ToString();
                                    record.Kursart = kursart;
                                    record.Kurs = dictStudentgroup["studentgroup.name"].ToString()!.Substring(0,
                                        Math.Min(dictStudentgroup["studentgroup.name"].ToString()!.Length, 20));
                                    record.Note = "";
                                    record.Abiturfach = "";
                                    record.WochenstdPUNKT = dictExp["periods"];
                                    record.ExterneLEERZEICHENSchulnrPUNKT = "";
                                    record.Zusatzkraft = zusatzlehrkraft;
                                    record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                                    record.Jahrgang = student.Jahrgang;
                                    record.Jahrgänge = "";
                                    record.FehlstdPUNKT = "";
                                    record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                                    records.Add(record);                                    
                                }
                            }
                        }
                    }
                }
            }
            Global.ZeileSchreiben("SchuelerLeistungsdaten.dat", records.Count().ToString());
        });

        zieldatei.AddRange(records);
        return zieldatei;
    }

    private List<dynamic>? GetUnterrichteMitDiesemFach(string fach, string klasse, List<dynamic>? exportLessons)
    {
        return exportLessons.Where(rec =>
        {
            var dictExp = (IDictionary<string, object>)rec;
            return dictExp["subject"].ToString() == fach && dictExp["klassen"].ToString().Split('~').Contains(klasse);
        }).OrderByDescending(rec =>
        {
            var dictExp = (IDictionary<string, object>)rec;
            return DateTime.ParseExact(dictExp["endDate"].ToString()!, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        }).ToList();
    }

    public List<string> VerschiedeneFaecher(string klasse, List<dynamic>? exportLessons)
    {
        return exportLessons.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["klassen"].ToString().Split('~').Contains(klasse);
            })
            .Select(record =>
            {
                var dict = (IDictionary<string, object>)record;
                return dict["subject"].ToString();
            }).Distinct().ToList();
    }
    
    static bool IsValidHttpUrl(string url)
    {
        return Uri.TryCreate(url, UriKind.Absolute, out Uri? uriResult) &&
               (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
    }


    public Datei? Faecher(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        var gpu002 = Quelldateien.GetMatchingList(configuration, "gpu002", IStudents, Klassen);
        if (!gpu002.Any())
        {
            return [];
        }

        var schildFaecher = Quelldateien.GetMatchingList(configuration, "faecher", IStudents, Klassen);
        if (!schildFaecher.Any())
        {
            return [];
        }
        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Fächer ...", ctx =>
        {
            foreach (var recExp in gpu002)
            {
                var dictExp = (IDictionary<string, object>)recExp;

                // Leere Fächer überspringen
                if (string.IsNullOrEmpty(dictExp["Field7"].ToString()))
                {
                    var panel = new Panel($"[red]In Untis ist in der Unterrichtsnummer {dictExp["Field1"]} kein Fach eingetragen. Diese Zeile wird übersprungen.[/]");
                    panel.Border = BoxBorder.Double;
                    AnsiConsole.Write(panel);
                    continue;
                }

                var schildFach = schildFaecher.Where(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["InternKrz"].ToString() == dictExp["Field7"].ToString();
                }).FirstOrDefault();

                // Wenn es das Fach in SchILD nicht gibt, ...

                if (schildFach != null) continue;
                {
                    // ... wird bei Fächern mit Suffix geprüft, ob es bereits ein Schildfach ohne Suffix gibt.

                    var subject = dictExp["Field7"].ToString();
                    var endetMitZiffer = subject.Length > 0 && char.IsDigit(subject[^1]);

                    if (endetMitZiffer)
                    {
                        var subjectOhneSuffix = subject.Substring(0, subject.Length - 1);
                        // Die Eigenschaften vom Mutterfach werden übernommen
                        var mutterfach = schildFaecher.Where(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternKrz"].ToString() == subjectOhneSuffix;
                        }).FirstOrDefault();

                        // Wenn es ein Mutterfach gibt, wird es mit neuem Namen hinzugefügt
                        if (mutterfach != null)
                        {
                            if (zieldatei.Any(x => x.InternKrz == dictExp["Field7"].ToString())) continue;
                            var dictMutterfach = (IDictionary<string, object>)mutterfach;
                            dynamic record = new ExpandoObject();
                            record.InternKrz = subject;
                            record.StatistikKrz = dictMutterfach["StatistikKrz"].ToString();
                            record.Bezeichnung = dictMutterfach["Bezeichnung"].ToString();
                            record.BezeichnungZeugnis = dictMutterfach["BezeichnungZeugnis"].ToString();
                            record.BezeichnungÜZeugnis = dictMutterfach["BezeichnungÜZeugnis"].ToString();
                            record.Unterrichtssprache = dictMutterfach["Unterrichtsprache"].ToString();
                            record.SortierungLEERZEICHENS1 = dictMutterfach["Sortierung S1"].ToString();
                            record.SortierungLEERZEICHENS2 = dictMutterfach["Sortierung S2"].ToString();
                            record.Gewichtung = dictMutterfach["Gewichtung"].ToString();

                            var gibtEsSchon = zieldatei.Any(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["InternKrz"].ToString() == subject;
                            });

                            if (!gibtEsSchon)
                            {
                                zieldatei.Add(record);
                            }
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.WriteLine("  Das Fach " + subject +
                                            " nicht gefunden. Es wird in SchILD angelegt. Bitte prüfen!");


                            dynamic record = new ExpandoObject();
                            record.InternKrz = subject;
                            record.StatistikKrz = "FB";
                            record.Bezeichnung = subject;
                            record.BezeichnungZeugnis = "";
                            record.BezeichnungÜZeugnis = "";
                            record.Unterrichtssprache = "";
                            record.SortierungLEERZEICHENS1 = "";
                            record.SortierungLEERZEICHENS2 = "";
                            record.Gewichtung = "";

                            var gibtEsSchon = zieldatei.Any(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["InternKrz"].ToString() == subject;
                            });

                            if (!gibtEsSchon)
                            {
                                zieldatei.Add(record);
                            }
                        }
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine("  Das Fach " + subject +
                                        " nicht gefunden. Es wird in SchILD angelegt. Bitte prüfen!");


                        dynamic record = new ExpandoObject();
                        record.InternKrz = subject;
                        record.StatistikKrz = "FB";
                        record.Bezeichnung = subject;
                        record.BezeichnungZeugnis = "";
                        record.BezeichnungÜZeugnis = "";
                        record.Unterrichtssprache = "";
                        record.SortierungLEERZEICHENS1 = "";
                        record.SortierungLEERZEICHENS2 = "";
                        record.Gewichtung = "";

                        var gibtEsSchon = zieldatei.Any(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternKrz"].ToString() == subject;
                        });

                        if (!gibtEsSchon)
                        {
                            zieldatei.Add(record);
                        }
                    }
                }
            }
        });
        //Global.ZeileSchreiben("Neue Fächer:", zieldatei.Count().ToString());

        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);

        return zieldatei;
    }

    public string LuLAnEintragungDerZeugnisnotenErinnern(IConfiguration configuration, Lehrers lehrers)
    {
        var leistungsdaten = Quelldateien.GetMatchingList(configuration, "leistungsdaten", IStudents, Klassen);
        var betreff = "";
        var adressen = "";
        var anrede = "";
        var lul = new List<string?>();
        var eintaege = new List<string>();

        int i = 1;
        

        AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("Lehrkräfte verarbeiten ...", ctx =>
        {
            foreach (var leistungsdatum in leistungsdaten)
            {
                var dict = (IDictionary<string, object>)leistungsdatum;

                // Wenn keine Note erteilt wurde ...
                if (dict["Note"].ToString() == "")
                {
                    var student = IStudents.FirstOrDefault(x =>
                        x.Vorname == dict["Vorname"].ToString() && x.Nachname == dict["Nachname"].ToString() &&
                        x.Geburtsdatum == dict["Geburtsdatum"].ToString());

                    eintaege.Add(i.ToString().PadLeft(4) + ". " + student.Klasse.ToString().PadRight(6) + ", " +
                                (dict["Nachname"] + ", " + dict["Vorname"]).ToString().PadRight(20).Substring(0, 19) +
                                ": " + dict["Fachlehrer"].ToString().PadRight(3) + ": " + dict["Fach"]);
                    i++;

                    var lehrer = lehrers.FirstOrDefault(x => x.Kürzel == dict["Fachlehrer"].ToString());

                    if (lehrer != null && !string.IsNullOrEmpty(lehrer.Mail) && !lul.Contains(lehrer.Mail))
                    {
                        lul.Add(lehrer.Mail);
                    }
                }
            }

            if (eintaege.Count > 0)
            {
                foreach (var eintrag in eintaege)
                {
                    Console.WriteLine(eintrag);
                }

                foreach (var lehrer in lehrers)
                {
                    if (lul.Contains(lehrer.Kürzel))
                    {
                        adressen += lehrer.Mail + ",";
                    }
                }

                Console.WriteLine("   " + adressen.TrimEnd(','));
            }
            else
            {
                Console.WriteLine("  Es fehlen keine Noten. Gut so.");
            }
            Global.ZeileSchreiben("Lehrkräfte", i.ToString());
        });
        return String.Join(',', lul);
    }

    public void ChatErzeugen(IConfiguration configuration, String mitgliederMail)
    {
        
    }
    
    public void WebuntisOderNetmanOderLitteraCsv(IConfiguration configuration, List<Datei> zieldateien)
    {
        try
        {
            List<dynamic>? webuntisStudents = Quelldateien.GetMatchingList(configuration, "student_", Students, Klassen);
            if (webuntisStudents == null || webuntisStudents.Count == 0) throw new Exception("Keine Webuntis-Students.csv");
            var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", Students, Klassen);
            if (schuelerZusatzdaten == null || schuelerZusatzdaten.Count == 0) throw new Exception("Keine schuelerZusatzdaten.dat");
            var schuelerLernabschnittsdaten = Quelldateien.GetMatchingList(configuration, "schuelerlernabschnittsdaten", Students, Klassen);
            if (schuelerLernabschnittsdaten == null || schuelerLernabschnittsdaten.Count == 0) throw new Exception("Keine schuelerlernabschnittsdaten.dat");
            var schuelerErzieher = Quelldateien.GetMatchingList(configuration, "schuelererzieher", Students, Klassen);
            if (schuelerErzieher == null || schuelerErzieher.Count == 0) throw new Exception("Keine schuelerErzieher.dat");
            var schuelerAdressen = Quelldateien.GetMatchingList(configuration, "schueleradressen", Students, Klassen);
            if (schuelerAdressen == null || schuelerAdressen.Count == 0) throw new Exception("Keine Schueleradressen.dat");
            var lehrkraefte = Quelldateien.GetMatchingList(configuration, "lehrkraefte", Students, Klassen);
            if (lehrkraefte == null || lehrkraefte.Count == 0) throw new Exception("Keine Lehrkraefte.dat");
            var klassen = Quelldateien.GetMatchingList(configuration, "klassen", Students, Klassen);
            if (klassen == null || klassen.Count == 0) return;
            var schuelerTelefonnummern = Quelldateien.GetMatchingList(configuration, "schuelertelefonnummern", IStudents, Klassen);
            if (schuelerTelefonnummern == null || !schuelerTelefonnummern.Any()) return;

            var table = new Spectre.Console.Table();
            table.Border(TableBorder.Rounded);
            table.Centered();
            table.Expand();

            // Add columns 
            table.AddColumn("Nr");
            table.AddColumn("Name");
            table.AddColumn("Vorname");
            table.AddColumn("ID");
            table.AddColumn("Klasse");
            table.AddColumn("Status");
            table.AddColumn("Änderung");

            var susMitÄnderung = new List<string>()
            {
                "Folgende Änderungen / Neuanlagen:",
                "Nr".PadRight(5) + "Name".PadRight(46) + "Status Änderung".PadRight(20)
            };

            var i = 1;
            var gelöschteSchüler = new List<IDictionary<string, object>>();

            // Ersetze die bisherige DistinctBy-Zeile durch folgende Gruppierung und Auswahlregel:
            //
            // Wenn derselbe Schüler mehrfach vorkommt UND in mindestens zwei der Einträge der Status "2" oder "6" ist,
            // dann wird aus diesen aktiven/Gast-Einträgen derjenige übernommen, dessen `BeginnDesBildungsganges` am spätesten ist.
            // Andernfalls (kein mehrfacher aktiver/externer Fall) wird weiterhin der Eintrag mit dem numerisch kleinsten Status gewählt
            // (wie bisher an anderen Stellen des Codes verwendet).
            //
            // Diese Logik an den beiden Stellen im File anwenden, an denen `uniqueStudents` gebildet wird.

            var uniqueStudents = Students
                .GroupBy(s => new { s.Vorname, s.Nachname, s.Geburtsdatum })
                .Select(g =>
                {
                    var list = g.ToList();
                    if (list.Count == 1) return list[0];

                    DateTime ParseDateOrMin(string? dt)
                    {
                        if (string.IsNullOrWhiteSpace(dt)) return DateTime.MinValue;
                        var formats = new[] { "dd.MM.yyyy", "d.M.yyyy", "yyyy-MM-dd", "yyyyMMdd" };
                        if (DateTime.TryParseExact(dt, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed)) return parsed;
                        if (DateTime.TryParse(dt, new CultureInfo("de-DE"), DateTimeStyles.None, out parsed)) return parsed;
                        if (DateTime.TryParse(dt, out parsed)) return parsed;
                        return DateTime.MinValue;
                    }

                    // Kandidaten mit Status aktiv (2) oder Gast/extern (6)
                    var activeOrGuest = list.Where(s => s.Status == "2" || s.Status == "6").ToList();
                    if (activeOrGuest.Count > 1)
                    {
                        // Wähle den Eintrag mit dem spätesten BeginnDesBildungsganges
                        return activeOrGuest.OrderByDescending(s => ParseDateOrMin(s.BeginnDesBildungsganges)).First();
                    }

                    // Fallback: wie bisher – wähle den Eintrag mit dem numerisch kleinsten Status
                    return list.OrderBy(s => int.TryParse(s.Status, out var st) ? st : int.MaxValue).First();
                })
                .OrderBy(s => s.Klasse)
                .ThenBy(s => s.Nachname)
                .ThenBy(s => s.Vorname);

            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Webuntis-Schüler*innen vorbereiten ...", ctx =>                
            {                    
                foreach (var rec in webuntisStudents)
                {                
                    if (rec is not IDictionary<string, object> webuntisStudent) continue;

                    // Ein Abgänger oder Absolvent, der als Externer oder Gast wiederkommt, hat in seiner aktiven Rolle den niedrigeren Status.
                    // Also wird bei mehrfach vorkommenden Schülern immer der mit dem niedrigsten Status genommen.                 
                    var schildStudent = uniqueStudents
                        .OrderBy(x => int.TryParse(x.Status, out var status) ? status : 0)
                        .FirstOrDefault(x =>
                            x.Nachname == webuntisStudent["longName"].ToString() &&
                            x.Vorname == webuntisStudent["foreName"].ToString() &&
                            x.Geburtsdatum == webuntisStudent["birthDate"].ToString());

                    if (schildStudent == null)
                    {
                        // Wenn der Schüler in Schüler nicht existiert, wird er in die Liste der gelöschten Schüler aufgenommen.
                        gelöschteSchüler.Add((IDictionary<string, object>)rec);
                        continue;
                    }

                    if(schildStudent.Nachname == "Schmitz" && schildStudent.Vorname == "Leon Noel")
                    {
                        var debug = 1;
                    }

                    var id = schildStudent.Id;
                    schildStudent.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);
                    var schildStudentMeldung = (schildStudent.Nachname + ", " + schildStudent.Vorname + ", " + id + " (" + schildStudent.Klasse + ")").PadRight(45);

                    // Wenn er aktiv oder Gast ist, wird seine Klassenzugehörigkeit gecheckt.
                    if (new List<string>() { "2", "6" }.Contains(schildStudent.Status))
                    {
                        if (webuntisStudent["klasse.name"].ToString() != schildStudent.Klasse)
                        {
                            susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      " + webuntisStudent["klasse.name"].ToString() + " -> " + schildStudent.Klasse);

                            table.AddRow(new Text[]{
                                new Text(i+".").RightJustified(),
                                new Text(schildStudent.Nachname).LeftJustified(),
                                new Text(schildStudent.Vorname).LeftJustified(),
                                new Text(id).LeftJustified(),
                                new Text(schildStudent.Klasse).LeftJustified(),
                                new Text(schildStudent.Status).LeftJustified(),
                                new Text(webuntisStudent["klasse.name"].ToString() + " -> " + schildStudent.Klasse).LeftJustified()});

                            i++;
                        }
                    }

                    // Wenn der SchildStudent nicht aktiv (2) ist und auch kein Gast (Externer) (6) ist ...
                    if (!new List<string>() { "2", "6" }.Contains(schildStudent.Status))
                    {
                        // Prüfen, ob ein Austrittsdatum vorhanden ist und ob es in der Vergangenheit liegt
                        string exitDateString = webuntisStudent["exitDate"]?.ToString() ?? string.Empty;

                        if (exitDateString != null && !string.IsNullOrEmpty(exitDateString))
                        {
                            DateTime exitDate;
                            bool isValidDate = DateTime.TryParseExact(
                            exitDateString,
                            "dd.MM.yyyy",  // Das erwartete Datumsformat (bspw. "31.07.2025")
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out exitDate);

                            if (isValidDate && exitDate >= DateTime.Now)
                            {
                                schildStudent.GetEntlassdatum(schuelerZusatzdaten);

                                if (schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse != null)
                                {
                                    //DateTime entl;
                                    //bool isValid = DateTime.TryParseExact(schildStudent.Entlassdatum, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out entl);

                                    if (schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse >= DateTime.Now)
                                    {
                                        susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      Austritt: " + schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString());

                                        // Das Entlassdatum wird einen Tag nach der Zeugnisausgabe bzw. einen Tag nach heute gesetzt. 
                                        var entlassdatum = DateTime.Now.Date;
                                        // Der spätere der Termine gewinnt.
                                        if (entlassdatum > schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse)
                                        {
                                            entlassdatum = entlassdatum.AddDays(1);
                                        }
                                        else
                                        {
                                            entlassdatum = schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.AddDays(1);
                                        }   

                                        schildStudent.Entlassdatum = entlassdatum.ToShortDateString();

                                        table.AddRow(new Text[]{
                                            new Text(i+".").RightJustified(),
                                            new Text(schildStudent.Nachname).LeftJustified(),
                                            new Text(schildStudent.Vorname).LeftJustified(),
                                            new Text(id).LeftJustified(),
                                            new Text(schildStudent.Klasse).LeftJustified(),
                                            new Text(schildStudent.Status).LeftJustified(),
                                            new Text("Austritt: " + entlassdatum.ToShortDateString())});


                                    }
                                    else
                                    {
                                        susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      Austritt: " + schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToString("dd.MM.yyyy"));

                                        table.AddRow(new Text[]{
                                            new Text(i+".").RightJustified(),
                                            new Text(schildStudent.Nachname).LeftJustified(),
                                            new Text(schildStudent.Vorname).LeftJustified(),
                                            new Text(id).LeftJustified(),
                                            new Text(schildStudent.Klasse).LeftJustified(),
                                            new Text(schildStudent.Status).LeftJustified(),
                                            new Text("Austritt: " + schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToString("dd.MM.yyyy"))});

                                        schildStudent.Entlassdatum = DateTime.Now.ToShortDateString();
                                    }
                                }
                                i++;
                                Thread.Sleep(10);
                            }
                        }
                    }
                }

                Global.ZeileSchreiben("Webuntis-Schüler*innen", webuntisStudents.Count().ToString());
            });

            //var zieldateien = new Dateien();

            foreach (var zieldatei in zieldateien)
            {
                var zieldateiname = Path.Combine(configuration["pfadDownloads"], zieldatei.AbsoluterPfad);
                zieldatei.AbsoluterPfad = zieldateiname;

                // Alle hart in Schild gelöschten Schüler, die aber noch in WebUntis vorhanden sind,     bekommen ein Entlassdatum, das auf heute gesetzt ist.
                if (Path.GetFileName(zieldateiname).ToLower().Contains("stammdaten-schueler"))
                {
                    foreach (var dict in gelöschteSchüler)
                    {
                        dynamic record = new ExpandoObject();
                        record.Schlüssel = dict["externKey"].ToString().Split('@')[0];
                        record.EMINUSMail = dict["address.email"].ToString();
                        record.Familienname = dict["longName"].ToString();
                        record.Vorname = dict["foreName"].ToString();
                        record.Klasse = dict["klasse.name"].ToString();
                        record.Kurzname = dict["name"].ToString().Split('@')[0];
                        record.Geschlecht = dict["gender"].ToString();
                        record.Geburtsdatum = dict["birthDate"].ToString();
                        record.Eintrittsdatum = "";
                        record.Austrittsdatum = DateTime.Now.ToShortDateString();
                        record.Telefon = dict["address.phone"].ToString();
                        record.Mobil = "";
                        record.Strasse = dict["address.street"].ToString();
                        record.PLZ = dict["address.postCode"].ToString();
                        record.Ort = dict["address.city"].ToString();
                        record.ErzName = "";
                        record.ErzMobil = "";
                        record.ErzTelefon = "";
                        record.Volljährig = VolljährigJaNein(dict["birthDate"].ToString());
                        record.BetriebName = "";
                        record.BetriebStrasse = "";
                        record.BetriebPlz = "";
                        record.BetriebOrt = "";
                        record.BetriebTelefon = "";
                        record.BetriebTelefon2 = "";
                        record.BetriebMail = "";
                        record.BetriebBetreuer = "";
                        record.SchildAdressId = "";
                        record.O365Identität = "";
                        record.Benutzername = "";
                        
                        if (dict["longName"].ToString() != "Familienname")
                        {
                            zieldatei.Add(record);

                            susMitÄnderung.Add((i + ". ").PadRight(5) + "Gelöscht in SchILD" + " " + "gelöscht" + "      Austritt: " + DateTime.Now.ToShortDateString());

                            i++;
                            table.AddRow(new Text[]{
                            new Text(i+".").RightJustified(),
                            new Text(dict["longName"].ToString()).LeftJustified(),
                            new Text(dict["foreName"].ToString()).LeftJustified(),
                            new Text(dict["name"].ToString().Split('@')[0]).LeftJustified(),
                            new Text(dict["klasse.name"].ToString()).LeftJustified(),
                            new Text("gelöscht").LeftJustified(),
                            new Text("Austritt: " + DateTime.Now.ToShortDateString())});
                        }
                    }
                }

                // Ab hier die Neuanlagen

                
                var aaaaa = uniqueStudents.Where(s => s.Nachname == "Sadiku" && s.Vorname == "Almir" && s.Status == "2").ToList();


                AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .Start(zieldateiname + " vorbereiten ...", ctx =>
                {
                    foreach (var studen in uniqueStudents)
                    {
                        if(studen.Nachname == "Sadiku" && studen.Vorname == "Almir")
                        {
                            string aaaa = "";
                        }

                        // Ein Abgänger oder Absolvent, der als Externer oder Gast wiederkommt, hat in seiner aktiven Rolle den niedrigeren Status.
                        // Also wird bei mehrfach vorkommenden Schülern immer der mit dem niedrigsten Status genommen.                 
                        var student = Students.OrderBy(x => int.TryParse(x.Status, out var status) ? status : 0).FirstOrDefault(x => x.Nachname == studen.Nachname && x.Vorname == studen.Vorname && x.Geburtsdatum == studen.Geburtsdatum && x.Klasse == studen.Klasse);

                        if (student == null) continue;

                        // Wenn der Schüler in Webuntis nicht existiert, ...
                        if (!webuntisStudents.Any(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["longName"].ToString() == student.Nachname && dict["foreName"].ToString() == student.Vorname && dict["birthDate"].ToString() == student.Geburtsdatum;
                            }))
                        {
                            // ... und der Schüler in Schild aktiv der Gast ist, wird er angelegt
                            if (student.Status is "2" or "6")
                            {
                                var id = student.Id;
                                susMitÄnderung.Add(((i + ". ").PadRight(5) + student.Nachname + ", " + student.Vorname + ", " + id + " (" + student.Klasse + ")").PadRight(51) + student.Status + "      Neu: " + student.Klasse);

                                table.AddRow(new Text[]{
                                                    new Text(i+".").RightJustified(),
                                                    new Text(student.Nachname).LeftJustified(),
                                        new Text(student.Vorname).LeftJustified(),
                                        new Text(id).LeftJustified(),
                                        new Text(student.Klasse).LeftJustified(),
                                                    new Text(student.Status).LeftJustified(),
                                                    new Text("Neu in: " + student.Klasse)});
                                i++;
                            }
                        }

                        var sz = schuelerZusatzdaten
                            .Where(rec =>
                            {
                                if (rec == null) return false;
                                var dict = (IDictionary<string, object>)rec;
                                return dict != null && dict["Nachname"] != null && dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                            }).LastOrDefault() as IDictionary<string, object>;

                        var se = schuelerErzieher
                            .Where(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                            }).LastOrDefault() as IDictionary<string, object>;

                        var sa = schuelerAdressen
                            .Where(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                    dict["Adressart"].ToString() == "Betrieb";
                            }).LastOrDefault() as IDictionary<string, object>;

                        var klasse = klassen
                        .Where(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternBez"].ToString() == student.Klasse;
                        }).LastOrDefault() as IDictionary<string, object>;

                        var klassenleitung = "";

                        if (klasse != null && klasse.ContainsKey("Klassenlehrer"))
                        {
                            var dictklassenleitung = lehrkraefte.Where(rec =>
                                {
                                    var dict = (IDictionary<string, object>)rec;
                                    return dict["InternKrz"].ToString() == klasse["Klassenlehrer"].ToString();
                                }).LastOrDefault() as IDictionary<string, object>;

                            if(dictklassenleitung == null)
                            {
                                throw new Exception("Kein Klassenleitung in SchILD in Klasse " + student.Klasse + ". Bitte zuerst korrigieren und dann erneut aus SchILD exportieren.");
                            }
                            else
                            {
                                klassenleitung = dictklassenleitung["Vorname"] + " " + dictklassenleitung["Nachname"];
                            }
                        }

                        int alter = -1;
                        if (DateTime.TryParseExact(student.Geburtsdatum, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime geburtsdatum))
                        {
                            alter = DateTime.Now.Year - geburtsdatum.Year;

                            // Falls der Geburtstag dieses Jahr noch nicht war, Alter um 1 verringern
                            if (DateTime.Now < geburtsdatum.AddYears(alter))
                            {
                                alter--;
                            }
                        }

                        dynamic record = new ExpandoObject();

                        if (Path.GetFileName(zieldateiname).ToLower().Contains("stammdaten-schueler"))
                        {
                            if (student.Nachname == "Sadiku" && student.Vorname == "Almir")
                            {
                                string aa = "";
                            }

                            if (configuration["Schulnummer"] == "177659")
                            {
                                record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                ? sz["Externe ID-Nr"].ToString()
                                : sz["schulische E-Mail"].ToString().Split('@')[0];
                            }
                            else
                            {
                                record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                            }

                            record.EMINUSMail = sz["schulische E-Mail"].ToString();
                            record.Familienname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Klasse = student.Klasse;
                            record.Kurzname = sz["schulische E-Mail"].ToString().Split('@')[0];
                            record.Geschlecht = student.Geschlecht?.ToString()?.ToUpper() ?? string.Empty;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Eintrittsdatum = "";
                            record.Austrittsdatum = student.Status == "2" || student.Status == "6" ? "31.07." + Global.AktSj[1] : student.ZeugnisdatumLetztesZeugnisInDieserKlasse != null ? student.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString() : sz?["Entlassdatum"].ToString();
                            record.Telefon = sz?["Telefon-Nr."].ToString();
                            record.Mobil = "";
                            record.Strasse = student.Straße.ToString();
                            record.PLZ = student.Postleitzahl.ToString();
                            record.Ort = student.Ort.ToString();
                            record.ErzName = alter >= 18 ? "" : se?["Vorname 1.Person"].ToString() + " " + se?["Nachname 1.Person"].ToString();
                            record.ErzMobil = alter >= 18 ? "" : "";
                            record.ErzTelefon = alter >= 18 ? "" : "";
                            record.Volljährig = alter >= 18 ? "1" : "0";
                            record.BetriebName = sa == null ? "" : sa["Name1"].ToString();
                            record.BetriebStrasse = sa == null ? "" : sa["Straße"].ToString();
                            record.BetriebPlz = sa == null ? "" : sa["PLZ"].ToString();
                            record.BetriebOrt = sa == null ? "" : sa["Ort"].ToString();
                            record.BetriebTelefon = sa == null ? "" : sa["1. Tel.-Nr."].ToString();
                            record.BetriebTelefon2 = sa == null ? "" : sa["2. Tel.-Nr."].ToString();
                            record.BetriebMail = sa == null ? "" : sa["E-Mail"].ToString();
                            record.BetriebBetreuer = sa == null ? "" : (sa["Betreuer Anrede"] == null || sa["Betreuer Anrede"].ToString() == "" ? "" : sa["Betreuer Anrede"].ToString() + " ") + (sa["Betreuer Vorname"] == null || sa["Betreuer Vorname"].ToString() == "" ? "" : sa["Betreuer Vorname"].ToString() + " ") + (sa["Betreuer Nachname"] == null ? "" : sa["Betreuer Nachname"].ToString());
                            record.SchildAdressId = sa == null ? "" : sa["SchILD-Adress-ID"].ToString();
                            record.O365Identität = student.MailSchulisch;
                            record.Benutzername = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");

                            // Es werden nur diejenigen Schüler exportiert die aktiv oder Gast sind, 
                            // und alle anderen, deren Entlassdatum heute oder in den letzten sechs Wochen war.                                         
                            if (student.Status == "2" || student.Status == "6" || student.NochKeineAnzahlWochenHer(6))
                            {
                                zieldatei.Add(record);
                            }
                        }
                        else if (Path.GetFileName(zieldateiname).ToLower().Contains("erzieh"))
                        {
                            foreach (var e in new List<string>() { "Vater", "Mutter", "(sonst.) gesetzl. Vertreter", "Wohngruppe", "Vormund" })
                            {
                                var see = schuelerErzieher
                                    .Where(rec =>
                                    {
                                        var dict = (IDictionary<string, object>)rec;
                                        return dict["Nachname"].ToString() == student.Nachname &&
                                            dict["Vorname"].ToString() == student.Vorname &&
                                            dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                            dict["Erzieherart"].ToString() == e;
                                    }).LastOrDefault() as IDictionary<string, object>;

                                if (see == null) continue;

                                var stt = schuelerTelefonnummern
                                    .Where(rec =>
                                    {
                                        var dict = (IDictionary<string, object>)rec;
                                        return dict["Nachname"].ToString() == student.Nachname &&
                                            dict["Vorname"].ToString() == student.Vorname &&
                                            dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                            dict["Art"].ToString() == e;
                                    })
                                    .Cast<IDictionary<string, object>>()
                                    .ToList();

                                for (int i = 1; i <= 2; i++)
                                {
                                    record = new ExpandoObject();

                                    if (configuration["Schulnummer"] == "177659")
                                    {
                                        record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                        ? sz["Externe ID-Nr"].ToString()
                                        : sz["schulische E-Mail"].ToString().Split('@')[0];
                                    }
                                    else
                                    {
                                        record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                                    }

                                    record.EMINUSMail = sz["schulische E-Mail"].ToString();
                                    record.Familienname = student.Nachname;

                                    if (student.Nachname == "Sadiku")
                                    {
                                        string aa = "";
                                    }

                                    record.Vorname = student.Vorname;
                                    record.Klasse = student.Klasse;
                                    // Kürze e auf die letzten 18 Zeichen. Lösche Leerzeichen und Punkte.
                                    // So wird aus "Vormund" -> "Vormund", aus "(sonst.) gesetzl. Vertreter" -> "gesetzlVertreter"
                                    var ee = "";
                                    if (e.Length > 18) ee = e.Substring(e.Length - 18);
                                    ee = e.Replace(" ", "").Replace(".", "").Replace("(", "").Replace(")", "").Replace("-", "");
                                    //record.Kurzname = sz["schulische E-Mail"].ToString().Split('@')[0] + "-" + ee + "-" + i;
                                    record.Geschlecht = student.Geschlecht?.ToString()?.ToUpper() ?? string.Empty;
                                    record.Geburtsdatum = student.Geburtsdatum;
                                    record.Eintrittsdatum = "";
                                    record.Austrittsdatum = student.Status == "2" || student.Status == "6" ? "31.07." + Global.AktSj[1] : student.ZeugnisdatumLetztesZeugnisInDieserKlasse != null ? student.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString() : sz?["Entlassdatum"].ToString();
                                    record.Telefon = sz?["Telefon-Nr."].ToString();
                                    record.ErzMail = see?["E-Mail " + i + ". Person"].ToString();
                                    record.ErzStrasse = see?["Straße"].ToString();
                                    record.ErzPLZ = see?["PLZ"].ToString();
                                    record.ErzOrt = see?["Ort"].ToString();

                                    var erzName = (see?["Vorname " + i + ".Person"].ToString() + " " + see?["Nachname " + i + ".Person"].ToString()).Trim();
                                    if (string.IsNullOrEmpty(erzName)) continue;
                                    record.ErzName = e + ": " + erzName;
                                    record.ErzNachname = see?["Nachname " + i + ".Person"].ToString();
                                    record.ErzVorname = see?["Vorname " + i + ".Person"].ToString();
                                    record.ErzAnrede = see?["Anrede " + i + ".Person"].ToString();
                                    record.ErzTitel = see?["Titel " + i + ".Person"].ToString();
                                    record.ErzArt = e;
                                    record.ErzMobil = stt != null && stt.Count > 0 ? stt[0]["Telefonnr."].ToString() : "";
                                    record.ErzTelefon = stt != null && stt.Count > 1 ? stt[1]["Telefonnr."].ToString() : "";
                                    record.Volljährig = alter >= 18 ? "1" : "0";
                                    record.O365Identität = student.MailSchulisch;
                                    record.Benutzername = student.MailSchulisch.Replace(configuration["MailDomain"], "");

                                    // Der Kurzname setzt sich aus den ersten vier Buchstaben des Nachnamens, den ersten vier Buchstaben des Vornamens,
                                    // den ersten vier Buchstaben der Straße und ee zusammen.
                                    // Wenn ein Wert weniger als 4 Buchstaben hat, wird der gesamte Wert genommen.
                                    record.Kurzname = (record.ErzNachname.Length >= 4 ? record.ErzNachname.Substring(0, 4) : record.ErzNachname) +
                                        (record.ErzVorname.Length >= 4 ? record.ErzVorname.Substring(0, 4) : record.ErzVorname) +
                                        (record.ErzStrasse.Length >= 4 ? record.ErzStrasse.Substring(0, 4) : record.ErzStrasse) + ee;

                                    // Es werden nur diejenigen Schüler exportiert die aktiv oder Gast sind, 
                                    // und alle anderen, deren Entlassdatum heute oder in den letzten sechs Wochen war.                                         
                                    if (student.Status == "2" || student.Status == "6" || student.NochKeineAnzahlWochenHer(6))
                                    {
                                        zieldatei.Add(record);
                                    }
                                }
                            }
                        }
                        else if (Path.GetFileName(zieldateiname).ToLower().Contains("ausbildungs"))
                        {
                            if (configuration["Schulnummer"] == "177659")
                            {
                                record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                ? sz["Externe ID-Nr"].ToString()
                                : sz["schulische E-Mail"].ToString().Split('@')[0];
                            }
                            else
                            {
                                record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                            }

                            record.EMINUSMail = sz["schulische E-Mail"].ToString();
                            record.Familienname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Klasse = student.Klasse;
                            record.Kurzname = sz["schulische E-Mail"].ToString().Split('@')[0];
                            record.Geschlecht = student.Geschlecht?.ToString()?.ToUpper() ?? string.Empty;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Eintrittsdatum = "";
                            record.Austrittsdatum = student.Status == "2" || student.Status == "6" ? "31.07." + Global.AktSj[1] : student.ZeugnisdatumLetztesZeugnisInDieserKlasse != null ? student.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString() : sz?["Entlassdatum"].ToString();
                            record.Telefon = sz?["Telefon-Nr."].ToString();
                            record.Mobil = "";
                            record.Strasse = student.Straße.ToString();
                            record.PLZ = student.Postleitzahl.ToString();
                            record.Ort = student.Ort.ToString();
                            record.ErzName = alter >= 18 ? "" : se?["Vorname 1.Person"].ToString() + " " + se?["Nachname 1.Person"].ToString();
                            record.ErzMobil = alter >= 18 ? "" : "";
                            record.ErzTelefon = alter >= 18 ? "" : "";
                            record.Volljährig = alter >= 18 ? "1" : "0";
                            record.BetriebName = sa == null ? "" : sa["Name1"].ToString() + ", " + (sa["Straße"] == null ? "" : sa["Straße"].ToString() + ", ") + (sa["PLZ"] == null ? "" : sa["PLZ"].ToString() + " ") + (sa["Ort"] == null ? "" : sa["Ort"].ToString());
                            record.BetriebStrasse = sa == null ? "" : sa["Straße"].ToString();
                            record.BetriebPlz = sa == null ? "" : sa["PLZ"].ToString();
                            record.BetriebOrt = sa == null ? "" : sa["Ort"].ToString();
                            record.BetriebTelefon = sa == null ? "" : sa["1. Tel.-Nr."].ToString();
                            record.BetriebTelefon2 = sa == null ? "" : sa["2. Tel.-Nr."].ToString();
                            record.BetriebMail = sa == null ? "" : sa["E-Mail"].ToString();
                            record.BetriebBetreuer = sa == null ? "" : (sa["Betreuer Anrede"] == null || sa["Betreuer Anrede"].ToString() == "" ? "" : sa["Betreuer Anrede"].ToString() + " ") + (sa["Betreuer Vorname"] == null || sa["Betreuer Vorname"].ToString() == "" ? "" : sa["Betreuer Vorname"].ToString() + " ") + (sa["Betreuer Nachname"] == null ? "" : sa["Betreuer Nachname"].ToString());
                            record.SchildAdressId = sa == null ? "" : sa["SchILD-Adress-ID"].ToString();
                            record.O365Identität = student.MailSchulisch;
                            record.Benutzername = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");

                            // Es werden nur diejenigen Schüler exportiert die aktiv oder Gast sind, 
                            // und alle anderen, deren Entlassdatum heute oder in den letzten sechs Wochen war.                                         
                            if (student.Status == "2" || student.Status == "6" || student.NochKeineAnzahlWochenHer(6))
                            {
                                // Nur wenn es einen Betrieb gibt
                                if (!string.IsNullOrEmpty(record.BetriebName))
                                    zieldatei.Add(record);
                            }
                        }
                        else if (Path.GetFileName(zieldateiname).ToLower().Contains("netman"))
                        {
                            // Netman
                            // ed123456	Dagobert	Eggemann	ed123456@students.berufskolleg-borken.de	E01.07.1992	BZ22A	Stappert, Markus

                            if (configuration["Schulnummer"] == "177659")
                            {
                                record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                ? sz["Externe ID-Nr"].ToString()
                                : sz["schulische E-Mail"].ToString().Split('@')[0];
                            }
                            else
                            {
                                record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                            }

                            record.Kurzname = sz["schulische E-Mail"].ToString().Replace("@students.berufskolleg-borken.de", "");
                            record.Vorname = student.Vorname;
                            record.Nachname = student.Nachname;
                            record.Mail = sz["schulische E-Mail"].ToString();
                            record.Passwort = student.Nachname.Substring(0, 1).ToUpper() + student.Geburtsdatum;
                            record.Klasse = student.Klasse;
                            record.Klassenleitung = klassenleitung;
                            record.BetriebName = sa == null ? "" : sa["Name1"].ToString();
                            record.BetriebStrasse = sa == null ? "" : sa["Straße"].ToString();
                            record.BetriebPlz = sa == null ? "" : sa["PLZ"].ToString();
                            record.BetriebOrt = sa == null ? "" : sa["Ort"].ToString();
                            record.BetriebTelefon = sa == null ? "" : sa["1. Tel.-Nr."].ToString();

                            student.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);

                            // Aktive SuS oder Schüler mit Abschluss/Abgang, deren letztes Zeugnis noch keine 42 Tage zurückliegt.
                            if (new List<string>() { "2", "6" }.Contains(student.Status) || (new List<string>() { "8", "9" }.Contains(student.Status) && student.ZeugnisdatumLetztesZeugnisInDieserKlasse.AddDays(42) >= DateTime.Now))
                            {
                                zieldatei.Add(record);
                            }
                        }
                        else if (Path.GetFileNameWithoutExtension(zieldateiname).ToLower().Contains("littera"))
                        {
                            // Littera
                            record.LGruppe = student.Klasse;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Titel = "";
                            record.Nachname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Strasse = alter >= 18 ? student.Straße.ToString() : se?["Straße"].ToString();
                            record.PLZ = alter >= 18 ? student.Postleitzahl.ToString() : se?["PLZ"].ToString();
                            record.Ort = alter >= 18 ? student.Ort.ToString() : se?["Ort"].ToString();
                            record.Geschlecht = student.Geschlecht.ToString();
                            record.Anmeldedatum = student.BeginnDesBildungsganges;
                            record.Telefon = sz?["Telefon-Nr."].ToString();
                            record.Mobiltelefon = sz?["Fax/Mobilnr"].ToString();
                            record.email = sz["schulische E-Mail"].ToString();
                            record.ZusatzInfo = "";
                            record.Bemerkung = "";
                            record.Geschlecht = student.Geschlecht.ToString().ToUpper();
                            student.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);

                            // Aktive SuS oder Schüler mit Abschluss/Abgang, deren letztes Zeugnis noch keine 42 Tage zurückliegt.
                            if (new List<string>() { "2", "6" }.Contains(student.Status) || (new List<string>() { "8", "9" }.Contains(student.Status) && student.ZeugnisdatumLetztesZeugnisInDieserKlasse.AddDays(42) >= DateTime.Now))
                            {
                                zieldatei.Add(record);
                            }
                        }
                    }

                    //Global.ZeileSchreiben(zieldateiname, string.Join(Environment.NewLine, zieldatei.Importhinweise));

                });

                foreach (var aktion in zieldatei.Funktionen)
                    aktion(zieldatei);
                                                
                if (zieldateiname.ToLower().Contains("schueler") && susMitÄnderung.Count() > 2)
                    AnsiConsole.Write(table);
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
            while (Console.KeyAvailable) Console.ReadKey(true);

            Console.ReadKey();
        }
    }

    private string VolljährigJaNein(string? geburtsdatum)
    {
        // Gibt "Ja" zurück, wenn das Geburtsdatum ein volljähriges Alter (>= 18) ergibt, sonst "Nein".
        if (string.IsNullOrWhiteSpace(geburtsdatum))
            return "Nein";

        DateTime parsed;
        var formats = new[]
        {
            "dd.MM.yyyy",
            "d.M.yyyy",
            "dd.MM.yy",
            "yyyy-MM-dd",
            "yyyyMMdd",
            "dd.MM.yyyy HH:mm",
            "d.M.yyyy HH:mm"
        };

        var input = geburtsdatum.Trim();

        // Versuche zuerst mehrere feste Formate, dann deutsch-kulturelle und zuletzt allgemeines TryParse.
        if (!DateTime.TryParseExact(input, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)
            && !DateTime.TryParse(input, new CultureInfo("de-DE"), DateTimeStyles.None, out parsed)
            && !DateTime.TryParse(input, out parsed))
        {
            return "Nein";
        }

        var today = DateTime.Today;
        var age = today.Year - parsed.Year;
        if (parsed > today.AddYears(-age)) age--;

        return age >= 18 ? "Ja" : "Nein";
    }

    public Dateien WebuntisUndCo(IConfiguration configuration, List<Datei> zieldateien)
    {
        var zielDateien = new Dateien();
        var zeitstempel = DateTime.Now.ToString("yyyyMMdd-HHmm");

        try
        {
            List<dynamic>? webuntisStudents = Quelldateien.GetMatchingList(configuration, "student_", Students, Klassen);
            if (webuntisStudents == null || webuntisStudents.Count == 0) return new Dateien();
            var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", Students, Klassen);
            if (schuelerZusatzdaten == null || schuelerZusatzdaten.Count == 0) return new Dateien();
            var schuelerLernabschnittsdaten = Quelldateien.GetMatchingList(configuration, "schuelerlernabschnittsdaten", Students, Klassen);
            if (schuelerLernabschnittsdaten == null || schuelerLernabschnittsdaten.Count == 0) return new Dateien();
            var schuelerErzieher = Quelldateien.GetMatchingList(configuration, "schuelererzieher", Students, Klassen);
            if (schuelerErzieher == null || schuelerErzieher.Count == 0) return new Dateien();
            var schuelerAdressen = Quelldateien.GetMatchingList(configuration, "schueleradressen", Students, Klassen);
            if (schuelerAdressen == null || schuelerAdressen.Count == 0) return new Dateien();
            var lehrkraefte = Quelldateien.GetMatchingList(configuration, "lehrkraefte", Students, Klassen);
            if (lehrkraefte == null || lehrkraefte.Count == 0) return new Dateien();
            var klassen = Quelldateien.GetMatchingList(configuration, "klassen", Students, Klassen);
            if (klassen == null || klassen.Count == 0) return new Dateien();


            var susMitÄnderung = new List<string>()
            {
                "Folgende Änderungen / Neuanlagen:",
                "Nr".PadRight(5) + "Name".PadRight(46) + "Status Änderung".PadRight(20)
            };

            var i = 1;

            var table = new Spectre.Console.Table();
            table.Border(TableBorder.Rounded);
            table.Centered();
            table.Expand();

            // Add columns 
            table.AddColumn("Nr");
            table.AddColumn("Name");
            table.AddColumn("Vorname");
            table.AddColumn("ID");
            table.AddColumn("Klasse");
            table.AddColumn("Status");
            table.AddColumn("Änderung");

            AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("Webuntis-Schüler*innen vorbereiten ...", ctx =>
            {
                foreach (var rec in webuntisStudents)
                {
                    //task.Increment(1);
                    if (rec is not IDictionary<string, object> webuntisStudent) continue;

                    // Ein Abgänger oder Absolvent, der als Externer oder Gast wiederkommt, hat in seiner aktiven Rolle den niedrigeren Status.
                    // Also wird bei mehrfach vorkommenden Schülern immer der mit dem niedrigsten Status genommen.                 
                    var schildStudent = Students
                        .OrderBy(x => int.TryParse(x.Status, out var status) ? status : 0)
                        .FirstOrDefault(x =>
                            x.Nachname == webuntisStudent["longName"].ToString() &&
                            x.Vorname == webuntisStudent["foreName"].ToString() &&
                            x.Geburtsdatum == webuntisStudent["birthDate"].ToString());

                    schildStudent.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);


                    var id = schildStudent.Id;

                    if (schildStudent == null) continue;

                    var schildStudentMeldung = (schildStudent.Nachname + ", " + schildStudent.Vorname + ", " + id + " (" + schildStudent.Klasse + ")").PadRight(45);

                    // Wenn er aktiv oder Gast ist, wird seine Klassenzugehörigkeit gecheckt.
                    if (new List<string>() { "2", "6" }.Contains(schildStudent.Status))
                    {
                        if (webuntisStudent["klasse.name"].ToString() != schildStudent.Klasse)
                        {
                            susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      " + webuntisStudent["klasse.name"].ToString() + " -> " + schildStudent.Klasse);

                            table.AddRow(new Text[]{
                            new Text(i+".").RightJustified(),
                            new Text(schildStudent.Nachname).LeftJustified(),
                            new Text(schildStudent.Vorname).LeftJustified(),
                            new Text(id).LeftJustified(),
                            new Text(schildStudent.Klasse).LeftJustified(),
                            new Text(schildStudent.Status).LeftJustified(),
                            new Text(webuntisStudent["klasse.name"].ToString() + " -> " + schildStudent.Klasse).LeftJustified()});

                            i++;
                        }
                    }

                    // Wenn der SchildStudent nicht aktiv (2) ist und auch kein Gast (Externer) (6) ist ...
                    if (!new List<string>() { "2", "6" }.Contains(schildStudent.Status))
                    {
                        // Prüfen, ob ein Austrittsdatum vorhanden ist und ob es in der Vergangenheit liegt
                        string exitDateString = webuntisStudent["exitDate"]?.ToString() ?? string.Empty;

                        if (exitDateString != null && !string.IsNullOrEmpty(exitDateString))
                        {
                            DateTime exitDate;
                            bool isValidDate = DateTime.TryParseExact(
                            exitDateString,
                            "dd.MM.yyyy",  // Das erwartete Datumsformat (bspw. "31.07.2025")
                            CultureInfo.InvariantCulture,
                            DateTimeStyles.None,
                            out exitDate);

                            if (isValidDate && exitDate >= DateTime.Now)
                            {
                                schildStudent.GetEntlassdatum(schuelerZusatzdaten);

                                if (schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse != null)
                                {
                                    //DateTime entl;
                                    //bool isValid = DateTime.TryParseExact(schildStudent.Entlassdatum, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out entl);

                                    if (schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse >= DateTime.Now)
                                    {
                                        susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      Austritt: " + schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString());

                                        table.AddRow(new Text[]{
                                        new Text(i+".").RightJustified(),
                                        new Text(schildStudent.Nachname).LeftJustified(),
                                        new Text(schildStudent.Vorname).LeftJustified(),
                                        new Text(id).LeftJustified(),
                                        new Text(schildStudent.Klasse).LeftJustified(),
                                        new Text(schildStudent.Status).LeftJustified(),
                                        new Text("Austritt: " + schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString())});

                                        schildStudent.Entlassdatum = DateTime.Now.ToShortDateString();
                                    }
                                    else
                                    {
                                        susMitÄnderung.Add((i + ". ").PadRight(5) + schildStudentMeldung + " " + schildStudent.Status + "      Austritt: " + schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToString("dd.MM.yyyy"));

                                        table.AddRow(new Text[]{
                                        new Text(i+".").RightJustified(),
                                        new Text(schildStudent.Nachname).LeftJustified(),
                                        new Text(schildStudent.Vorname).LeftJustified(),
                                        new Text(id).LeftJustified(),
                                        new Text(schildStudent.Klasse).LeftJustified(),
                                        new Text(schildStudent.Status).LeftJustified(),
                                        new Text("Austritt: " + schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToString("dd.MM.yyyy"))});

                                        schildStudent.Entlassdatum = DateTime.Now.ToShortDateString();
                                    }
                                }
                                i++;
                                Thread.Sleep(10);
                            }
                        }
                    }
                }

                Global.ZeileSchreiben("Webuntis-Schüler*innen", webuntisStudents.Count().ToString());
            });

            // Ab hier die Neuanlagen

            var uniqueStudents = Students
                .DistinctBy(s => new { s.Vorname, s.Nachname, s.Geburtsdatum })
                .OrderBy(s => s.Klasse)
                .ThenBy(s => s.Nachname)
                .ThenBy(s => s.Vorname);

            foreach (var zieldatei in zieldateien)
            {
                var datei = new Datei();
                datei.AbsoluterPfad = zieldatei.AbsoluterPfad;
                datei.Delimiter = zieldatei.Delimiter;
                datei.Quote = zieldatei.Quote;
                datei.Encoding = zieldatei.Encoding;

                table = new Spectre.Console.Table();
                table.Title = new TableTitle($"Änderungen in [{Global.GetColor(Global.ColorPfadInDateien)}]{zeitstempel + Path.GetFileName(zieldatei.AbsoluterPfad)}[/]");
                table.Border(TableBorder.Rounded);
                table.Centered();
                table.Expand();

                // Add columns 
                table.AddColumn("Nr");
                table.AddColumn("Name");
                table.AddColumn("Vorname");
                table.AddColumn("ID");
                table.AddColumn("Klasse");
                table.AddColumn("Status");
                table.AddColumn("Änderung");

                zieldatei.AbsoluterPfad = Path.Combine(configuration["pfadDownloads"], zeitstempel + Path.GetFileName(zieldatei.AbsoluterPfad));

                AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start(Path.GetFileName(zeitstempel + zieldatei.AbsoluterPfad) + " vorbereiten ...", ctx =>
                {
                    int i = 0;

                    foreach (var studen in uniqueStudents)
                    {
                        // Ein Abgänger oder Absolvent, der als Externer oder Gast wiederkommt, hat in seiner aktiven Rolle den niedrigeren Status.
                        // Also wird bei mehrfach vorkommenden Schülern immer der mit dem niedrigsten Status genommen.                 
                        var student = Students.OrderBy(x => int.TryParse(x.Status, out var status) ? status : 0).FirstOrDefault(x => x.Nachname == studen.Nachname && x.Vorname == studen.Vorname && x.Geburtsdatum == studen.Geburtsdatum);

                        if (student == null) continue;

                        // Wenn der Schüler in Webuntis nicht existiert, ...
                        if (!webuntisStudents.Any(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["longName"].ToString() == student.Nachname && dict["foreName"].ToString() == student.Vorname && dict["birthDate"].ToString() == student.Geburtsdatum;
                            }))
                        {
                            // ... und der Schüler in Schild aktiv oder Gast ist, wird er angelegt
                            if (student.Status is "2" or "6")
                            {
                                var id = student.Id;
                                susMitÄnderung.Add(((i + ". ").PadRight(5) + student.Nachname + ", " + student.Vorname + ", " + id + " (" + student.Klasse + ")").PadRight(51) + student.Status + "      Neu: " + student.Klasse);

                                table.AddRow(new Text[]{
                                                    new Text(i+".").RightJustified(),
                                                    new Text(student.Nachname).LeftJustified(),
                                        new Text(student.Vorname).LeftJustified(),
                                        new Text(id).LeftJustified(),
                                        new Text(student.Klasse).LeftJustified(),
                                                    new Text(student.Status).LeftJustified(),
                                                    new Text("Neu in: " + student.Klasse)});
                                i++;
                            }
                        }

                        var sz = schuelerZusatzdaten
                            .Where(rec =>
                            {
                                if (rec == null) return false;
                                var dict = (IDictionary<string, object>)rec;
                                return dict != null && dict["Nachname"] != null && dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                            }).LastOrDefault() as IDictionary<string, object>;

                        var se = schuelerErzieher
                            .Where(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                            }).LastOrDefault() as IDictionary<string, object>;

                        var sa = schuelerAdressen
                            .Where(rec =>
                            {
                                var dict = (IDictionary<string, object>)rec;
                                return dict["Nachname"].ToString() == student.Nachname &&
                                    dict["Vorname"].ToString() == student.Vorname &&
                                    dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                    dict["Adressart"].ToString() == "Betrieb";
                            }).LastOrDefault() as IDictionary<string, object>;

                        var klasse = klassen
                        .Where(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternBez"].ToString() == student.Klasse;
                        }).LastOrDefault() as IDictionary<string, object>;

                        var klassenleitung = "";

                        if (klasse != null && klasse.ContainsKey("Klassenlehrer"))
                        {
                            var dictklassenleitung = lehrkraefte.Where(rec =>
                                {
                                    var dict = (IDictionary<string, object>)rec;
                                    return dict["InternKrz"].ToString() == klasse["Klassenlehrer"].ToString();
                                }).LastOrDefault() as IDictionary<string, object>;

                            klassenleitung = dictklassenleitung["Vorname"] + " " + dictklassenleitung["Nachname"];
                        }

                        int alter = -1;
                        if (DateTime.TryParseExact(student.Geburtsdatum, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime geburtsdatum))
                        {
                            alter = DateTime.Now.Year - geburtsdatum.Year;

                            // Falls der Geburtstag dieses Jahr noch nicht war, Alter um 1 verringern
                            if (DateTime.Now < geburtsdatum.AddYears(alter))
                            {
                                alter--;
                            }
                        }

                        dynamic record = new ExpandoObject();

                        if (Path.GetFileName(zieldatei.AbsoluterPfad).ToLower().Contains("schueler"))
                        {
                            if (configuration["Schulnummer"] == "177659")
                            {
                                record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                ? sz["Externe ID-Nr"].ToString()
                                : sz["schulische E-Mail"].ToString().Split('@')[0];
                            }
                            else
                            {
                                record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                            }

                            record.EMINUSMail = sz["schulische E-Mail"].ToString();
                            record.Familienname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Klasse = student.Klasse;
                            record.Kurzname = sz["schulische E-Mail"].ToString().Split('@')[0];
                            record.Geschlecht = student.Geschlecht?.ToString()?.ToUpper() ?? string.Empty;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Eintrittsdatum = "";
                            record.Austrittsdatum = student.Status == "2" || student.Status == "6" ? "31.07." + Global.AktSj[1] : student.ZeugnisdatumLetztesZeugnisInDieserKlasse != null ? student.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString() : sz?["Entlassdatum"].ToString();
                            record.Telefon = sz?["Telefon-Nr."].ToString();
                            record.Mobil = "";
                            record.Strasse = alter >= 18 ? student.Straße.ToString() : se?["Straße"].ToString();
                            record.PLZ = alter >= 18 ? student.Postleitzahl.ToString() : se?["PLZ"].ToString();
                            record.Ort = alter >= 18 ? student.Ort.ToString() : se?["Ort"].ToString();
                            record.ErzName = alter >= 18 ? "" : se?["Vorname 1.Person"].ToString() + " " + se?["Nachname 1.Person"].ToString();
                            record.ErzMobil = alter >= 18 ? "" : "";
                            record.ErzTelefon = alter >= 18 ? "" : "";
                            record.Volljährig = alter >= 18 ? "1" : "0";
                            record.BetriebName = sa == null ? "" : sa["Name1"].ToString();
                            record.BetriebStrasse = sa == null ? "" : sa["Straße"].ToString();
                            record.BetriebPlz = sa == null ? "" : sa["PLZ"].ToString();
                            record.BetriebOrt = sa == null ? "" : sa["Ort"].ToString();
                            record.BetriebTelefon = sa == null ? "" : sa["1. Tel.-Nr."].ToString();
                            record.BetriebTelefon2 = sa == null ? "" : sa["2. Tel.-Nr."].ToString();
                            record.BetriebMail = sa == null ? "" : sa["E-Mail"].ToString();
                            record.BetriebBetreuer = sa == null ? "" : (sa["Betreuer Anrede"] == null || sa["Betreuer Anrede"].ToString() == "" ? "" : sa["Betreuer Anrede"].ToString() + " ") + (sa["Betreuer Vorname"] == null || sa["Betreuer Vorname"].ToString() == "" ? "" : sa["Betreuer Vorname"].ToString() + " ") + (sa["Betreuer Nachname"] == null ? "" : sa["Betreuer Nachname"].ToString());
                            record.SchildAdressId = sa == null ? "" : sa["SchILD-Adress-ID"].ToString();
                            record.O365Identität = student.MailSchulisch;
                            record.Benutzername = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");

                            // Es werden nur diejenigen Schüler exportiert die aktiv oder Gast sind, 
                            // und alle anderen, deren Entlassdatum heute oder in den letzten sechs Wochen war.                                         
                            if (student.Status == "2" || student.Status == "6" || student.NochKeineAnzahlWochenHer(6))
                            {
                                zieldatei.Add(record);
                                i++;
                                Console.WriteLine(i.ToString());
                            }
                        }
                        else if (Path.GetFileName(zieldatei.AbsoluterPfad).ToLower().Contains("erzieher"))
                        {
                            if (configuration["Schulnummer"] == "177659")
                            {
                                record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                ? sz["Externe ID-Nr"].ToString()
                                : sz["schulische E-Mail"].ToString().Split('@')[0];
                            }
                            else
                            {
                                record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                            }

                            record.EMINUSMail = sz["schulische E-Mail"].ToString();
                            record.Familienname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Klasse = student.Klasse;
                            record.Kurzname = sz["schulische E-Mail"].ToString().Split('@')[0];
                            record.Geschlecht = student.Geschlecht?.ToString()?.ToUpper() ?? string.Empty;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Eintrittsdatum = "";
                            record.Austrittsdatum = student.Status == "2" || student.Status == "6" ? "31.07." + Global.AktSj[1] : student.ZeugnisdatumLetztesZeugnisInDieserKlasse != null ? student.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString() : sz?["Entlassdatum"].ToString();
                            record.Telefon = sz?["Telefon-Nr."].ToString();
                            record.Mobil = "";
                            record.Strasse = alter >= 18 ? student.Straße.ToString() : se?["Straße"].ToString();
                            record.PLZ = alter >= 18 ? student.Postleitzahl.ToString() : se?["PLZ"].ToString();
                            record.Ort = alter >= 18 ? student.Ort.ToString() : se?["Ort"].ToString();
                            record.ErzName = alter >= 18 ? "" : se?["Vorname 1.Person"].ToString() + " " + se?["Nachname 1.Person"].ToString();
                            record.ErzMobil = alter >= 18 ? "" : sa == null ? "" : sa["1. Tel.-Nr."].ToString();
                            record.ErzTelefon = alter >= 18 ? "" : sa == null ? "" : sa["2. Tel.-Nr."].ToString();
                            record.Volljährig = alter >= 18 ? "1" : "0";
                            record.SchildAdressId = sa == null ? "" : sa["SchILD-Adress-ID"].ToString();
                            record.O365Identität = student.MailSchulisch;
                            record.Benutzername = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");

                            // Es werden nur diejenigen Schüler exportiert die aktiv oder Gast sind, 
                            // und alle anderen, deren Entlassdatum heute oder in den letzten sechs Wochen war.                                         
                            if (student.Status == "2" || student.Status == "6" || student.NochKeineAnzahlWochenHer(6))
                            {
                                if (alter < 18)
                                    zieldatei.Add(record);
                            }
                        }
                        else if (Path.GetFileName(zieldatei.AbsoluterPfad).ToLower().Contains("betrieb"))
                        {
                            if (configuration["Schulnummer"] == "177659")
                            {
                                record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                ? sz["Externe ID-Nr"].ToString()
                                : sz["schulische E-Mail"].ToString().Split('@')[0];
                            }
                            else
                            {
                                record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                            }

                            record.EMINUSMail = sz["schulische E-Mail"].ToString();
                            record.Familienname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Klasse = student.Klasse;
                            record.Kurzname = sz["schulische E-Mail"].ToString().Split('@')[0];
                            record.Geschlecht = student.Geschlecht?.ToString()?.ToUpper() ?? string.Empty;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Eintrittsdatum = "";
                            record.Austrittsdatum = student.Status == "2" || student.Status == "6" ? "31.07." + Global.AktSj[1] : student.ZeugnisdatumLetztesZeugnisInDieserKlasse != null ? student.ZeugnisdatumLetztesZeugnisInDieserKlasse.ToShortDateString() : sz?["Entlassdatum"].ToString();
                            record.Telefon = sz?["Telefon-Nr."].ToString();
                            record.Mobil = "";
                            record.Strasse = alter >= 18 ? student.Straße.ToString() : se?["Straße"].ToString();
                            record.PLZ = alter >= 18 ? student.Postleitzahl.ToString() : se?["PLZ"].ToString();
                            record.Ort = alter >= 18 ? student.Ort.ToString() : se?["Ort"].ToString();
                            record.ErzName = alter >= 18 ? "" : se?["Vorname 1.Person"].ToString() + " " + se?["Nachname 1.Person"].ToString();
                            record.ErzMobil = alter >= 18 ? "" : "";
                            record.ErzTelefon = alter >= 18 ? "" : "";
                            record.Volljährig = alter >= 18 ? "1" : "0";
                            record.BetriebName = sa == null ? "" : sa["Name1"].ToString() + ", " + (sa["Straße"] == null ? "" : sa["Straße"].ToString() + ", ") + (sa["PLZ"] == null ? "" : sa["PLZ"].ToString() + " ") + (sa["Ort"] == null ? "" : sa["Ort"].ToString());
                            record.BetriebStrasse = sa == null ? "" : sa["Straße"].ToString();
                            record.BetriebPlz = sa == null ? "" : sa["PLZ"].ToString();
                            record.BetriebOrt = sa == null ? "" : sa["Ort"].ToString();
                            record.BetriebTelefon = sa == null ? "" : sa["1. Tel.-Nr."].ToString();
                            record.BetriebTelefon2 = sa == null ? "" : sa["2. Tel.-Nr."].ToString();
                            record.BetriebMail = sa == null ? "" : sa["E-Mail"].ToString();
                            record.BetriebBetreuer = sa == null ? "" : (sa["Betreuer Anrede"] == null || sa["Betreuer Anrede"].ToString() == "" ? "" : sa["Betreuer Anrede"].ToString() + " ") + (sa["Betreuer Vorname"] == null || sa["Betreuer Vorname"].ToString() == "" ? "" : sa["Betreuer Vorname"].ToString() + " ") + (sa["Betreuer Nachname"] == null ? "" : sa["Betreuer Nachname"].ToString());
                            record.SchildAdressId = sa == null ? "" : sa["SchILD-Adress-ID"].ToString();
                            record.O365Identität = student.MailSchulisch;
                            record.Benutzername = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");

                            // Es werden nur diejenigen Schüler exportiert die aktiv oder Gast sind, 
                            // und alle anderen, deren Entlassdatum heute oder in den letzten sechs Wochen war.                                         
                            if (student.Status == "2" || student.Status == "6" || student.NochKeineAnzahlWochenHer(6))
                            {
                                // Nur wenn es einen Betrieb gibt
                                if (!string.IsNullOrEmpty(record.BetriebName))
                                    zieldatei.Add(record);
                            }
                        }
                        else if (Path.GetFileName(zieldatei.AbsoluterPfad).ToLower().Contains("netman"))
                        {
                            // Netman
                            // ed123456	Dagobert	Eggemann	ed123456@students.berufskolleg-borken.de	E01.07.1992	BZ22A	Stappert, Markus

                            if (configuration["Schulnummer"] == "177659")
                            {
                                record.Schlüssel = !string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) && sz["Externe ID-Nr"].ToString().Length == 6 && sz["Externe ID-Nr"].ToString().StartsWith("15")
                                ? sz["Externe ID-Nr"].ToString()
                                : sz["schulische E-Mail"].ToString().Split('@')[0];
                            }
                            else
                            {
                                record.Schlüssel = sz["schulische E-Mail"].ToString().Split('@')[0];
                            }

                            if (student.Nachname == "Boulos")
                            {
                                string aa = "";
                            }

                            record.Kurzname = sz["schulische E-Mail"].ToString().Replace("@students.berufskolleg-borken.de", "");
                            record.Vorname = student.Vorname;
                            record.Nachname = student.Nachname;
                            record.Mail = sz["schulische E-Mail"].ToString();
                            record.Passwort = student.Nachname.Substring(0, 1).ToUpper() + student.Geburtsdatum;
                            record.Klasse = student.Klasse;
                            record.Klassenleitung = klassenleitung;
                            record.BetriebName = sa == null ? "" : sa["Name1"].ToString();
                            record.BetriebStrasse = sa == null ? "" : sa["Straße"].ToString();
                            record.BetriebPlz = sa == null ? "" : sa["PLZ"].ToString();
                            record.BetriebOrt = sa == null ? "" : sa["Ort"].ToString();
                            record.BetriebTelefon = sa == null ? "" : sa["1. Tel.-Nr."].ToString();

                            student.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);

                            // Aktive SuS oder Schüler mit Abschluss/Abgang, deren letztes Zeugnis noch keine 42 Tage zurückliegt.
                            if (new List<string>() { "2", "6" }.Contains(student.Status) || (new List<string>() { "8", "9" }.Contains(student.Status) && student.ZeugnisdatumLetztesZeugnisInDieserKlasse.AddDays(42) >= DateTime.Now))
                            {
                                zieldatei.Add(record);
                            }
                        }
                        else if (Path.GetFileNameWithoutExtension(zieldatei.AbsoluterPfad).ToLower().Contains("littera"))
                        {
                            // Littera
                            record.LGruppe = student.Klasse;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Titel = "";
                            record.Nachname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Strasse = alter >= 18 ? student.Straße.ToString() : se?["Straße"].ToString();
                            record.PLZ = alter >= 18 ? student.Postleitzahl.ToString() : se?["PLZ"].ToString();
                            record.Ort = alter >= 18 ? student.Ort.ToString() : se?["Ort"].ToString();
                            record.Geschlecht = student.Geschlecht.ToString();
                            record.Anmeldedatum = student.BeginnDesBildungsganges;
                            record.Telefon = sz?["Telefon-Nr."].ToString();
                            record.Mobiltelefon = sz?["Fax/Mobilnr"].ToString();
                            record.email = sz["schulische E-Mail"].ToString();
                            record.ZusatzInfo = "";
                            record.Bemerkung = "";
                            record.Geschlecht = student.Geschlecht.ToString().ToUpper();
                            student.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);

                            // Aktive SuS oder Schüler mit Abschluss/Abgang, deren letztes Zeugnis noch keine 42 Tage zurückliegt.
                            if (new List<string>() { "2", "6" }.Contains(student.Status) || (new List<string>() { "8", "9" }.Contains(student.Status) && student.ZeugnisdatumLetztesZeugnisInDieserKlasse.AddDays(42) >= DateTime.Now))
                            {
                                zieldatei.Add(record);
                            }
                        }
                        // Wenn der Pfad "webuntis" enthält, dann wird geprüft, ob es in webuntis einen Schüler gibt, den es in Schülerbasisdaten nicht gibt.
                        // Das kann passieren, wenn ein Schüler sich z.B. einmal mit Zweitem Vornamen angeleldet hat und einmal ohne.
                        if (zieldatei.AbsoluterPfad.ToLower().Contains("schueler"))
                        {
                            foreach (var rec in webuntisStudents)
                            {
                                var dict = (IDictionary<string, object>)rec;

                                if (!Students.Any(x => x.Nachname == dict["longName"].ToString() && x.Vorname == dict["foreName"].ToString() && x.Geburtsdatum == dict["birthDate"].ToString()))
                                {
                                    Console.WriteLine($"Schüler {dict["longName"]} {dict["foreName"]} ({dict["birthDate"]}) ist in Webuntis vorhanden, aber nicht in Schülerbasisdaten.");
                                }
                            }
                        }
                    }
                    zielDateien.Add(zieldatei);
                    Global.ZeileSchreiben(zeitstempel + Path.GetFileName(zieldatei.AbsoluterPfad), zieldatei.Count().ToString());
                });
                if (susMitÄnderung.Count() > 2)
                    AnsiConsole.Write(table);
            }
            return zielDateien;
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public void KlassenErstellen(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null,
        string defaultwert = "",
        Global.Modus modus = Global.Modus.Update)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        var schildKlassen = Quelldateien.GetMatchingList(configuration, "klassen", Students, Klassen);
        if (schildKlassen.Count == 0) return;
        var untisKlassen = Quelldateien.GetMatchingList(configuration, "GPU003", Students, Klassen);
        if (untisKlassen.Count == 0) return;
        List<dynamic> gpu002 = Quelldateien.GetMatchingList(configuration, "gpu002", IStudents, Klassen);
        if (gpu002 == null || gpu002.Count == 0) return;

        var records = new List<dynamic>();
        bool neueKlassen = false;
        var klassenOhneUnterricht = "";

        foreach (var untisKlasse in untisKlassen)
        {
            var dictUntis = (IDictionary<string, object>)untisKlasse;

            // Prüfe, ob es Unterricht zur Klasse gibt
            var klasseHatUnterricht = gpu002.Any(rec =>
                {
                    if (rec == null) return false;
                    var dict = (IDictionary<string, object>)rec;
                    return dict != null && dict["Field5"] != null && dict["Field5"].ToString() == dictUntis["Field1"].ToString();
                });

            // Wenn die Klasse keinen Unterricht hat, wird sie übersprungen
            if (!klasseHatUnterricht)
            {
                klassenOhneUnterricht += $"{dictUntis["Field1"]}, ";
                continue;
            }

            var klasseVonDerKopiertWird = dictUntis["Field1"].ToString();

            // Wenn es die Klasse in Schild nicht gibt
            if (!schildKlassen.Any(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["InternBez"].ToString() == dictUntis["Field1"].ToString();
                }))
            {
                klasseVonDerKopiertWird = DecreaseNumberInString(dictUntis["Field1"].ToString());
                neueKlassen = true;
            }

            // Suche die korrespondierende (Vorgänger-)klasse mit allen Schildeigenschaften
            var schildKlasseVonDerKopiertWird = schildKlassen.FirstOrDefault(zeile =>
            {
                var dict = (IDictionary<string, object>)zeile;
                return dict["InternBez"].ToString() == klasseVonDerKopiertWird;
            });

            if (schildKlasseVonDerKopiertWird != null)
            {
                var s = (IDictionary<string, object>)schildKlasseVonDerKopiertWird;
                dynamic record = new ExpandoObject();
                record.InternBez = dictUntis["Field1"].ToString();
                record.StatistikBez = dictUntis["Field1"].ToString();
                record.SonstigeBez = "";
                record.Jahrgang = dictUntis["Field14"].ToString();
                record.Folgeklasse = dictUntis["Field1"].ToString();
                record.Klassenlehrer = dictUntis.ContainsKey("Field30") && dictUntis["Field30"] != null
                    ? dictUntis["Field30"].ToString().Split(',').FirstOrDefault() ?? ""
                    : "";
                record.OrgForm = s["OrgForm"].ToString();
                record.Klassenart = s["Klassenart"].ToString();
                record.Gliederung = s["Gliederung"].ToString();
                record.Fachklasse = s["Fachklasse"].ToString();
                zieldatei.Add(record);
            }
        }

        if (neueKlassen)
        {
            // Erstelle ein panel
            var panel = new Panel("Es werden neue Klassen angelegt. Die Prüfungsordnung und die stellv. Klassenleitung müssen manuell eingetragen werden.")
                .HeaderAlignment(Justify.Left)
                .Header($"  Hinweis  ")
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Red);
        }

        if (klassenOhneUnterricht != "")
        {
            klassenOhneUnterricht = klassenOhneUnterricht.TrimEnd(' ', ',');
            var panel = new Panel($"Klassen ohne Unterricht bleiben unberücksichtigt: {klassenOhneUnterricht}.")
                .HeaderAlignment(Justify.Left)
                .Header($"  Hinweis  ")
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Red);
            AnsiConsole.Write(panel);
        }
        
        foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
    }

    private string DecreaseNumberInString(string? input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        // Regex sucht eine zweistellige Zahl (\d{2})
        Match match = Regex.Match(input, @"\d{2}");

        if (match.Success)
        {
            int number = int.Parse(match.Value) - 1; // Zahl um 1 verringern
            string newNumber = number.ToString("D2"); // Sicherstellen, dass es zweistellig bleibt

            // Ersetze die Zahl
            string result = input.Replace(match.Value, newNumber);

            // Ersetze den letzten Buchstaben durch 'A'
            result = result.Substring(0, result.Length - 1) + "A";

            return result;
        }

        return input; // Falls keine Zahl gefunden wird, bleibt der String unverändert
    }


    public void Sprechtag(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string hinweis)
    {
        var exportLessons = Quelldateien.GetMatchingList(configuration, "exportlessons", IStudents, Klassen);
        if (exportLessons == null || !exportLessons.Any()) return;

        var raums = Quelldateien.GetMatchingList(configuration, "gpu005", IStudents, Klassen);
        if (raums == null || !raums.Any()) return;

        var gpu004 = Quelldateien.GetMatchingList(configuration, "gpu004", IStudents, Klassen);
        if (gpu004 == null || !gpu004.Any()) return;

        var lehrkraefte = Quelldateien.GetMatchingList(configuration, "lehrkraefte", IStudents, Klassen);
        if (lehrkraefte == null || lehrkraefte.Count == 0) return;

        var zieldatei = new Datei(zieldateiname, funktionen, configuration);

        var lehrers = new Lehrers();
        foreach (var rec in lehrkraefte)
        {
            var dict = (IDictionary<string, object>)rec;

            var lehGpu004 = gpu004.Where(rec =>
            {
                if (rec == null) return false;
                var dic = (IDictionary<string, object>)rec;
                return dic != null && dic["Field1"] != null && dic["Field1"].ToString() == dict["InternKrz"].ToString();
            }).LastOrDefault() as IDictionary<string, object>;
            
            if(lehGpu004 == null)
            {
                continue;
            }

            var l = new Lehrer
            {
                Kürzel = dict["InternKrz"].ToString(),
                Nachname = dict["Nachname"].ToString(),
                Vorname = dict["Vorname"].ToString(),
                Anrede = dict["Anrede"].ToString(),
                Titel = dict["Titel"].ToString(),
                Raum = lehGpu004["Field5"].ToString(),
                Text2 = lehGpu004["Field43"].ToString()
            };
            lehrers.Add(l);
        }

        configuration = Global.Konfig("SprechtagsDatum", Global.Modus.Update, configuration);

        hinweis = hinweis.Replace(" nach der allgemeinen Zeugnisausgabe", ", **" + configuration["SprechtagsDatum"] + "**,");

        // Distincte und sortierte Liste aller Lehrer-Kürzel, die im Unterricht sind  
        var alleLehrerImUnterrichtKürzel = exportLessons
    .Select(rec =>
    {
        var dict = (IDictionary<string, object>)rec;
        return dict["teacher"].ToString();
    })
    .Where(x => !string.IsNullOrWhiteSpace(x)) // Nur nicht-leere Strings
    .Distinct()
    .OrderBy(x => x)
    .ToList();

        var alleLehrerImUnterricht = new Lehrers();
        var vergebeneRäume = new Raums();

        foreach (var lehrer in lehrers.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname))
        {
            if (!(from l in alleLehrerImUnterrichtKürzel where lehrer.Kürzel == l select l).Any()) continue;
            // Wenn Raum und Text2 leer sind, dann wird der Lehrer ignoriert 
            if (lehrer.Raum is null or "" && lehrer.Text2 == "") continue;
            alleLehrerImUnterricht.Add(lehrer);

            var r = (from v in vergebeneRäume where v.Raumnummer == lehrer.Raum select v).FirstOrDefault();

            if (r == null)
            {
                if (lehrer.Raum == null) continue;
                // Wenn der Lehrer außer Haus ist, wird sein Raum freigegeben
                if (!lehrer.Text2.ToLower().Contains("außer haus"))
                {
                    vergebeneRäume.Add(new Raum(lehrer.Raum));
                }
            }
            else
            {
                r.Anzahl++;
            }
        }

        zieldatei.Add("====== Sprechtag ======");
        zieldatei.Add("");

        zieldatei.Add(hinweis);

        var i = 1;
        zieldatei.Add("");
        zieldatei.Add("<WRAP column 15em>");
        zieldatei.Add("");
        zieldatei.Add("^Name^Raum^");

        var lehrerProSpalteAufSeite2 = ((alleLehrerImUnterricht.Count - 60) / 3) + 1;

        foreach (var l in alleLehrerImUnterricht.OrderBy(x => x.Nachname))
        {
            var raum = string.IsNullOrEmpty(l.Raum) ? "" : l.Raum;

            // Wenn ein KuK außer Haus ist, wird der Raum aus Untis unterdrückt, bleibt aber in Untis für das nächste Jahr stehen. 
            if (!string.IsNullOrEmpty(l.Text2) && l.Text2.ToLower().Contains("außer haus"))
            {
                raum = "";
            }

            zieldatei.Add(
                "|" + l.Anrede + " " + (l.Titel == "" ? "" : l.Titel + " ") +
                l.Nachname + (l.Text2 == "" ? "" : " ((" + l.Text2 + "))") + "|" + raum + "|");

            if (i == 20 || i == 40 || i == 60 || i == 60 + lehrerProSpalteAufSeite2 ||
                i == 60 + lehrerProSpalteAufSeite2 * 2)
            {
                zieldatei.Add("</WRAP>");
                zieldatei.Add("");

                if (i == 60)
                {
                    zieldatei.Add("<WRAP pagebreak>");
                }

                zieldatei.Add("<WRAP column 15em>");
                zieldatei.Add("");
                zieldatei.Add("^Name^Raum^");
            }

            i++;
        }

        zieldatei.Add("</WRAP>");

        zieldatei.Add(
            "Klassenleitungen finden die Einladung als Kopiervorlage im [[sharepoint>:f:/s/Kollegium2/EjakJvXmitdCkm_iQcqOTLwB-9EWV5uqXE8j3BrRzKQQAw?e=OwxG0N|Sharepoint]].\r\n" +
            Environment.NewLine);

        zieldatei.Add("");

        var xx = raums.Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Field1"].ToString();
        }).Distinct().OrderBy(x => x).ToList();

        var räume = new Raums();

        foreach (var dd in xx)
        {
            if (!(from v in vergebeneRäume where v.Raumnummer == dd select v).Any())
            {
                räume.Add(new Raum(dd));
            }
        }

        var freieR = räume.OrderBy(x => x.Raumnummer)
            .Where(raum => !(from v in vergebeneRäume where v.Raumnummer == raum.Raumnummer select v).Any()).Aggregate(
                @"Verfügbare, freie Räume für den Sprechtag: ",
                (current, raum) => current + (raum.Raumnummer + " "));

        var panel = new Panel(freieR)
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Red);

        AnsiConsole.Write(panel);
        
        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);
    }

    public Datei Zusatzdaten(IConfiguration configuration, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);

        var datums = Quelldateien.GetMatchingList(configuration, "DatumsAusAtlantis", IStudents, Klassen);
        if (!datums.Any())
        {
            return [];
        }

        foreach (var student in IStudents)
        {
            var datumsDiesesSchuelers = datums.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).LastOrDefault();

            var dictD = (IDictionary<string, object>)datumsDiesesSchuelers;

            if (dictD != null)
            {
                dynamic record = new ExpandoObject();

                // Nachname|Vorname|Geburtsdatum|Namenszusatz|Geburtsname|Geburtsort|Ortsteil|Telefon-Nr.|E-Mail|2. Staatsang.|Externe ID-Nr|Sportbefreiung|Fahrschülerart|Haltestelle|Einschulungsart|Entlassdatum|Entlassjahrgang|Datum Schulwechsel|Bemerkungen
                // 

                record.Nachname = 
                    student.Nachname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Vorname =
                    student.Vorname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Geburtsdatum =
                    student.Geburtsdatum; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Namenszusatz = "";
                record.Geburtsname = "";
                record.Geburtsort = "";
                record.Ortsteil = "";
                record.TelefonMINUSNrPUNKT = "";
                record.EMINUSMail = "";
                record.ZWEIPUNKTLEERZEICHENStaatsangPUNKT = "";
                record.ExterneLEERZEICHENIDMINUSNr = "";
                record.Sportbefreiung = "";
                record.Fahrschülerart = "";
                record.Haltestelle = "";
                record.Einschulungsart = "";
                record.Entlassdatum = string.IsNullOrEmpty(dictD["Austrittsdatum"].ToString())
                    ? ""
                    : DateTime.Parse(dictD["Austrittsdatum"].ToString()).ToString("dd.MM.yyyy");
                record.Entlassjahrgang = "";
                record.DatumLEERZEICHENSchulwechsel = "";
                record.Bemerkungen = "";
                record.BKAZVO = "";
                record.BeginnBildungsgang = string.IsNullOrEmpty(dictD["Ausbildungsbeginn"].ToString())
                    ? ""
                    : DateTime.Parse(dictD["Ausbildungsbeginn"].ToString()).ToString("dd.MM.yyyy");
                record.Anmeldedatum = string.IsNullOrEmpty(dictD["Aufnahmedatum"].ToString())
                    ? ""
                    : DateTime.Parse(dictD["Aufnahmedatum"].ToString()).ToString("dd.MM.yyyy");
                record.Bafög = "";
                record.EPMINUSJahre = "";
                record.FaxSLASHMobilnr = "";
                record.Ausweisnummer = "";
                record.schulischeLEERZEICHENEMINUSMail = "";
                record.MasernMINUSImpfnachweis = "";
                zieldatei.Add(record);
            }
            else
            {
                string a = "";
            }
        }

        return zieldatei;
    }

    public Datei Basisdaten(IConfiguration configuration, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        var records = new List<dynamic>();
        var datums = Quelldateien.GetMatchingList(configuration, "DatumsAusAtlantis", IStudents, Klassen);
        var basis = Quelldateien.GetMatchingList(configuration, "basisdaten", IStudents, Klassen);

        foreach (var student in IStudents)
        {
            var basisDiesesSchuelers = basis.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).FirstOrDefault();

            var dictB = (IDictionary<string, object>)basisDiesesSchuelers;

            var datumsDiesesSchuelers = datums.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).ToList();

            if (datumsDiesesSchuelers != null && datumsDiesesSchuelers.Count() > 0)
            {

                // Das älteste Datum zählt.
                var recEintrittsdatum = datumsDiesesSchuelers.OrderBy(dynamic =>
                {
                    var dict = (IDictionary<string, object>)dynamic;
                    return dict["Eintrittsdatum"].ToString();
                }).FirstOrDefault();

                var dictEintrittsdatum = (IDictionary<string, object>)recEintrittsdatum;

                var eintrittsdatum = DateTime.Parse(dictEintrittsdatum["Eintrittsdatum"].ToString())
                    .ToString("dd.MM.yyyy");


                dynamic record = new ExpandoObject();

                // Nachname|Vorname|Geburtsdatum|Geschlecht|Status|PLZ|Ort|Straße|Aussiedler|EINSPUNKTLEERZEICHENStaatsangPUNKT|Konfession|StatistikKrzLEERZEICHENKonfession|Aufnahmedatum|AbmeldedatumLEERZEICHENReligionsunterricht|AnmeldedatumLEERZEICHENReligionsunterricht|SchulpflichtLEERZEICHENerfPUNKT|Reform-Pädagogik|NrPUNKTLEERZEICHENStammschule|Jahr|Abschnitt|Jahrgang|Klasse|Schulgliederung|OrgForm|Klassenart|Fachklasse|NochLEERZEICHENfrei|VerpflichtungLEERZEICHENSprachförderkurs|TeilnahmeLEERZEICHENSprachförderkurs|Einschulungsjahr|ÜbergangsempfPUNKTLEERZEICHENJG5|JahrLEERZEICHENWechselLEERZEICHENS1|1PUNKTLEERZEICHENSchulformLEERZEICHENS1|JahrLEERZEICHENWechselLEERZEICHENS2|Förderschwerpunkt|2PUNKTLEERZEICHENFörderschwerpunkt|Schwerstbehinderung|Autist|LSLEERZEICHENSchulnrPUNKT|LSLEERZEICHENSchulform|Herkunft|LSLEERZEICHENEntlassdatum|LSLEERZEICHENJahrgang|LSLEERZEICHENVersetzung|LSLEERZEICHENReformpädagogik|LSLEERZEICHENGliederung|LSLEERZEICHENFachklasse|LSLEERZEICHENAbschluss|Abschluss|SchulnrPUNKTLEERZEICHENneueLEERZEICHENSchule|Zuzugsjahr|GeburtslandLEERZEICHENSchüler|GeburtslandLEERZEICHENMutter|GeburtslandLEERZEICHENVater|Verkehrssprache|DauerLEERZEICHENKindergartenbesuch|EndeLEERZEICHENEingliederungsphase|EndeLEERZEICHENAnschlussförderung
                // 

                record.Nachname =
                    student.Nachname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Vorname =
                    student.Vorname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Geburtsdatum =
                    student.Geburtsdatum; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                record.Geschlecht = string.IsNullOrEmpty(dictB["Geschlecht"].ToString())
                    ? ""
                    : dictB["Geschlecht"].ToString();
                record.Status = string.IsNullOrEmpty(dictB["Status"].ToString()) ? "" : dictB["Status"].ToString();
                record.PLZ = "";
                record.Ort = "";
                record.Straße = "";
                record.Aussiedler = "";
                record.Staatsang = "";
                record.Konfession = "";
                record.StatistikKrzKonfession = "";
                record.Aufnahmedatum = eintrittsdatum;
                record.AbmeldedatumReligionsunterricht = "";
                record.AnmeldedatumReligionsunterricht = "";
                record.Schulpflichterf = string.IsNullOrEmpty(dictB["Schulpflicht erf."].ToString())
                    ? ""
                    : dictB["Schulpflicht erf."].ToString();
                record.ReformPädagogik = "";
                record.NrStammschule = "";
                record.Jahr = "";
                record.Abschnitt = "";
                record.Jahrgang = "";
                record.Klasse = "";
                record.Schulgliederung = "";
                record.OrgForm = "";
                record.Klassenart = "";
                record.Fachklasse = "";
                record.Nochfrei = "";
                record.VerpflichtungSprachförderkurs = "N";
                record.TeilnahmeSprachförderkurs = "N";
                record.Einschulungsjahr = "";
                record.ÜbergangsempJG5 = "";
                record.JahrWechselS1 = "";
                record.SchulformS1 = "";
                record.JahrWechselS2 = "";
                record.Förderschwerpunkt = "";
                record.Förderschwerpunkt2 = "";
                record.Schwerstbehinderung = string.IsNullOrEmpty(dictB["Schwerstbehinderung"].ToString())
                    ? ""
                    : dictB["Schwerstbehinderung"].ToString();
                record.Autist = string.IsNullOrEmpty(dictB["Autist"].ToString()) ? "" : dictB["Autist"].ToString();
                record.LSSchulnr = "";
                record.LSSchulform = "";
                record.Herkunft = "";
                record.LSEntlassdatum = "";
                record.LSJahrgang = "";
                record.LSVersetzung = "";
                record.LSReformpädagogik = "";
                record.LSGliederung = "";
                record.LSFachklasse = "";
                record.LSAbschluss = "";
                record.Abschluss = "";
                record.SchulnrneueSchule = "";
                record.Zuzugsjahr = "";
                record.GeburtslandSchüler = "";
                record.GeburtslandMutter = "";
                record.GeburtslandVater = "";
                record.Verkehrssprache = "";
                record.DauerKindergartenbesuch = "";
                record.EndeEingliederungsphase = "";
                record.EndeAnschlussförderung = "";
                zieldatei.Add(record);
            }
        }

        return zieldatei;
    }

    public void GetFaecher(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string zieldateiname,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {

        var faecher = Quelldateien.GetMatchingList(configuration, "gpu006", IStudents, Klassen);
        var zieldatei = new Datei(zieldateiname, funktionen, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        var verschiedeneFaecher = faecher.Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Field2"];
        }).ToList().Distinct();

        foreach (var langname in verschiedeneFaecher)
        {
            var fach = faecher.FirstOrDefault(rec =>
            {
                var dict = rec as IDictionary<string, object>;
                return dict != null && dict["Field2"].ToString() == langname.ToString();
            });

            if (langname == "") continue;
            dynamic record = new ExpandoObject();
            record.Name = langname;            
            var dictFach = fach as IDictionary<string, object>;
            if (dictFach != null && dictFach.ContainsKey("Field1"))
            {
                record.Kuerzel = dictFach["Field1"].ToString();
            }            
            zieldatei.Add(record);
        }
        foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
    }

    public void GetLehrer(
        Anrechnungen anrechnungen,
        List<Action<Datei>> funktionen,
        Lehrers lehrers,
        string zieldateiname,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        // alle verschiendenen Lehrkäfte aus anrechnungen und lehrers
        var alleverschiendenenLehrerKürzel = anrechnungen.Select(a => a.LehrerKuerzel)
            .Union(lehrers.Select(l => l.Kürzel))
            .Distinct();

        foreach (var lehrerKuerzel in alleverschiendenenLehrerKürzel)
        {
            var anrechnung = anrechnungen.FirstOrDefault(a => a.LehrerKuerzel == lehrerKuerzel);

            if(anrechnung == null){
                string vorname = "";
            }

            dynamic record = new ExpandoObject();
            record.Kürzel = anrechnung != null ? anrechnung.LehrerKuerzel : lehrerKuerzel;
            record.Vorname = anrechnung != null ? anrechnung.Vorname : "";
            record.Nachname = anrechnung != null ? anrechnung.Nachname : "";
            record.Name = (anrechnung != null && anrechnung.Titel == "" ? "" : anrechnung.Titel + " ") + anrechnung.Vorname + " " + anrechnung.Nachname;
            record.Mail = anrechnung != null ? anrechnung.Mail : "";
            zieldatei.Add(record);
        }

        foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
    }

    public void Praktikanten(
        List<Action<Datei>> funktionen,
        List<string> interessierendeKlassenUndJg,
        string zieldateiname,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var records = new List<dynamic>();
        var zieldatei = new Datei(zieldateiname, funktionen, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        var praktikanten = new List<Student>();

        foreach (var item in interessierendeKlassenUndJg)
        {
            var anzahlStr = item.Split(',')[2];
            int anzahl = 0;
            int.TryParse(anzahlStr, out anzahl);

            for (int i = 1; i <= anzahl; i++)
            {
                foreach (var praktikant in (from s in Students
                                       where s.Klasse.StartsWith(item.Split(',')[0])
                                       where s.Jahrgang.EndsWith(item.Split(',')[1])
                                       where s.Status == "2" || s.Status == "6"
                                       select s).ToList())
                                       {
                                            dynamic record = new ExpandoObject();
                                            record.Name = praktikant.Nachname + ", " + praktikant.Vorname;
                                            record.Klasse = praktikant.Klasse;
                                            record.Jahrgang = praktikant.Jahrgang;
                                            record.Betrieb = "";
                                            record.Betreuung = "";
                                            record.LaufendeNummer = i;
                                            zieldatei.Add(record);
                                       }
            }
        }
        foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
    }

    public void KlassenAnlegen(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string zieldateiname,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var std = Students;
        var kl = Klassen;

        var klassen = Quelldateien.GetMatchingList(configuration, "klassen", Students, new Klassen());

        if (klassen.Count == 0)
        {
            return;
        }

        var records = new List<dynamic>();

        var zieldatei = new Datei(zieldateiname, funktionen, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        foreach (var klasse in klassen.OrderBy(x =>
        {
            var dictKlasse = x as IDictionary<string, object>;
            return dictKlasse?["InternBez"]?.ToString() ?? "";
        }))
        {
            var dictKlasse = klasse as IDictionary<string, object>;            

            dynamic record = new ExpandoObject();
            record.Name = dictKlasse["InternBez"].ToString();
            record.Klassenleitung = dictKlasse["Klassenlehrer"].ToString();
            record.Klassensprecher = "";
            record.Klassensprecher2 = "";
            zieldatei.Add(record);
        }

        foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
    }

    public void Schulpflichtüberwachung(
            IConfiguration configuration,
            string zieldateiname,
            List<string> maßnahmen,
            int anzahlUnentschFehlstunden,
            int schonfrist,
            int warnungAbAnzahl,
            int verjaehrungUnbescholtene,
            int nachSovielenTagenVerjährenFehlzeitenBeiMaßnahme,
            Lehrers lehrers,
            List<Action<Datei>> funktionen        
        )
    {
        Students schuelerMitAbwesenheiten = GetSchuelerMitSovielenUnentschFehlzeiten(configuration, anzahlUnentschFehlstunden);
        schuelerMitAbwesenheiten.GetMassnahmen(configuration, maßnahmen, Quelldateien);
        schuelerMitAbwesenheiten.SchulpflichtüberwachungTxt(
            configuration,
            funktionen,
            zieldateiname,
            schonfrist,                                         // Schonfrist: So viele Tage hat die Klassenleitung Zeit offene Stunden
                                                                // zu bearbeiten, bevor eine Warnung ausgelöst wird.
            warnungAbAnzahl,                                    // Nach so vielen unent. Stunden ohne Maßnahme wird eine Warnung ausgelöst.
            verjaehrungUnbescholtene,                           // Nach so vielen Tagen verjähren unentschuldigte Fehlstunden für Unbescholtene.
            nachSovielenTagenVerjährenFehlzeitenBeiMaßnahme,    // Nach so vielen Tagen verjähren unentschuldigte Fehlstunden für SuS mit Maßnahme
            Klassen,
            lehrers,
            Quelldateien
        );
    }

    public Students GetSchuelerMitSovielenUnentschFehlzeiten(IConfiguration configuration, int anzahl)
    {
        var sMitAbwesenheiten = new Students();

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("SuS mit mehr als " + anzahl + " unentschuldigten Fehlstunden ermitteln ...", ctx =>
        {
            var absencePerStudent = Quelldateien.GetMatchingList(configuration, "absenceperstudent", IStudents, Klassen);

            foreach (Student student in Students)
            {
                if(student.Vorname.StartsWith("Ju") && student.Nachname.StartsWith("Bä"))
                {
                    string aa = "";
                }

                student.GetUnentschFehlzeiten(absencePerStudent);

                if (student.Abwesenheiten.Count != 0)                    
                    if(student.MehrAlsSovieleUnentschuldigteFehlstunden(anzahl))
                        sMitAbwesenheiten.Add(student);
            }
        });

        Global.ZeileSchreiben($"SuS mit mehr als " + anzahl + " unentschuldigten Fehlstunden:", $"{sMitAbwesenheiten.Count}");

        return sMitAbwesenheiten;
    }

    public void GetGruppen(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        Anrechnungen anrechnungen,
        string zieldateiname,
        Lehrers lehrers,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, configuration, funktionen, delimiter, quote, encoding, shouldAllQuote, lehrers, importhinweise);

        var gpu002 = Quelldateien.GetMatchingList(configuration, "gpu002", IStudents, Klassen);
        if (gpu002 == null || gpu002.Count == 0) return;

        var gpu003 = Quelldateien.GetMatchingList(configuration, "gpu003", IStudents, Klassen);
        if (gpu003 == null || gpu003.Count == 0) return;

        Gruppen = new Gruppen();
        Gruppen.AddRange(new Gruppen().GetBildungsgaenge(gpu002, anrechnungen, lehrers));
        Gruppen.AddRange(new Gruppen().GetSchulformen(gpu002, anrechnungen, lehrers));
        Gruppen.Add(new Gruppe().Get(gpu002, lehrers,
                "versetzung:blaue_briefe",
                new List<string>() { "BS", "HBG", "HBT", "HBW", "FS" },
                new List<int>() { 1 }));
        Gruppen.Add(new Gruppe().Get(gpu002, lehrers,
                "termine:fhr:start",
                new List<string>() { "BS", "HBG", "HBT", "HBW", "FS", "FM" },
                new List<int>() { 2 }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:deutsch_kommunikation",
            new List<string>() { "D", "D FU", "D1", "D2", "D G1", "D G2", "D L1", "D L2", "D L", "DL", "DL1", "DL2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:englisch",
            new List<string>() { "E", "E FU", "E1", "E2", "E G1", "E G2", "E L1", "E L2", "E L", "EL", "EL1", "EL2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:englisch",
            new List<string>() { "E", "E FU", "E1", "E2", "E G1", "E G2", "E L1", "E L2", "E L", "EL", "EL1", "EL2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:religionslehre",
            new List<string>() { "KR", "KR FU", "KR1", "KR2", "KR G1", "KR G2", "ER", "ER G1" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:mathematik_physik",
            new List<string>() { "M", "M FU", "M1", "M2", "M G1", "M G2", "M L1", "M L2", "M L", "ML", "ML1", "ML2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:politik_gesellschaftslehre",
            new List<string>() { "PK", "PK FU", "PK1", "PK2", "GG G1", "GG G2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
                ":fachschaften:wirtschaftslehre_in_nicht_kaufm_klassen",
                new List<string>() { "WL", "WBL" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:sport",
            new List<string>() { "SP", "SP G1", "SP G2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(gpu002, lehrers,
            ":fachschaften:biologie",
            new List<string>() { "BI", "Bi", "Bi FU", "Bi1", "Bi G1", "Bi G2", "BI G1", "BI L1" }));

        Gruppen.Add(new Gruppe().GetKollegium(gpu002, lehrers,
            ":kollegium:start"));
        Gruppen.Add(new Gruppe().GetLehrerinnen(anrechnungen, lehrers,
            "kollegium:lehrerinnen"));
        Gruppen.Add(new Gruppe().GetRefs(lehrers,
            "kollegium:referendar_innen"));
        Gruppen.Add(new Gruppe().GetKlassenleitungen(gpu003, lehrers,
            "kollegium:klassenleitungen"));
        Gruppen.Add(new Gruppe().GetBildungsgangleitungen(anrechnungen, lehrers,
            "kollegium:bildungsgangleitungen"));
        Gruppen.Add(new Gruppe().GetByWikilink(anrechnungen, lehrers,
            "kollegium:schulleitung:erweiterte:start"));
        Gruppen.Add(new Gruppe().GetByWikilink(anrechnungen, lehrers,
            "kollegium:lehrerrat"));

        foreach (var gruppe in Gruppen)
            zieldatei.Add(gruppe.Record);
        
        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);
    }

 public void Kalender2Wiki(
  IConfiguration configuration,
  List<Action<Datei>> funktionen,
  string kalender,
  string zieldateiname,
  string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
 {
  var zieldatei = new Datei(zieldateiname + ".csv", funktionen, delimiter, quote, encoding, shouldAllQuote, importhinweise);

  var kalenderRec = Quelldateien.GetMatchingList(configuration, kalender, Students, Klassen);

  if (kalenderRec?.Count != 0)
  {
   var records = new List<dynamic>();

   var sortedRecords = kalenderRec?
    .Where(rec =>
    {
     var beginnString = (string)((IDictionary<string, object>)rec)["Beginn"];
     var kategorienString = (string)((IDictionary<string, object>)rec)["Kategorien"] ?? "";

     return beginnString.Split(" ").Length > 0
     && !string.IsNullOrEmpty(kategorienString)
     && DateTime.ParseExact(beginnString.Split(" ")[1], "dd.MM.yyyy", new CultureInfo("de-DE")) >=
         new DateTime(Convert.ToInt32(Global.AktSj[0]), 07, 31); // keine alten SJ
    })
    .OrderBy(rec =>
    {
     var beginnString = (string)((IDictionary<string, object>)rec)["Beginn"];
     return DateTime.ParseExact(beginnString.Substring(3, beginnString.Length - 3), "dd.MM.yyyy HH:mm",
      new CultureInfo("de-DE"));
    })
    .ToList();

   if (sortedRecords != null)
   {
    foreach (var rec in sortedRecords)
    {
     var dict = (IDictionary<string, object>)rec;
     var beginnString = (string)((IDictionary<string, object>)rec)["Beginn"];
     var endeString = (string)((IDictionary<string, object>)rec)["Ende"];
     var beginnDatum = DateTime.ParseExact(beginnString.Substring(3, beginnString.Length - 3),
      "dd.MM.yyyy HH:mm", new CultureInfo("de-DE"));
     var endeDatum = DateTime.ParseExact(endeString.Substring(3, endeString.Length - 3), "dd.MM.yyyy HH:mm",
      new CultureInfo("de-DE"));
     var dat = beginnDatum.ToString("ddd dd.MM.yyyy", new CultureInfo("de-DE"));
     var zeit = "";

     // Wenn zwischen beginn und ende exakt 24 Stunden oder ein Vielfaches von 24 liegen, dann ist das Ereignis ganztägig
     bool ganztaegig = (endeDatum - beginnDatum).TotalHours % 24 == 0;

     // Bei mehrtägiges, ganztägigen Ereignissen muss das Endedatum um einen Tag nach vorne geschoben werden

     if ((endeDatum - beginnDatum).TotalHours >= 24 && endeDatum.Hour == 0 && endeDatum.Minute == 0 &&
      endeDatum.Second == 0)
     {
      endeDatum = endeDatum.AddDays(-1);
     }

     if (beginnDatum.Hour != 0)
     {
      zeit = ", " + beginnDatum.ToShortTimeString();

      if (endeDatum.Hour != 0)
      {
       zeit += " - " + endeDatum.ToShortTimeString();
      }

      zeit += " Uhr";
     }

     if (ganztaegig && beginnDatum.Date != endeDatum.Date)
     {
      dat += " - " + endeDatum.ToString("ddd dd.MM.yyyy", new CultureInfo("de-DE"));
     }

     var sj = "vergangene";

     if (new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1) < beginnDatum &&
      beginnDatum < new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31))
     {
      sj = "aktuelles";
     }
     if (beginnDatum > new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31))
     {
      sj = "kommendes";
     }

     if (dict["Betreff"].ToString().Contains("QA"))
     {
      string a = "";
     }
     // Wenn in der Nachricht ein Hyperlink enthalten ist, der nach bkb.wiki zeigt, dann wird der Hyperlink aus dem Inhalt der Seite isoliert und einer Variablen namen link zugewiesen.
     var link = dict["Nachricht"].ToString()!.Split(' ').FirstOrDefault(x => x.Contains("bkb.wiki"));

     // Vom link wird nur der Teil hinter dem letzten Slash behalten
     if (link != null && link.Contains("/"))
     {
      link = link.Substring(link.LastIndexOf('/') + 1);
     }
     else
     {
      link = null;
     }

     dynamic record = new ExpandoObject();
     record.Betreff = dict["Betreff"].ToString()!.Trim();
     record.Seite = string.IsNullOrEmpty(link) ? dict["Kategorien"].ToString().Split(';')[0] : link;
     record.Hinweise = "";
     record.Datum = dat + zeit;
     record.Kategorien = GetKategorien(link, dict["Kategorien"].ToString());
     record.Verantwortlich = "";
     record.Ort = dict["Ort"].ToString()!.Trim();
     record.Ressourcen = dict["Ressourcen"].ToString()!.Trim();
     record.SJ = sj;

     if (zieldatei.AbsoluterPfad != null && !zieldatei.AbsoluterPfad.Contains("kollegium"))
     {
     }
     else
     {
      record.Links = "";
     }

     zieldatei.Add(record);
    }
   }

   // Zeige den Link in spectre console   
   AnsiConsole.MarkupLine($"[link=https://bkb.wiki/start?do=admin&page=struct_schemas&table={kalender}#plugin__struct_delete][{Global.GetColor(Global.ColorHyperlink)}]https://bkb.wiki/start?do=admin&page=struct_schemas&table={kalender}[/][/]");

   foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
  }
 }

    private string GetKategorien(string? link, string? toString)
    {
        if (string.IsNullOrEmpty(toString))
        {
            return string.Empty;
        }

        var kategorien = toString.ToLower().Split(';').Aggregate("", (current, str) => current + (str.Trim() + ","));

        if (!string.IsNullOrEmpty(link) && !kategorien.Contains(link.ToLower()))
        {
            kategorien = link + "," + kategorien;
        }

        return kategorien.TrimEnd(',');
    }

    public void Teilleistungen(
        IConfiguration configuration,
        string zieldateiname,
        Lehrers lehrers,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        configuration = Global.Konfig("Teilleistungsarten", Global.Modus.Update, configuration);
        configuration = Global.Konfig("Abschnitt", Global.Modus.Update, configuration);

        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);
        var records = new List<dynamic>();

        var marksPerLs = Quelldateien.GetMatchingList(configuration, "marksperlesson", IStudents, Klassen);
        if (marksPerLs == null) return;
        if (marksPerLs.Count == 0) throw new Exception("In der Quelldatei marksperlesson wurden keine Datensätze gefunden.");

        try
        {
            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Teilleistungen aus MarksPerLesson einlesen ...", ctx =>
            {
                foreach (var student in IStudents)
                {
                    foreach (var recMar in marksPerLs)
                    {
                        var dictMar = (IDictionary<string, object>)recMar;
                        if (!(dictMar["Name"].ToString() == student.Nachname + " " + student.Vorname && dictMar["Klasse"].ToString() == student.Klasse)) continue;
                        if (!(string.IsNullOrEmpty(configuration["Teilleistungsarten"].ToString()) || configuration["Teilleistungsarten"].ToString().ToLower().Trim().Split(',').Contains(dictMar["Prüfungsart"].ToString().ToLower()))) continue;
                        dynamic record = new ExpandoObject();
                        record.Nachname = $"{student.Nachname}#{student.Klasse}";
                        record.Vorname = student.Vorname;
                        record.Geburtsdatum = student.Geburtsdatum;
                        record.Jahr = Global.AktSj[0];
                        record.Abschnitt = configuration["Abschnitt"];
                        record.Fach = dictMar["Fach"].ToString();
                        record.Datum = dictMar["Datum"].ToString();
                        record.Teilleistung = dictMar["Prüfungsart"].ToString();
                        record.Note = student.GetNote(dictMar["Note"].ToString());
                        record.Bemerkung = dictMar["Bemerkung"].ToString();
                        record.Lehrkraft = dictMar["Benutzer"].ToString();
                        zieldatei.Add(record);
                    }
                }

                if (zieldatei.Count == 0)
                {
                    var panel = new Panel($"Keine Teilleistungen gefunden. Haben Sie die Teilleistungsart(en) exakt so eingegeben, wie sie in SchILD bzw. Webuntis mit Langname heißt?")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Color.Red);

                    AnsiConsole.Write(panel);
                }
            });
            
            var teilleistungen = Quelldateien.GetMatchingList(configuration, "schuelerteilleistungen", IStudents, Klassen);

            var leistungsdaten = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
            if (leistungsdaten == null ||leistungsdaten.Count == 0)
            {
                var panel = new Panel($"In der Quelldatei schuelerLeistungsdaten.dat wurden keine Datensätze gefunden. Also kann nicht ermittelt werden, ob Teilleistungen fehlen.")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Red);

                AnsiConsole.Write(panel);
            }
            else
            {
                var alleVerschiedenenLehrer = leistungsdaten.Select(r =>
                {
                    var dict = (IDictionary<string, object>)r;
                    return dict["Fachlehrer"].ToString();
                }).Distinct().OrderBy(f => f).ToList();

                var meldung = new List<string>();
                var urlMitte = "";

                AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Leistungsdaten aus Schülerleistungsdaten einlesen ...", ctx =>
                {
                    foreach (var lehrer in alleVerschiedenenLehrer)
                    {
                        // Durchlaufe alle Leistungsdaten.
                        foreach (var leistungsdatensatz in leistungsdaten.Where(r =>
                        {
                            var dict = (IDictionary<string, object>)r;
                            return dict["Fachlehrer"].ToString() == lehrer;
                        }))
                        {
                            // Suche in den zieldatei.Records nach einem Eintrag, der zu dem Leistungsdatensatz passt.
                            var dictLeistungsdatensatz = (IDictionary<string, object>)leistungsdatensatz;

                            var matchingRecord = zieldatei.FirstOrDefault(r =>
                            {
                                var dictRecord = (IDictionary<string, object>)r;
                                return dictRecord["Nachname"].ToString().Split('#')[0] == $"{dictLeistungsdatensatz["Nachname"]}"
                                    && dictRecord["Vorname"].ToString() == dictLeistungsdatensatz["Vorname"].ToString()
                                    && dictRecord["Fach"].ToString() == dictLeistungsdatensatz["Fach"].ToString()
                                    && configuration["Teilleistungsarten"].ToString().ToLower().Contains(dictRecord["Teilleistung"].ToString().ToLower());
                            });

                            if(matchingRecord == null)
                            {
                                // Wenn kein passender Wert gefunden wurde, durchsuche die isherigen Teilleistungsdaten.
                                matchingRecord = teilleistungen.FirstOrDefault(r =>
                                {
                                    var dictRecord = (IDictionary<string, object>)r;
                                    return dictRecord["Nachname"].ToString().Split('#')[0] == $"{dictLeistungsdatensatz["Nachname"]}"
                                        && dictRecord["Vorname"].ToString() == dictLeistungsdatensatz["Vorname"].ToString()
                                        && dictRecord["Fach"].ToString() == dictLeistungsdatensatz["Fach"].ToString()
                                        && configuration["Teilleistungsarten"].ToString().Contains(dictRecord["Teilleistung"].ToString());
                                });                                
                            }

                            // Wenn kein passender Eintrag gefunden wurde, gebe eine Warnung aus.
                            if (matchingRecord == null)
                            {
                                meldung.Add($"{dictLeistungsdatensatz["Nachname"]} {dictLeistungsdatensatz["Vorname"]} ({dictLeistungsdatensatz["Fach"]}, {lehrer})");

                                var lehrerMail = lehrers.FirstOrDefault(l => l.Kürzel == lehrer)?.Mail;

                                if (urlMitte.Split(',').Contains(lehrerMail) == false)
                                {
                                    urlMitte += lehrerMail + ",";                                    
                                }
                            }
                        }
                    }
                });

                // Nachricht an den Lehrer erstellen
                if (meldung.Count > 0)
                {
                    var anzahl = meldung.Count;
                
                    var nursoviele = 10;
                    var panel = new Panel($"{anzahl}x wurden keine Teilleistungen eingetragen. Die LuL werden im Chat informiert.\n\n" +
                        $"{meldung.Take(nursoviele).Select(m => $"- {m}").Aggregate((current, next) => current + "\n" + next)}")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Red);

                    AnsiConsole.Write(panel);
                }
                zieldatei.UrlMitte = urlMitte.TrimEnd(',');
                zieldatei.UrlRechts = "&message=" + Uri.EscapeDataString("Hallo LuL ");
            }

            foreach (var aktion in zieldatei.Funktionen)
                aktion(zieldatei);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            while (Console.KeyAvailable) Console.ReadKey(true);

            Console.ReadKey();
        }
    }

    internal Datei GetLehrerDerKlassen(IConfiguration configuration, Lehrers lehrers)
    {
        var datei = new Datei();
        var mitgliederMail = "";

        var recGpu002 = Quelldateien.GetMatchingList(configuration, "gpu002", IStudents, Klassen);
        if (recGpu002.Count == 0) return datei;

        var verschiedeneLulKuerzel = recGpu002
            .Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                var klassenString = dict["Field5"].ToString();
                var klassenListe = klassenString.Split('~'); // Zerlegt den String in eine Liste
                return IKlassen.Any(klasse => klassenListe.Contains(klasse)) &&
                       !string.IsNullOrEmpty(dict["Field6"].ToString());
            }).Select(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Field6"].ToString();
            }).Distinct().ToList().OrderBy(x => x).ToList();

        dynamic record = new ExpandoObject();

        foreach (var lulKuerzel in verschiedeneLulKuerzel)
        {
            var lehrer = lehrers.Where(rec =>
            {
                return rec.Kürzel == lulKuerzel.ToString();
            }).FirstOrDefault();

            if (lehrer != null)
            {
                mitgliederMail += lehrer.Mail + ";";
            }
        }

        record.MitgliederMail = mitgliederMail.TrimEnd(';');
        datei.Add(record);
        return datei;
    }

    /// <summary>
    /// Erstellt die Datei LehrkraefteSonderzeiten.dat mit den Lehrkräften und deren Anrechnungsgründen.
    /// Diese Datei kann in SchILD importiert werden.
    /// </summary>
    /// <param name="configuration"></param>
    /// <param name="zieldateiname"></param>
    /// <param name="nurDieseGründe"></param>
    /// <returns></returns>
    internal void LehrkraefteSonderzeiten(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null,
        string defaultwert = "",        
        Global.Modus modus = Global.Modus.Update)
    {
        var akt = int.Parse(Global.AktSj[0]);

        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        var lehrkraefte = Quelldateien.GetMatchingList(configuration, "lehrkraefte", IStudents, Klassen);
        if (lehrkraefte == null || lehrkraefte.Count == 0) return;

        var lehrkraefteSonderzeiten = Quelldateien.GetMatchingList(configuration, "lehrkraefteSonderzeiten", IStudents, Klassen);
        if (lehrkraefteSonderzeiten == null) return;

        var gpu020 = Quelldateien.GetMatchingList(configuration, "gpu020", IStudents, Klassen); // Anrechnungen
        if (gpu020 == null) return;

        var gpu004 = Quelldateien.GetMatchingList(configuration, "gpu004", IStudents, Klassen); // Lehrkraefte
        if (gpu004 == null) return;

        var lehrers = lehrkraefte
            .Where(rec => ((IDictionary<string, object>)rec)["statistik-relevant"].ToString() == "J")
            .ToList();

        var lehrerString = lehrers
            .Select(rec => ((IDictionary<string, object>)rec)["InternKrz"].ToString())
            .Aggregate("", (current, str) => current + (str + ","))
            .TrimEnd(',');        

        configuration = Global.Konfig("NurDieseLehrer", modus, configuration, "", -1, -1, null, "", null, lehrerString);

        var nurDieseLehrer = configuration["NurDieseLehrer"].ToString().Trim();

        // Wenn die "Altersermäßigung"-Funktion aufgerufen wird (200), werden alle Lehrkräfte als Default hinzugefügt.
        if (defaultwert == "200")
            nurDieseLehrer = lehrerString;
        
        var alleMöglichenVerschiedenenGründe = gpu020
                    .Where(rec => ((IDictionary<string, object>)rec)["Field13"].ToString() != "0")
                    .Where(rec => ((IDictionary<string, object>)rec)["Field5"].ToString().Length != 0)
                    .Select(rec => ((IDictionary<string, object>)rec)["Field5"].ToString().Trim().TrimStart('0'))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

        // Wenn die Sonderzeitenfunktion aufgerufen wird, besteht Wahlfreiheit
        if (defaultwert != "200")
            configuration = Global.Konfig("NurDieseGruende", modus, configuration, "", -1, -1, null, "", null, string.Join(',', alleMöglichenVerschiedenenGründe));

        var interessierendeGründe = configuration["NurDieseGruende"].ToString().Trim();

        // Wenn die "Altersermäßigung"-Funktion aufgerufen wird (200), wird nur dieser Grund gesetzt.
        if (defaultwert == "200")
            interessierendeGründe = "200";

        // Wenn die Sonderzeitenfunktion aufgerufen wird, besteht Wahlfreiheit
        if (defaultwert != "200")
            configuration = Global.Konfig("DieseGruendeIgnorieren", Global.Modus.Update, configuration, "", -1, -1, null, "", null, string.Join(',', alleMöglichenVerschiedenenGründe));

        var ignorierteGründe = configuration["DieseGruendeIgnorieren"].ToString().Trim();
        
        if (defaultwert == "200")
            ignorierteGründe = "";

        // Alle interessierenden Gründe minus ignorierte Gründe
        var interessierendeMinusIgnorierte = interessierendeGründe
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Where(gr => !ignorierteGründe.Split(',', StringSplitOptions.RemoveEmptyEntries).Contains(gr))
                .ToList();
                
        // Wenn die Altersermäßigung (auch) interessiert, dann wird die Stelle abgefragt
        if (interessierendeMinusIgnorierte.Contains("200"))
            configuration = Global.Konfig("VolleStelle", Global.Modus.Update, configuration);

        var panel1 = new Panel($"Die Anrechnungen aus der Untis-Datei [aqua]GPU020.txt[/] werden mit der SchILD-Datei [aqua]LehrkraefteSonderzeiten.dat[/] abgeglichen." +
                        $"\nDie Datei [aqua]{zieldateiname}[/] wird neu erstellt und kann nach ScHILD importiert werden." +
                        $"\nAnrechnungsgründe [aqua]...........................[/] werden ignoriert und auf 0 gesetzt.")
                        .Header($"[bold springGreen2] Anrechnungen [/]")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Spectre.Console.Color.SpringGreen2);

        // Wenn die Altersermäßigung (auch) interessiert, dann wird das Panel mit den Hinweisen angezeigt
        if (interessierendeMinusIgnorierte.Contains("200"))
        {
            var panel = new Panel($"Im Folgenden werden alle Lehrkräfte mit Altersermäßigung und alle mit fehlendem Deputat bzw. fehlendem Geburtsdatum angezeigt. " +
                                "Ohne Angabe des Deputats und Geburtsdatums in SchILD, findet keine Berechnung statt. " +
                                "Sch/Unt zeigt (möglicherweise) abweichende Werte in SchILD und in Untis. " +
                                $"Die Datei [aqua]{zieldateiname}[/] wird neu erstellt und kann nach ScHILD importiert werden. " +
                                $"Da die Datei nur zusammen mit einer (leeren) lehrkraefte.dat importiert werden kann, wird eine leere lehrkraefte.dat ebenfalls erzeugt. " +
                                $"Wenn in der Spalte [bold springGreen2]Alter am 31.7.{akt + 1}[/] die Zahl 55 oder 60 steht, dann ändern sich die Werte im kommenden Jahr. ")
                    .HeaderAlignment(Justify.Left)
                    .SquareBorder()
                    .Expand()
                    .BorderColor(Color.Grey);

            AnsiConsole.Write(panel);
        }
        
        var tableAltersermäßigung = new Table();

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Lehrkraefte prüfen ...", ctx =>
        {            
            if (interessierendeMinusIgnorierte.Contains("200"))
            {
                //table.AddColumn(new TableColumn("Nr."));
                tableAltersermäßigung.AddColumn(new TableColumn("LuL"));
                tableAltersermäßigung.AddColumn(new TableColumn("Geb.dat."));
                tableAltersermäßigung.AddColumn(new TableColumn("Deput. lt. Sch/Unt"));
                tableAltersermäßigung.AddColumn(new TableColumn("Stelle in %"));
                tableAltersermäßigung.AddColumn(new TableColumn($"Alter am 31.7.{akt}"));
                tableAltersermäßigung.AddColumn(new TableColumn($"Alter am 31.7.{akt + 1}"));
                tableAltersermäßigung.AddColumn(new TableColumn($"IST laut Sch/Unt"));
                tableAltersermäßigung.AddColumn(new TableColumn($"SOLL {akt}/ {akt + 1} lt. SchILD"));
                tableAltersermäßigung.AddColumn(new TableColumn($"SOLL {akt + 1}/ {akt + 2} lt.SchILD"));
            }

            foreach (var lehrerRec in lehrers)
            {
                var lehrer = new Lehrer();
                // Spalte 2
                lehrer.Kürzel = ((IDictionary<string, object>)lehrerRec)["InternKrz"].ToString();

                if (nurDieseLehrer != "" && !nurDieseLehrer.Split(',').ToList().Contains(lehrer.Kürzel)) continue;

                // Spalte 3
                lehrer.Geburtsdatum = lehrer.GetGeburtsdatum(((IDictionary<string, object>)lehrerRec)["Geburtsdatum"].ToString());
                // Spalte 4
                lehrer.DeputatLautSchild = lehrer.GetDeputatSchild(((IDictionary<string, object>)lehrerRec)["Pflichtstunden-Soll"].ToString());

                foreach (var grund in interessierendeMinusIgnorierte)
                {
                    var istWertSonderzeitLautSchild = lehrer.GetWertSonderzeiten(grund, lehrkraefteSonderzeiten);
                    var istWertGpu020LautUntis = lehrer.GetAnrechnungswertGPU020Soll(gpu020, grund);
                    lehrer.DeputatLautUntis = lehrer.GetDeputatLautUntis(gpu004);

                    if (interessierendeGründe.Contains("200") && grund == "200")
                    {
                        // Spalte 4      
                        lehrer.ProzentStelleInSchild = lehrer.GetProzentStelle(configuration);
                        // Spalte 5
                        lehrer.AlterAmErstenSchultagDesJahres = lehrer.GetAlterAmErstenSchultagDesSchuljahres(akt);
                        // Spalte 6
                        lehrer.AlterAmErstenSchultagDesKommendenJahres = lehrer.GetAlterAmErstenSchultagDesSchuljahres(akt + 1);
                        // Spalte 7
                        lehrer.AltersermäßigungSoll = lehrer.CheckAltersermäßigungSoll(lehrer.AlterAmErstenSchultagDesJahres);
                        // Spalte 8
                        lehrer.AltersermäßigungSollKommendes = lehrer.CheckAltersermäßigungSoll(lehrer.AlterAmErstenSchultagDesKommendenJahres);

                        // Wenn der Lehrer spätestens am 31.7. des kommenden SJ 55 Jahre alt ist, erscheint er in der Liste
                        if (lehrer.AlterAmErstenSchultagDesKommendenJahres >= 55)
                        {
                            var deputatLautSchildUndUntis = (lehrer.DeputatLautSchild == 0 ? "-" : lehrer.DeputatLautSchild % 1 == 0
                                ? ((int)lehrer.DeputatLautSchild).ToString()
                                : lehrer.DeputatLautSchild.ToString("0.##", CultureInfo.InvariantCulture)) + "/" + (lehrer.DeputatLautUntis == 0 ? "-" : lehrer.DeputatLautUntis % 1 == 0
                                ? ((int)lehrer.DeputatLautUntis).ToString()
                                : lehrer.DeputatLautUntis.ToString("0.##", CultureInfo.InvariantCulture));

                            var istSchildUndUntis = (istWertSonderzeitLautSchild == 0 ? "-" : istWertSonderzeitLautSchild.ToString("0.##", CultureInfo.InvariantCulture)) + "/" +
                                (istWertGpu020LautUntis == 0 ? "-" : istWertGpu020LautUntis.ToString("0.##", CultureInfo.InvariantCulture));

                            tableAltersermäßigung.AddRow(
                                lehrer.Kürzel,
                                lehrer.Geburtsdatum.ToString("dd.MM.yy"),
                                deputatLautSchildUndUntis,
                                lehrer.ProzentStelleInSchild == 0 ? "-" : lehrer.ProzentStelleInSchild.ToString("0.##", CultureInfo.InvariantCulture),
                                lehrer.AlterAmErstenSchultagDesJahres == 0 ? "-" : lehrer.AlterAmErstenSchultagDesJahres.ToString(),
                                lehrer.AlterAmErstenSchultagDesKommendenJahres == 0 ? "-" : lehrer.AlterAmErstenSchultagDesKommendenJahres.ToString(),
                                istSchildUndUntis,
                                lehrer.AltersermäßigungSoll == 0 ? "-" : lehrer.AltersermäßigungSoll.ToString("0.##", CultureInfo.InvariantCulture),
                                lehrer.AltersermäßigungSollKommendes == 0 ? "-" : lehrer.AltersermäßigungSollKommendes.ToString("0.##", CultureInfo.InvariantCulture));
                        }
                    }

                    var wertDerAnrechnung = 0.0;
                    
                    var wertIst = lehrkraefteSonderzeiten
                    .Where(rec => ((IDictionary<string, object>)rec)["Lehrkraft"].ToString() == lehrer.Kürzel)
                    .Where(rec => ((IDictionary<string, object>)rec)["Grund"].ToString() == grund)
                    .Select(rec => ((IDictionary<string, object>)rec)["Anzahl Stunden"].ToString())
                    .FirstOrDefault();

                    if (string.IsNullOrWhiteSpace(wertIst))
                        wertIst = "0";
                    else
                        wertIst = wertIst.Replace(',', '.');

                    wertDerAnrechnung = lehrer.GetAnrechnungswertGPU020Soll(gpu020, grund);

                    dynamic record = new ExpandoObject();
                    record.Lehrkraft = lehrer.Kürzel;
                    record.Zeitart = grund.ToString().StartsWith("1") ? "MEHRLEISTUNG" : grund.ToString().StartsWith("2") ? "MINDERLEISTUNG" : "ANRECHNUNG";
                    record.Grund = grund;
                    record.AnzahlLEERZEICHENStunden = interessierendeGründe != "200" ? wertDerAnrechnung.ToString().Replace('.', ',') : lehrer.AltersermäßigungSoll.ToString("F2", CultureInfo.InvariantCulture).Replace('.', ',');
                    // Nullwerte werden nicht exportiert, weil das in ASDPC zu Fehlermeldungen führt.
                    if (record.AnzahlLEERZEICHENStunden != "0,00" && record.AnzahlLEERZEICHENStunden != "0")
                        zieldatei.Add(record);
                }
            }
        });

        if(interessierendeGründe.Contains("200"))
            if(tableAltersermäßigung.Rows.Count > 0)
                AnsiConsole.Write(tableAltersermäßigung);
    }

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

                AnsiConsole.Write(new Rule($"[bold fuchsia] {datFiles.Count} Dateien verschoben von [bold aqua]{configuration["PfadSchilddatenaustausch"]}[/] nach [bold aqua]{configuration["PfadDownloads"]}[/][/] ").RuleStyle("fuchsia").LeftJustified());                
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

    public void Lehrkraefte(
        IConfiguration configuration,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert, string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        dynamic record = new ExpandoObject();
        record.InternKrz = "";
        record.StatistikKrz = "";
        record.Anrede = "";
        record.Titel = "";
        record.Nachname = "";
        record.Vorname = "";
        record.Geschlecht = "";
        record.Geburtsdatum = "";
        record.Staatsang = "";
        record.PLZ = "";
        record.Ort = "";
        record.Straße = "";
        record.TelPUNKTLEERZEICHENFestnetz = "";
        record.TelPUNKTLEERZEICHENmobil = "";
        record.EMINUSMail = "";
        record.Rechtsverhältnis = "";
        record.Beschäftigungsart = "";
        record.Einsatzstatus = "";
        record.PflichtstundenMINUSSoll = "";
        record.Schulleitung = "";
        record.statistikMINUSrelevant = "";
        record.StammschulnrPUNKT = "";
        record.dienstlPUNKTLEERZEICHENEMINUSMail = "";
        zieldatei.Add(record);
    }

    public void RenderAuswahlÜberschrift(IConfiguration configuration)
    {
        var panel3 = new Panel($"[grey85]{string.Join("\n", Beschreibung)}[/]")
            .Header($"[bold yellow3_1] {Titel.Split(':')[2]} [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Color.Yellow3_1);
        
        AnsiConsole.Write(panel3);
    }

 internal void SchuelerZusatzdatenUmMailAdresseErgaenzen(
  IConfiguration configuration,
  string zieldateiname,
  List<Action<Datei>> funktionen,
  string[] anhandDieserAttributeWirdVerglichen,
  string[] dieseAttributeWerdenBeimVergleichIgnoriert,
  string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
 {
  var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", IStudents, Klassen);
  if (schuelerZusatzdaten == null || !schuelerZusatzdaten.Any()) throw new Exception("Keine SchuelerZusatzdaten.dat");

  var schuelerBasisdaten = Quelldateien.GetMatchingList(configuration, "schuelerbasisdaten", IStudents, Klassen);
  if (schuelerBasisdaten == null || !schuelerBasisdaten.Any()) throw new Exception("Keine SchuelerBasisdaten.dat");

  bool problem = false;

  var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

  if (schuelerZusatzdaten.Count != Students.Count)
  {
   var panel = new Panel($"Die Anzahl der in der Datei SchuelerZusatzdaten.dat ({schuelerZusatzdaten.Count}) stimmt nicht mit der Anzahl der SchuelerBasisdaten.dat ({Students.Count}) überein. Exportieren Sie die Dateien erneut. Anschließend kehren Sie hierher zurück!")
      .HeaderAlignment(Justify.Left)
      .SquareBorder()
      .Expand()
      .BorderColor(Global.ColorFehler);

   AnsiConsole.Write(panel);
   return;
  }

  var mailDomain = configuration["MailDomain"].Trim();
  var mehrfachVorhanden = new List<dynamic>();

  AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Schuelerzusatzdaten: Mails anpassen ...", ctx =>
  {
   for (int i = 0; i < schuelerZusatzdaten.Count; i++)
   {    
    var dictSz = (IDictionary<string, object>)schuelerZusatzdaten[i];
    var dictSb = (IDictionary<string, object>)schuelerBasisdaten[i];

    // Wenn der Schüler nicht aktiv oder extern ist, überspringe diese Zeile
    if (!(dictSb["Status"].ToString() == "2" || dictSb["Status"].ToString() == "6")) continue;

    if (string.IsNullOrEmpty(dictSb["Geburtsdatum"].ToString()))
    {
     var panel = new Panel($"Der Schüler {dictSz["Nachname"]} {dictSz["Vorname"]} hat kein Geburtsdatum in den Individualdaten I hinterlegt. Bitte ergänzen Sie das Geburtsdatum in SchILD, damit eine schulische E-Mail-Adresse generiert werden kann.")
       .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
       .HeaderAlignment(Justify.Left)
       .SquareBorder()
       .Expand()
       .BorderColor(Global.ColorFehler);
     AnsiConsole.Write(panel);
     problem = true;
     continue;
    }

    dynamic record = new ExpandoObject();

    if (MehrfachVorhanden(
     schuelerZusatzdaten,
     dictSz["schulische E-Mail"].ToString(),
     dictSz["Nachname"].ToString(),
     dictSz["Vorname"].ToString(),
     dictSz["Geburtsdatum"].ToString()))
    {
     mehrfachVorhanden.Add(schuelerZusatzdaten[i]);
    }

    foreach (var prop in dictSz)
    {
     var name = prop.Key;
     var value = prop.Value;

     if (name == "Nachname")
     {
      var klasse = dictSb["Klasse"].ToString();

      ((IDictionary<string, object>)record)[name] = $"{value}#{klasse}";
     }
     else if (name == "schulische E-Mail")
     {
      // Schüler mit vorhandener Mail überspringen
      if (!string.IsNullOrEmpty(value.ToString()))
      {
       ((IDictionary<string, object>)record)[name] = value;
      }
      else
      {   
       var student = Students.FirstOrDefault(s => (s.Status == "2" || s.Status == "6") && s.Nachname == dictSz["Nachname"].ToString() && s.Vorname == dictSz["Vorname"].ToString() && s.Geburtsdatum == dictSz["Geburtsdatum"].ToString());

       if (student == null)
       {
        // Wenn der Schüler nicht gefunden wurde, überspringe diese Zeile
        ((IDictionary<string, object>)record)[name] = value;
        continue;
       }
       else
       {
        // Wenn es in den Zusatzdaten einen Schüler gibt, mit identischem Namen, Vornamen, Geburtsdatum,
        // und bei dem die schulische E-Mail-Adresse ebenfalls übereinstimmt, dann gib die Mail-Adresse aus:
        var mail = schuelerZusatzdaten
         .Where(s =>
         {
          var dic = s as IDictionary<string, object>;
          return dic != null &&
           dic["Nachname"].ToString() == dictSz["Nachname"].ToString() &&
           dic["Vorname"].ToString() == dictSz["Vorname"].ToString() &&
           dic["Geburtsdatum"].ToString() == dictSz["Geburtsdatum"].ToString() &&
           !string.IsNullOrEmpty(dic["schulische E-Mail"].ToString());
         })
         .Select(s => ((IDictionary<string, object>)s)["schulische E-Mail"].ToString())
         .FirstOrDefault();

        if (!string.IsNullOrEmpty(mail))
        {
         ((IDictionary<string, object>)record)["schulische E-Mail"] = mail;
        }
        else
        {
         if (DateTime.TryParse(student.Geburtsdatum, out DateTime gebDatum))
         {
          var n = student.Bereinigen(student.Nachname.ToLower()).Substring(0, 1);
          var v = student.Bereinigen(student.Vorname.ToLower()).Substring(0, 1);
          var geburtsjahr = gebDatum.Year.ToString().Substring(2, 2);
          var geburtsmonat = gebDatum.Month.ToString("D2");
          var geburtstag = gebDatum.Day.ToString("D2");

          var schulischeEmail = $"{n}{v}{geburtsjahr}{geburtsmonat}{geburtstag}{mailDomain}";

          // Wenn die E-Mail-Adresse bereits existiert, dann hänge eine Zahl an       
          var counter = 1;
          while (schuelerZusatzdaten.Any(s => ((IDictionary<string, object>)s)["schulische E-Mail"].ToString() == schulischeEmail))
          {
           schulischeEmail = $"{n}{v}{geburtsjahr}{geburtsmonat}{geburtstag}{counter}{mailDomain}";
           counter++;
          }

          // Wenn der Counter größer als 1 ist, dann gib ein Panel aus
          if (counter > 1)
          {
           var panel = new Panel($"Die E-Mail-Adresse für {student.Vorname} {student.Nachname} ({student.Geburtsdatum}) soll neu angelegt werden, wurde aber bereits zuvor in SchILD vergeben. Deswegen schreibt [{Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] jetzt für {student.Vorname} {student.Nachname} einen Zähler vor das [{Global.GetColor(Global.ColorHinweise)}]@[/]: [{Global.GetColor(Global.ColorZahlen)}]{schulischeEmail}[/]. So wird Eindeutigkeit gewährleistet.\nOptional können Sie nach dem SchILD-Import die E-Mail-Adresse von {student.Vorname} nochmal ändern, wenn Sie anderweitig Eindeutigkeit herstellen wollen (z.B. Buchstabe statt Zähler).\nWeiter mit [{Global.GetColor(Global.ColorActionInMenüs)}]ENTER[/].")
            .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] Hinweis: Doppelung bei E-Mail-Adresse [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Global.ColorHinweise);
           AnsiConsole.Write(panel);
           Console.ReadKey(true);
          }

          student.MailSchulisch = schulischeEmail;

          ((IDictionary<string, object>)record)[name] = schulischeEmail;

          if (MehrfachVorhanden(
           schuelerZusatzdaten,
           dictSz["schulische E-Mail"].ToString(),
           dictSz["Nachname"].ToString(), dictSz["Vorname"].ToString(), dictSz["Geburtsdatum"].ToString()))
          {
           mehrfachVorhanden.Add(schuelerZusatzdaten[i]);
          }
         }
        }
       }
      }
     }
     else
     {
      ((IDictionary<string, object>)record)[name] = value;
     }
    }

    zieldatei.Add(record);
   }  
  });

  Global.ZeileSchreiben("Schuelerzusatzdaten: Mails überprüft", zieldatei.Count.ToString());

  foreach (var aktion in zieldatei.Funktionen)
   aktion(zieldatei);

  // Sortiere mehrfachVorhanden nach Geburtsdatum
  mehrfachVorhanden = mehrfachVorhanden.OrderBy(s => s.Geburtsdatum).ToList();

  if (mehrfachVorhanden.Count > 0)
  {
   var schüler = string.Join("\n ", mehrfachVorhanden.Select(s => $"{s.Geburtsdatum} {s.Nachname} {s.Vorname}"));
   var fehler = $"Unter [{Global.GetColor(Global.ColorPfadInProgrammen)}]Individualdaten I[/] haben mehrere dieselbe schulinterne Mailadresse: \n {schüler} \nLösen Sie das Problem, indem Sie in SchILD unter [{Global.GetColor(Global.ColorPfadInProgrammen)}]Individualdaten I[/] händisch Eindeutigkeit herstellen. Sie könnten z.B. bei einer/einem Schüler*in händisch eine [{Global.GetColor(Global.ColorZahlen)}]1[/] anhängen.\nAnschließend exportieren Sie alle *.dat-Dateien erneut und kehren hierher zurück.";
   throw new Exception(fehler);
  }

  if (problem)
  {
   throw new Exception($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren![/]");
  }
 }

    private string TelefonNummerFormatieren(string? v)
    {
        var telefon = v.ToString();

        // Leere Telefonnummern werden ignoriert
        if (string.IsNullOrEmpty(telefon))
        {
            return telefon;
        }

        // Ausländische Telefonnummern werden ignoriert
        if (
            !string.IsNullOrEmpty(telefon) &&
            (telefon.StartsWith("00") || telefon.StartsWith("+")) &&
            !(telefon.StartsWith("0049") || telefon.StartsWith("+49"))
            )
        {
            return telefon;
        }

        // Manchmal ist eine Mailadresse als Telefonnummer eingetragen. Diese wird ignoriert.
        if (!string.IsNullOrEmpty(telefon) && (telefon.Contains("@")))
        {
            return telefon;
        }

        // Bei Deutschen Telefonnummern wird das Ländervorwahl-Präfix +49 entfernt, wenn es vorhanden ist.                    
        if (telefon.StartsWith("+49"))
        {
            telefon = telefon.Replace("+49 ", "0");
        }
        if (telefon.StartsWith("+49"))
        {
            telefon = telefon.Replace("+49", "0");
        }

        if (telefon.StartsWith("0049"))
        {
            telefon = telefon.Replace("0049", "0");
        }
        if (telefon.StartsWith("0049 "))
        {
            telefon = telefon.Replace("0049 ", "0");
        }

        // Alle Bindestriche, Schrägstriche, Klammern werden entfernt.
        var tel = telefon.Replace(" ", "").Replace("-", " ").Replace("/", " ").Replace("(", "").Replace(")", " ").Trim();

        // Doppelte Leerzeichen werden entfernt
        while (tel.Contains("  "))
        {
            tel = tel.Replace("  ", " ");
        }

        // Mit vorwahlen abgleichen und die Leerstelle entsprechend gesetzt
        foreach (var vorwahl in Global.Vorwahlen)
        {
            if (tel.StartsWith(vorwahl) && tel.Length > vorwahl.Length)
            {
                // Entferne zuerst alle Leerstellen
                tel = tel.Replace(" ", "");
                // Füge dann eine Leerstelle nach der Vorwahl ein
                tel = tel.Insert(vorwahl.Length, " ");
                break;
            }
        }
        return tel;
    }

    private bool MehrfachVorhanden(List<dynamic> schuelerZusatzdaten, string mail, string nachname, string vorname, string geburtsdatum)
    {
        // Prüfe auf doppelte
        var doppelte = schuelerZusatzdaten.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return mail == dict["schulische E-Mail"].ToString() && mail != "" &&
                !string.IsNullOrEmpty(dict["Nachname"].ToString()) &&
                !string.IsNullOrEmpty(dict["Vorname"].ToString()) &&
                !string.IsNullOrEmpty(dict["Geburtsdatum"].ToString());
        }).ToList();

        // Wenn mehr als 1 Eintrag existiert, aber alle Felder identisch sind, ist es KEIN Duplikat
        if (doppelte.Count > 1)
        {
            // Prüfe, ob alle Einträge in Name, Vorname und Geburtsdatum identisch sind
            bool alleIdentisch = doppelte.All(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Nachname"].ToString() == nachname &&
                    dict["Vorname"].ToString() == vorname &&
                    dict["Geburtsdatum"].ToString() == geburtsdatum;
            });

            if (alleIdentisch)
                return false; // Kein Duplikat, sondern dieselbe Person
            return true; // Unterschiedliche Personen mit gleicher Mail
        }
        return false;
    }

    internal bool NichtAlleSusHabenEineEindeutigeMailAdresse(IConfiguration configuration, Students students)
    {
        var problem = false;
        var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", students, Klassen);
        if (schuelerZusatzdaten == null || schuelerZusatzdaten.Count == 0) return false;

        var sz = schuelerZusatzdaten
                .Where(rec =>
                {
                    if (rec == null) return false;
                    var dict = (IDictionary<string, object>)rec;
                    return dict != null && string.IsNullOrEmpty(dict["schulische E-Mail"].ToString());
                }).ToList();

        if (sz.Count > 0)
        {
            AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("Fehlende Mailadressen ...", ctx =>
            {
                foreach (var s in sz)
                {
                    var dict = (IDictionary<string, object>)s;
                    var panel = new Panel(
                                $"Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]Schuelerzusatzdaten.dat[/] enthält keine schulische E-Mail-Adresse für {dict["Vorname"]} {dict["Nachname"]}. " +
                                $"\nNutzen Sie [{Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/], um automatisch alle Mailadressen in [{Global.GetColor(Global.ColorPfadInDateien)}]Schuelerzusatzdaten.dat[/] zu erstellen." +
                                $"\nAnschließend kehren Sie hierher zurück.")
                            .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Global.ColorFehler);
                    AnsiConsole.Write(panel);
                }
                Global.ZeileSchreiben($"Fehlende Mailadressen:", $"{sz.Count}");
            });

            problem = true;            
        }

        // Annahme: schuelerZusatzdaten ist List<IDictionary<string, object>>
        var doppelteMitAbweichung = schuelerZusatzdaten
            .Where(sz => sz is IDictionary<string, object> dict && !string.IsNullOrEmpty(dict["schulische E-Mail"]?.ToString()))
            .GroupBy(sz => ((IDictionary<string, object>)sz)["schulische E-Mail"]?.ToString())
            .Where(g =>
                g.Count() > 1 &&
                g.Select(sz => (
                    Vorname: ((IDictionary<string, object>)sz)["Vorname"]?.ToString(),
                    Nachname: ((IDictionary<string, object>)sz)["Nachname"]?.ToString()
                ))
                .Distinct(new TupleStringComparer()).Count() > 1
            )
            .SelectMany(g => g) // Alle betroffenen Zeilen ausgeben
            .ToList();

        AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start("Doppelt vergebene Mailadressen ...", ctx =>
        {
            foreach (IDictionary<string, object> ssz in doppelteMitAbweichung)
            {
                var panel = new Panel($"Die E-Mail-Adresse [{Global.GetColor(Global.ColorPfadInDateien)}]{ssz["schulische E-Mail"]}[/] ist mehrfach vergeben, aber die Namen sind unterschiedlich: " +
                                    $"{ssz["Vorname"]} {ssz["Nachname"]}")
                            .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Global.ColorFehler);
                AnsiConsole.Write(panel);
                problem = true;
            }
            if (doppelteMitAbweichung.Count > 0)
            {
                Global.ZeileSchreiben($"Doppelt vergebene Mailadressen:", $"{doppelteMitAbweichung.Count}");
            }
        });

        if (problem)
        {
            //AnsiConsole.MarkupLine($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren.[/]");
                throw new Exception($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren![/]");
            }

        return problem;
    }

    internal List<string> GetFotosAusSchildPfade(IConfiguration configuration, Students Students, Enum zipModus)
    {
        var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", Students, Klassen);
        if (schuelerZusatzdaten == null || schuelerZusatzdaten.Count == 0) return [];
            
        var absolutePfade = new List<string>();

        var pfadFotosImSchILD = configuration["PfadFotosImSchILD-Ordner"];

        // Suche nach allen Fotos im SchILD-Ordner und in Unterordnern        
        var alleFotos = Directory.GetFiles(pfadFotosImSchILD, "*.jpg", SearchOption.AllDirectories);

        var students = Students;

        AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .Start($"Schüler*innen für {zipModus} verarbeiten ...", ctx =>
        {
            foreach (var student in students)
            {
                var bisherigerDateinameUndPfad = alleFotos.FirstOrDefault(f =>
                    f.Contains(student.Nachname, StringComparison.OrdinalIgnoreCase) &&
                    f.Contains(student.Vorname, StringComparison.OrdinalIgnoreCase) &&
                    f.Contains(student.Geburtsdatum, StringComparison.OrdinalIgnoreCase));

                // Benenne das Foto im Ordner nach kurzname um
                if (bisherigerDateinameUndPfad != null)
                {
                    var neuerDateiname = "";

                    var sz = schuelerZusatzdaten
                        .Where(rec =>
                        {
                            if (rec == null) return false;
                            var dict = (IDictionary<string, object>)rec;
                            return dict != null && dict["Nachname"] != null && dict["Nachname"].ToString() == student.Nachname &&
                                dict["Vorname"].ToString() == student.Vorname &&
                                dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                        }).LastOrDefault() as IDictionary<string, object>;

                    if (sz != null)
                    {
                        if ((Global.ZipModus)zipModus == Global.ZipModus.Webuntis)
                        {
                            neuerDateiname = sz["schulische E-Mail"].ToString().Replace(configuration["MailDomain"], "").Replace("@", "");
                        }
                        else if ((Global.ZipModus)zipModus == Global.ZipModus.Geevoo)
                        {
                            neuerDateiname = sz["schulische E-Mail"].ToString();
                        }
                    }

                    if (!string.IsNullOrEmpty(neuerDateiname))
                    {
                        var dateityp = Path.GetExtension(bisherigerDateinameUndPfad).ToLowerInvariant();
                        var pfad = Path.GetDirectoryName(bisherigerDateinameUndPfad);
                        var neuerPfadUndDateinameWebuntis = Path.Combine(pfad, neuerDateiname + dateityp);

                        // Wenn die neue Datei schon existiert, überspringe sie
                        if (File.Exists(neuerPfadUndDateinameWebuntis)) continue;

                        // Erstelle Webuntis-kompatiblen Dateinamen
                        File.Copy(bisherigerDateinameUndPfad, neuerPfadUndDateinameWebuntis, true);

                        absolutePfade.Add(neuerPfadUndDateinameWebuntis);
                    }
                }
            }
            Global.ZeileSchreiben($"Fotos aus SchILD für {zipModus}", $"{absolutePfade.Count}");
        });

        if (absolutePfade.Count == 0)
            {
                Global.ZeileSchreiben("Keine neuen Fotos gefunden", "0", ConsoleColor.Red, ConsoleColor.White);
            }
        return absolutePfade;
    }

    internal Datei? KurseLehrkraefte(IConfiguration configuration, string zieldateiname, Unterrichte kurse)
    {
        var zieldatei = new Datei(zieldateiname);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("KurseLehrkräfte: ...", ctx =>
        {
            foreach (var kurs in kurse)
            {
                // Alle Lehrkraefte des Kurses werden durchlaufen
                for (int i = 0; i < kurs.Lehrkraefte.Count; i++)
                {
                    var lehrer = kurs.Lehrkraefte[i];
                    if (lehrer == null || string.IsNullOrEmpty(kurs.Lehrkraefte[i])) continue;

                    // Prüfen, ob der Lehrer bereits in der Datei ist
                    // if (zieldatei.Any(rec => rec.Lehrkraft == kurs.Lehrkraefte[i])) continue;

                    dynamic record = new ExpandoObject();
                    record.KursBez = kurs.KursBez.Substring(0, Math.Min(kurs.KursBez.Length, 20));
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = configuration["Abschnitt"].ToString();
                    record.Lehrkraft = kurs.Lehrkraefte[i];
                    record.Wochenstd = kurs.LehrkraefteWochenstunden[i];
                    record.Jahrgang = ""; // kann leer bleiben                
                    zieldatei.Add(record);
                }
            }
            Global.ZeileSchreiben("KurseLehrkräfte:", zieldatei.Count().ToString());
        });
        return zieldatei;
    }

    internal void NeueFotosAusSchildOrdnerErstellenUndAlteFotosVerschieben(IConfiguration configuration)
    {
        var pfadFotosAusSchild = configuration["PfadFotosAusSchild"];
        var ordnername = new DirectoryInfo(pfadFotosAusSchild).Name;
        var übergeordneterOrdner = Directory.GetParent(pfadFotosAusSchild)?.FullName;

        var ordnernameBisher = ordnername + "_bisher";

        var neuerOrdnername = $"{ordnername}_{DateTime.Now:yyyyMMdd_HHmmss}";
        var neuerOrdnerPfad = Path.Combine(Path.Combine(übergeordneterOrdner ?? "", ordnernameBisher), neuerOrdnername);
        
        var panel = new Panel(
            $"Nach dem erfolgreichen Zippen werden nun die Fotos aus [{Global.GetColor(Global.ColorPfadInDateien)}]{pfadFotosAusSchild}[/] nach [{Global.GetColor(Global.ColorPfadInDateien)}]{neuerOrdnerPfad}[/] verschoben. Das ist notwendig, damit BKB-Tool beim nächsten Mal aus dem Vergleich der neuen und alten Fotos die Zipdatei passend erstellen kann. " +
            $"Weiter mit [{Global.GetColor(Global.ColorActionInMenüs)}]ENTER[/]. Abbruch mit [bold red]x[/].");
        panel.Header($"[bold {Global.GetColor(Global.ColorHinweise)}] Hinweis: Fotos-Ordner wird verschoben [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Global.ColorHinweise);
        AnsiConsole.Write(panel);

        var key = Console.ReadKey(true);
        if (key.Key == ConsoleKey.X)
        {
            throw new OperationCanceledException("Sie haben abgebrochen.");
        }

        // Erstelle neuen Ordner mit aktuellem Datum
        Directory.CreateDirectory(neuerOrdnerPfad);

        // Verschiebe alle jpg-Dateien aus dem alten Ordner in den neuen Ordner
        var dateien = Directory.GetFiles(pfadFotosAusSchild, "*.jpg");
        Global.ZeileSchreiben($"Neuer Ordner erstellt:", neuerOrdnername);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Alte Fotos verschieben ...", ctx =>
        {
            foreach (var datei in dateien)
            {
                var dateiname = Path.GetFileName(datei);
                var neuerDateipfad = Path.Combine(neuerOrdnerPfad, dateiname);
                File.Move(datei, neuerDateipfad);
            }
        });
        
        Global.ZeileSchreiben($"Alte Fotos verschoben:", dateien.Length.ToString());
    }

 internal void OeffneExistierendeDateienOderNeueInEditor(IConfiguration configuration, List<string> dateien)
 {
  var notepadPlusPlusCandidates = new[]
  {
   @"C:\Program Files\Notepad++\notepad++.exe",
   @"C:\Program Files (x86)\Notepad++\notepad++.exe"
  };

  var pfadDownloads = configuration["PfadDownloads"];
  if (string.IsNullOrEmpty(pfadDownloads) || !Directory.Exists(pfadDownloads))
  {
   AnsiConsole.MarkupLine($"[red]Der Ordner {pfadDownloads} existiert nicht.[/]");
   return;
  }

  // Baue die vollständigen Pfade zusammen
  var vollstaendigeDateien = dateien
   .Select(datei => Path.Combine(pfadDownloads, datei))
   .Where(File.Exists)
   .ToList();

  if (vollstaendigeDateien.Count == 0)
  {
   if (this.Quelldateien.Where(q => !string.IsNullOrEmpty(q.Fehlermeldung) && !q.IstOptional && !q.Nur177659).Any())
   {
    AnsiConsole.MarkupLine($"[red]Keine der angegebenen Dateien wurde im Ordner {pfadDownloads} gefunden.[/]");
    // Falls nicht optinale Dateen fehlen, dann Abbruch
    return;
   }
   else
   {
    AnsiConsole.MarkupLine($"[yellow]Keine der angegebenen Dateien wurde im Ordner {pfadDownloads} gefunden. Sie werden jetzt erstellt.[/]");
    // Fehlende optinale Dateien werden erstellt
    vollstaendigeDateien = dateien
     .Select(datei => Path.Combine(pfadDownloads, datei))
     .ToList();
     foreach (var datei in vollstaendigeDateien)     
     {
      File.WriteAllText(datei, "", Encoding.UTF8);
     }
   }}

  // Prüfe auf veraltete Dateien (bestehendes Verhalten beibehalten)
  var maxAlter = DateTime.Now.Date.AddDays(-0);
  var dateienDieAelterSind = vollstaendigeDateien
   .Where(datei => File.GetLastWriteTime(datei) < maxAlter)
   .ToList();

  if (dateienDieAelterSind.Count > 0)
  {
   AnsiConsole.MarkupLine($"[yellow]Die folgenden Dateien sind älter als vom {maxAlter:dd.MM.yyyy} Tage:[/]");
   foreach (var datei in dateienDieAelterSind)
   {
    AnsiConsole.MarkupLine($"[yellow]- {Path.GetFileName(datei)} (letzte Änderung: {File.GetLastWriteTime(datei)})[/]");
   }
   throw new Exception("Aktualisieren Sie die Dateien. Kehren Sie dann hierher zurück.");
  }

  try
  {
   // Windows: versuche Notepad++ zuerst, sonst Standardeditor
   if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
   {
    var notepadPlusPlusPath = notepadPlusPlusCandidates.FirstOrDefault(File.Exists);

    if (!string.IsNullOrEmpty(notepadPlusPlusPath))
    {
     // Notepad++ kann mehrere Dateien in einem Aufruf öffnen
     var args = string.Join(" ", vollstaendigeDateien.Select(f => $"\"{f}\""));
     Process.Start(new ProcessStartInfo
     {
      FileName = notepadPlusPlusPath,
      Arguments = args,
      UseShellExecute = false
     });
     return;
    }

    // Fallback: Standardeditor (Dateizuordnung) verwenden
    foreach (var file in vollstaendigeDateien)
    {
     Process.Start(new ProcessStartInfo
     {
      FileName = file,
      UseShellExecute = true
     });
    }
    return;
   }

   // Linux: öffne mit xdg-open (Standardanwendung). Falls xdg-open nicht verfügbar, versuche UseShellExecute.
   if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux))
   {
    foreach (var file in vollstaendigeDateien)
    {
     try
     {
      Process.Start(new ProcessStartInfo
      {
       FileName = "xdg-open",
       Arguments = $"\"{file}\"",
       UseShellExecute = false
      });
     }
     catch (Exception)
     {
      // Fallback: direkte Shell-Öffnung (funktioniert auf einigen Umgebungen)
      try
      {
       Process.Start(new ProcessStartInfo
       {
        FileName = file,
        UseShellExecute = true
       });
      }
      catch (Exception ex)
      {
       AnsiConsole.MarkupLine($"[red]Datei konnte nicht geöffnet werden: {file} — {ex.Message}[/]");
      }
     }
    }
    return;
   }

   // Sonstige OS: Standardöffner verwenden
   foreach (var file in vollstaendigeDateien)
   {
    Process.Start(new ProcessStartInfo
    {
     FileName = file,
     UseShellExecute = true
    });
   }
  }
  catch (Exception ex)
  {
   AnsiConsole.MarkupLine($"[red]Fehler beim Öffnen der Datei(en): {ex.Message}[/]");
  }
 }

 internal void OeffneWebseite(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Fehler beim Öffnen der Webseite: {ex.Message}[/]");
        }
    }

    internal Datei SchueleradresseTelefonFormatieren(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string zieldateiname,
        string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert, string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var schuelerAdressen = Quelldateien.GetMatchingList(configuration, "schueleradressen", IStudents, Klassen);
        if (schuelerAdressen == null || !schuelerAdressen.Any()) return [];

        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Schueleradressen: Mails und Telefonnummern anpassen ...", ctx =>
        {
            foreach (var schueler in schuelerAdressen)
            {
                var dict = (IDictionary<string, object>)schueler;

                dynamic record = new ExpandoObject();

                foreach (var prop in dict)
                {
                    var name = prop.Key;
                    var value = prop.Value;

                    if (name == "Nachname")
                    {
                        var student = Students.LastOrDefault(s =>
                            s.Nachname == dict["Nachname"].ToString() &&
                            s.Vorname == dict["Vorname"].ToString() &&
                            s.Geburtsdatum == dict["Geburtsdatum"].ToString());

                        ((IDictionary<string, object>)record)[name] = value.ToString();
                    }
                    else if (name.Contains(". Tel.-Nr."))
                    {
                        var tel = TelefonNummerFormatieren(value.ToString());
                        ((IDictionary<string, object>)record)[name] = tel;
                    }
                    else
                    {
                        ((IDictionary<string, object>)record)[name] = value;
                    }
                }
                zieldatei.Add(record);
            }
        });
        Global.ZeileSchreiben("Schueleradressen: Mails und Telefonnummern angepasst", zieldatei.Count.ToString());

        return zieldatei;
    }

    internal Datei AdresseTelefonFormatieren(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string zieldateiname,
        string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert, string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var adressen = Quelldateien.GetMatchingList(configuration, "adressen", IStudents, Klassen);
        if (adressen == null || !adressen.Any()) return [];

        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Adressen: Mails und Telefonnummern anpassen ...", ctx =>
        {
            foreach (var schueler in adressen)
            {
                var dict = (IDictionary<string, object>)schueler;

                dynamic record = new ExpandoObject();

                foreach (var prop in dict)
                {
                    var name = prop.Key;
                    var value = prop.Value;

                    if (name == "Nachname")
                    {
                        var student = Students.LastOrDefault(s =>
                            s.Nachname == dict["Nachname"].ToString() &&
                            s.Vorname == dict["Vorname"].ToString() &&
                            s.Geburtsdatum == dict["Geburtsdatum"].ToString());

                        ((IDictionary<string, object>)record)[name] = $"{value}#{student.Klasse}";
                    }
                    else if (name.Contains("Telefonnr. ") || name.Contains("Fax"))
                    {
                        var tel = TelefonNummerFormatieren(value.ToString());
                        ((IDictionary<string, object>)record)[name] = tel;
                    }
                    else
                    {
                        ((IDictionary<string, object>)record)[name] = value;
                    }
                }
                zieldatei.Add(record);
            }
        });
        Global.ZeileSchreiben("Adressen: Mails und Telefonnummern angepasst", zieldatei.Count.ToString());

        return zieldatei;
    }

    /// <summary>
    /// Achtung: Telefonnummern sind immer additiv.
    /// </summary>
    /// <param name="configuration"></param>
    /// <param name="zieldateiname"></param>
    /// <param name="anhandDieserAttributeWirdVerglichen"></param>
    /// <param name="dieseAttributeWerdenBeimVergleichIgnoriert"></param>
    /// <param name="delimiter"></param>
    /// <param name="quote"></param>
    /// <param name="encoding"></param>
    /// <param name="shouldAllQuote"></param>
    /// <param name="importhinweise"></param>
    /// <returns></returns>
    internal Datei SchuelerTelefonnummernFormatieren(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string zieldateiname,
        string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert, string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var schuelertelefonnummern = Quelldateien.GetMatchingList(configuration, "schuelertelefonnummern", IStudents, Klassen);
        if (schuelertelefonnummern == null || !schuelertelefonnummern.Any()) return [];

        var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("SchuelerTelefonnummern: Telefonnummern anpassen ...", ctx =>
        {
            foreach (var schueler in schuelertelefonnummern)
            {
                var dict = (IDictionary<string, object>)schueler;

                dynamic record = new ExpandoObject();

                foreach (var prop in dict)
                {
                    var name = prop.Key;
                    var value = prop.Value;

                    if (name.Contains("Telefonnr."))
                    {
                        var tel = TelefonNummerFormatieren(value.ToString());
                        ((IDictionary<string, object>)record)[name] = tel;
                    }
                    else
                    {
                        ((IDictionary<string, object>)record)[name] = value;
                    }
                }
                zieldatei.Add(record);
            }
        });
        Global.ZeileSchreiben("SchuelerTelefonnummern: Telefonnummern angepasst", zieldatei.Count.ToString());

        return zieldatei;
    }

    public void GetUntisAnrechnungen(
        Anrechnungen anrechnungen,
        string zieldateiname,
        List<Action<Datei>> funktionen,
        List<int> nurDieseGrunde,
        List<int> furDieseGrundeKeinenWert,
        List<string?> furDieseLehrerKeineWerte,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, funktionen, delimiter, quote, encoding, shouldAllQuote, importhinweise);
        
        try
        {
            foreach (var anrechnung in anrechnungen.OrderBy(a => a.LehrerKuerzel).ThenBy(a => a.Grund))
            {
                if (anrechnung.LehrerKuerzel == "KUH")
                {
                    var debug = 1;
                }

                if (!nurDieseGrunde.Contains(anrechnung.Grund)) continue;
                var wert = (anrechnung.Wert == 0 ? "" : anrechnung.Wert.ToString(CultureInfo.InvariantCulture));

                if (!furDieseGrundeKeinenWert.Contains(anrechnung.Grund))
                {
                    wert = "";
                }

                if (furDieseLehrerKeineWerte.Contains(anrechnung.LehrerKuerzel))
                {
                    wert = "";
                }

                var kategorien = "";
                if (anrechnung.Kategorien != null)
                    kategorien = anrechnung.Kategorien.Aggregate(kategorien, (current, c) => current + (c + ","));

                anrechnung.Name = (anrechnung.Titel == "" ? "" : anrechnung.Titel + " ") +
                                  anrechnung.Vorname + " " + anrechnung.Nachname;

                dynamic record = new ExpandoObject();
                record.Name = anrechnung.Name;
                record.Kuerzel = anrechnung.LehrerKuerzel;
                record.Mail = anrechnung.Mail;
                record.Wert = wert;
                record.von = (anrechnung.Von.Year == 1 ? "" : anrechnung.Von.ToShortDateString());
                record.bis = (anrechnung.Bis.Year == 1 ? "" : anrechnung.Bis.ToShortDateString());
                record.Rolle = anrechnung.Rolle;
                record.Amt = anrechnung.Amt;
                record.Grund = anrechnung.Grund.ToString();
                record.Beschreibung = (anrechnung.Beschr == "" ? "" : "[[" + anrechnung.Beschr + "]]");
                record.Hinweis = anrechnung.Hinweis;
                record.Kategorien = kategorien.TrimEnd(',');

                zieldatei.Add(record);
            }

            foreach (var aktion in zieldatei.Funktionen) aktion(zieldatei);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            while (Console.KeyAvailable) Console.ReadKey(true);

            Console.ReadKey();
        }
    }

 internal void KlausurbelegungWikiSeiteErstellen(
  IConfiguration configuration,
  string zieldateiname,
  List<Action<Datei>> funktionen)
 {
  var leistungsdaten = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
  if (leistungsdaten == null || !leistungsdaten.Any()) return;

  var gpu002 = Quelldateien.GetMatchingList(configuration, "gpu002", IStudents, Klassen);
  if (gpu002 == null || !gpu002.Any()) return;

  var faecher = Quelldateien.GetMatchingList(configuration, "faecher", IStudents, Klassen);
  if (faecher == null || !faecher.Any()) return;

  var zieldatei = new Datei(zieldateiname, funktionen, configuration);

  // Gib alle Fächer distinct als List<string> zurück, sortiert nach Sortierung S2
  var alleFaecherDistinct = faecher
   .Select(rec => (IDictionary<string, object>)rec)
   .OrderBy(dict => dict["Sortierung S2"].ToString())
   .Select(dict => dict["InternKrz"].ToString()) // Nur Teil vor Leerzeichen
   .Distinct()   
   .ToList();

  var faecherInDieserKlasseAusLeistungsdaten = alleFaecherDistinct
   .Where(fach =>
    leistungsdaten.Any(rec =>
    {
     var dict = (IDictionary<string, object>)rec;
     // Nur Teil vor Leerzeichen vergleichen
     return dict.ContainsKey("Fach") && dict["Fach"].ToString() == fach;
    }))
   .ToList();

  var faecherInDieserKlasseAusGpu002 = alleFaecherDistinct
   .Where(fach =>
    gpu002.Any(rec =>
    {
     var dict = (IDictionary<string, object>)rec;
     // Nur Teil vor Leerzeichen vergleichen
     return dict.ContainsKey("Field7") && dict["Field7"].ToString() == fach;
    }))
   .ToList();

  var verschiedeneKlassenDerStudents = IStudents.Select(s => s.Klasse).Distinct().OrderBy(k => k).ToList();

  foreach (var klasse in verschiedeneKlassenDerStudents)
  {
   zieldatei.Name = $"oeffentlich:Klausurbelegung:{klasse}-{configuration["InteressierendesSchuljahr"]}-{configuration["Abschnitt"]}";
   zieldatei.Add($"====== Klausurbelegung {klasse} Schuljahr: {configuration["InteressierendesSchuljahr"]} Abschnitt: {configuration["Abschnitt"]} ======");
   zieldatei.Add("");

   var kopfzeile = "^  Name  ^ ";
   var kopfzeile2 = "^  ^";

   // Prüfe für alle IStudents dieser Klasse, ob die Initialen (z.B. A.B.) mehrfach vorkommen.
   // Wenn ja, dann muss für diese Klasse die Initialen soweit verlängert werden, bis der Wert eindeutig ist.
   int anzahlZeichen = 1;
   for (int i = 1; i <= 5; i++) // Maximal bis 5 Zeichen
   {
    var initialenGruppiert = IStudents
     .Where(s => s.Klasse == klasse)
     .Select(s => new
     {
      Student = s,
      Initialen = s.Nachname.Substring(0, i) + "." + s.Vorname.Substring(0, i) + "."
     })
     .GroupBy(x => x.Initialen)
     .Where(g => g.Count() > 1)
     .ToList();

    if (initialenGruppiert.Count == 0)
    {
     anzahlZeichen = i;
     break; // Alle Initialen sind eindeutig
    }
   }

   if (configuration["Klausurbelegung"] == "1")
   {
    foreach (var fach in faecherInDieserKlasseAusGpu002)
    {
     // Alle verschiedenen Lehrerkürzel, die das Fach (ohne Ziffer und nur vor dem Leerzeichen) unterrichten
     var alleLehrerDiesesFachs = gpu002
      .Where(rec =>
      {
       var dict = (IDictionary<string, object>)rec;
       return dict["Field7"].ToString().StartsWith(fach);
      })
      .Select(rec => (IDictionary<string, object>)rec)
      .GroupBy(dict => dict["Field6"].ToString())
      .Select(g => g.First())
      .Select(dict => new
      {
       Kürzel = dict["Field6"].ToString()
      })
      .OrderBy(l => l.Kürzel)
      .ToList();

     foreach (var lehrer in alleLehrerDiesesFachs)
     {
      kopfzeile2 += lehrer.Kürzel + " ";
     }
     kopfzeile += fach + "  ^  ";
     kopfzeile2 += "^";
    }
   }
   else if (configuration["Klausurbelegung"] == "2")
   {
    foreach (var fach in faecherInDieserKlasseAusLeistungsdaten)
    {
     // Alle verschiedenen Lehrerkürzel, die das Fach (ohne Ziffer und nur vor dem Leerzeichen) unterrichten
     var alleLehrerDiesesFachs = leistungsdaten
      .Where(rec =>
      {
       var dict = (IDictionary<string, object>)rec;
       return dict["Fach"].ToString() == fach;
      })
      .Select(rec => (IDictionary<string, object>)rec)
      .GroupBy(dict => dict["Fachlehrer"].ToString().Split(' ')[0])
      .Select(g => g.First())
      .Select(dict => new
      {
       Kürzel = dict["Fachlehrer"].ToString().Split(' ')[0]
      })
      .OrderBy(l => l.Kürzel)
      .ToList();

     foreach (var lehrer in alleLehrerDiesesFachs)
     {
      kopfzeile2 += lehrer.Kürzel + " ";
     }
     kopfzeile += fach + "  ^  ";
     kopfzeile2 += "  ^  ";
    }
   }
   
   zieldatei.Add(kopfzeile);
   zieldatei.Add(kopfzeile2);

   foreach (var student in IStudents.Where(s => s.Klasse == klasse).OrderBy(s => s.Nachname).ThenBy(s => s.Vorname))
   {
    var nameKurz = student.Nachname.Substring(0, anzahlZeichen) + "." + student.Vorname.Substring(0, anzahlZeichen) + ".";

    var zeile = "| " + nameKurz;

    if (configuration["Klausurbelegung"] == "1")
    {
     foreach (var fach in faecherInDieserKlasseAusGpu002)
     {
      zeile += "| ";
     }
    }
    else if (configuration["Klausurbelegung"] == "2")
    {   // Wenn Leistungsdaten genutzt werden, dann auch die Inhalte in die Zellen übergeben
     foreach (var fach in faecherInDieserKlasseAusLeistungsdaten)
     {
      // Suche die Kursart aus den Leistungsdaten für diesen Schüler
      var kursart = leistungsdaten
       .Where(rec =>
       {
        var dict = (IDictionary<string, object>)rec;
        return dict["Fach"].ToString() == fach &&
            dict["Nachname"].ToString() == student.Nachname &&
            dict["Vorname"].ToString() == student.Vorname &&
            dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
       })
       .Select(rec => (IDictionary<string, object>)rec)
       .Select(dict => dict["Kursart"].ToString())
       .FirstOrDefault();
      kursart = kursart;// (kursart == null ? "" : new List<string> { "PUK", "GK" }.Contains(kursart) ? "X" : kursart);

      zeile += "| " + kursart + " ";
     }
    }
    else if (configuration["Klausurbelegung"] == "3")
    {
     throw new Exception("Unbekannte Einstellung für Klausurbelegung: " + configuration["Klausurbelegung"]);
    }
    
    zeile += "| ";
    zieldatei.Add(zeile);
   }
   zieldatei.Add("Klassenleitungen (zusammen mit Schüler*innen) nehmen hier im Wiki die Klausurbelegung vor. Unmittelbar vor den Zeugnissen wird die Tabelle nach SchILD importiert. Zulässige Werte:");
   zieldatei.Add("  * LK1, LK2, AB3, AB4 (für die 4 Abiturfächer)");
   zieldatei.Add("  * GKS, GKM (für schriftliche und mündliche Grundkurse)");

   foreach (var aktion in zieldatei.Funktionen)
    aktion(zieldatei);
  }
 }

 internal void KlausurbelegungAusWikiNachSchildEinlesen(
  IConfiguration configuration,
  string zieldateiname,
  List<Action<Datei>> funktionen,
  string[] anhandDieserAttributeWirdVerglichen,
  string[] dieseAttributeWerdenBeimVergleichIgnoriert,
  string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
 {
  var zieldatei = new Datei(zieldateiname, funktionen, anhandDieserAttributeWirdVerglichen, dieseAttributeWerdenBeimVergleichIgnoriert, delimiter, quote, encoding, shouldAllQuote, importhinweise);

  var leistungsdaten = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
  if (!leistungsdaten.Any())
   throw new Exception("Es sind keine Leistungsdaten vorhanden.");  

  var dokuwikiZugriff = new DokuwikiZugriff(configuration);

  dokuwikiZugriff.Options = new XmlRpcStruct
  {
   { "sum", "Automatische Aktualisierung" },
   { "minor", Global.WikiSprechtagKleineAenderung } // Kein Minor-Edit
  };

  var verschiedeneKlassenDerStudents = IStudents.Select(s => s.Klasse).Distinct().OrderBy(k => k).ToList();

  foreach (var klasse in verschiedeneKlassenDerStudents)
  {
   // Wiki-Seite lesen
   string seitenName = $"oeffentlich:Klausurbelegung:{klasse}-{configuration["InteressierendesSchuljahr"]}-{configuration["Abschnitt"]}"; // Pfad im Wiki
   var seitenInhalt = dokuwikiZugriff.Proxy.GetPage(seitenName);

   // Tabelle parsen
   var tabelle = ParseDokuWikiTable(seitenInhalt);

   var maxSpalten = 20; // Maximal 9 Spalten anzeigen
   var table = new Spectre.Console.Table();
   table.Title = new TableTitle($"Tabelle [{Global.GetColor(Global.ColorPfadInDateien)}]{seitenName}[/] aus Wiki");
   table.Border(TableBorder.Rounded);
   table.Expand();
   
   // Die ersten beiden Zeilen als Header verwenden
   int spalten = Math.Min(tabelle[0].Count, maxSpalten);
   for (int j = 0; j < spalten; j++)
   {
    var header = tabelle[0][j];
    if (tabelle.Count > 1 && tabelle[1].Count > j)
     header += "\n" + tabelle[1][j];
    table.AddColumn(header);
   }
   if (tabelle[0].Count > maxSpalten)
   {
    table.AddColumn("...");
   }
   
   int laengeDesNamens = 0;

   // Ab Zeile 2 die Daten einfügen
   for (int i = 2; i < tabelle.Count; i++)
   {
    var row = tabelle[i].Take(maxSpalten).ToList();

    laengeDesNamens = row[0].Split('.')[0].Length;

    if (tabelle[i].Count > maxSpalten)
     row.Add("...");
    table.AddRow(row.ToArray());
   }
   
   AnsiConsole.Write(table);

   // Beispiel: Zugriff auf einzelne Zelle
   //Console.WriteLine("\nZelle [2][0] = " + tabelle[2][0]);

   //"Durchlaufe alle Zeilen der bisherigen Leistungsdaten mit dem Ziel die Kursart zu aktualisieren"
   
   foreach (var leistung in leistungsdaten)
   {
    var dict = (IDictionary<string, object>)leistung;

    // suche den Spaltenindex aus Tabelle, der dem Fach entspricht
    var fach = dict["Fach"].ToString();
    var spaltenIndex = tabelle[0].IndexOf(fach);
    if (spaltenIndex == -1)
    {
     // Fach nicht gefunden, überspringe diesen Eintrag
     continue;
    }
    // suche den Zeilenindex aus Tabelle, der dem Schüler entspricht
    var schuelername = dict["Nachname"].ToString().Substring(0, laengeDesNamens) + "." + dict["Vorname"].ToString().Substring(0, laengeDesNamens) + "."; // Nur erstes Zeichen des Vornamens
    var zeilenIndex = tabelle.FindIndex(row =>
     row.Count > 0 &&
     row[0] == schuelername
    );
    if (zeilenIndex == -1)
    {
     // Schüler nicht gefunden, überspringe diesen Eintrag
     continue;
    }

    dynamic record = new ExpandoObject();
    
    var valKursart = "";
    var abschnitt = "";
    
    foreach (var prop in dict)
    {
     var key = prop.Key;
     var val = prop.Value.ToString();

     // Es werden nur Veränderungen im interessierenden Abschnitt vorgenommen.
     if(key == "Abschnitt")
      abschnitt = key;

     // Die Spalte "Kursart" & "Abiturfach" wird im interessierenden Abschnitt angepasst.
     if(abschnitt == configuration["Abschnitt"])
      if (key == "Kursart" || key == "Abiturfach")
       if (spaltenIndex < tabelle[zeilenIndex].Count)
         if(key == "Kursart")
         {
          // Der val wird aus Wiki genommen.
          val = tabelle[zeilenIndex][spaltenIndex].Trim();
          valKursart = val;
         }
         if(key == "Abiturfach" && new List<string> { "LK1", "LK2", "AB3", "AB4" }.Contains(valKursart))
         {
          val = valKursart.ToString().Substring(2, 1); // Nur die Ziffer  
         }
     ((IDictionary<string, object>)record)[key] = val;
    }
     zieldatei.Add(record);
    }
  }
  foreach (var aktion in zieldatei.Funktionen)
   aktion(zieldatei);
 }

    static List<List<string>> ParseDokuWikiTable(string wikiText)
    {
        var result = new List<List<string>>();

        // Nur Zeilen mit Tabelle (| oder ^)
        var lines = wikiText
            .Split('\n')
            .Select(l => l.Trim())
            .Where(l => l.StartsWith("^") || l.StartsWith("|"))
            .ToList();

        foreach (var line in lines)
        {
            char sep = line.StartsWith("^") ? '^' : '|';

            // WICHTIG: StringSplitOptions.None => behält leere Zellen bei
            var cells = line
                .Split(new[] { sep }, StringSplitOptions.None)
                .Select(c => c.Trim())
                .ToList();

            // Entferne NUR leere Elemente am Anfang/Ende, falls sie durch äußere Trennzeichen entstehen
            if (cells.Count > 0 && cells.First() == "")
                cells.RemoveAt(0);
            if (cells.Count > 0 && cells.Last() == "")
                cells.RemoveAt(cells.Count - 1);

            result.Add(cells);
        }

        return result;
    }

    internal void Chat(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string zieldateiname,
        Lehrers lehrers,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise = null)
    {
        var zieldatei = new Datei(zieldateiname, configuration, funktionen, delimiter, quote, encoding, shouldAllQuote, lehrers, importhinweise);

        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);
    }

    internal void GetStudentsVonAtlantisCsv(IConfiguration configuration)
    {
        var inputFolder = Path.Combine(configuration["PfadDownloads"], "PDF-Input");

        if (!Directory.Exists(inputFolder))
        {
            Directory.CreateDirectory(inputFolder);
            var path = new TextPath(inputFolder);

            path.RootStyle = new Style(foreground: Spectre.Console.Color.Red);
            path.SeparatorStyle = new Style(foreground: Spectre.Console.Color.SpringGreen2);
            path.StemStyle = new Style(foreground: Spectre.Console.Color.DodgerBlue1);
            path.LeafStyle = new Style(foreground: Spectre.Console.Color.Yellow);

            var panel = new Panel(path)
                .Header("[bold greenYellow] Neuer Ordner für PDF-Dateien: [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Spectre.Console.Color.SpringGreen2);

            AnsiConsole.Write(panel);
        }

        do
        {
            // Wenn eine einzige CSV-Datei vorhanden ist

            if (Directory.GetFiles(inputFolder, "*.csv").Length == 1)
            {
                var csvPath = Directory.GetFiles(inputFolder, "*.csv").FirstOrDefault();

                if (csvPath != null)
                {
                    try
                    {
                        var csvLines = File.ReadAllLines(csvPath, Encoding.UTF8);
                        if (csvLines.Length > 0)
                        {
                            foreach (var line in csvLines.Skip(1)) // Erste Zeile überspringen (Header)
                            {
                                var columns = line.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                                if (columns.Length >= 3)
                                {
                                    try
                                    {
                                        DateTime.ParseExact(columns[2].Trim().Trim('"'), "dd.MM.yyyy", CultureInfo.InvariantCulture);
                                    }
                                    catch (FormatException)
                                    {
                                        var panel = new Panel($"Ungültiges Datum in Zeile: {line}")
                                            .Header("[bold red] Fehler [/]")
                                            .HeaderAlignment(Justify.Left)
                                            .SquareBorder()
                                            .Expand()
                                            .BorderColor(Spectre.Console.Color.Red);
                                        AnsiConsole.Write(panel);
                                        continue; // Diese Zeile überspringen
                                    }

                                    if(columns[0] == "Alhamed")
                                    {
                                        string aaa = "";
                                    }
                                        


                                    var student = new Student
                                    {
                                        Nachname = columns[0].Trim().Trim('"'), // Entfernt führende/trailing Leerzeichen und Anführungszeichen
                                        Vorname = columns[1].Trim().Trim('"'),
                                        Geburtsdatum = DateTime.ParseExact(columns[2].Trim().Trim('"'), "dd.MM.yyyy", CultureInfo.InvariantCulture).ToString("dd.MM.yyyy")
                                    };
                                    Students.Add(student);
                                }
                                else
                                {
                                    var panel = new Panel($"Ungültige Zeile in CSV-Datei: {line}")
                                        .Header("[bold red] Fehler [/]")
                                        .HeaderAlignment(Justify.Left)
                                        .SquareBorder()
                                        .Expand()
                                        .BorderColor(Spectre.Console.Color.Red);
                                    AnsiConsole.Write(panel);
                                }
                            }
                            Global.ZeileSchreiben(csvPath, Students.Count().ToString(), ConsoleColor.Yellow, ConsoleColor.Gray);                            
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Fehler beim Einlesen der CSV-Datei: {ex.Message}");
                    }
                }
                else
                {
                    var panel = new Panel("Die Datei 'schueler.csv' wurde nicht gefunden. Bitte erstellen Sie die Datei im UTF8-Format.")
                            .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Global.ColorFehler);
                    AnsiConsole.Write(panel);
                    throw new Exception($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren![/]");
                }
            }
            else if (Directory.GetFiles(inputFolder, "*.csv").Length == 0)
            {
                var panel = new Panel($"{Path.Combine(inputFolder, "schueler.csv")} existiert nicht. Bitte erstellen Sie die Datei im UTF8-Format.\nFolgende Spalten sind Pflicht: Vorname, Nachname, Geburtsdatum (DD-MM-YYYY), Klasse")
                            .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Global.ColorFehler);
                AnsiConsole.Write(panel);
                throw new Exception($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren![/]");
            }
            else if (Directory.GetFiles(inputFolder, "*.csv").Length > 1)
            {
                var panel = new Panel($"Es gibt mehrere CSV-Dateien in {inputFolder}. Es darf nur eine CSV-Datei vorhanden sein.")
                            .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Global.ColorFehler);
                AnsiConsole.Write(panel);
                throw new Exception($"[grey]  Zuerst die Hinweise [/][bold red]!?[/][grey] bearbeiten, dann hierher zurückkehren![/]");
            }
        }
        while (Students.Count == 0);
    }

    internal void PdfDateienVerarbeiten(IConfiguration configuration)
    {
        List<string> schlüsselwörter = configuration["Schluesselwoerter"].ToString().Trim().Split(",").ToList();

        foreach (string dateiName in Directory.GetFiles(Path.Combine(configuration["PfadDownloads"], "PDF-Input"), "*.*").Where(file => file.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)).OrderBy(file => file))
        {
            try
            {
                var pdfDatei = new PdfDatei(dateiName);
                pdfDatei.Seiten.Einlesen(dateiName);                
                pdfDatei.Art = pdfDatei.GetArt(schlüsselwörter);
                if (string.IsNullOrEmpty(pdfDatei.Art))
                    continue;
                pdfDatei.AnzahlElementeInDieserDatei = pdfDatei.GetAnzahlElementeProDatei(configuration);
                if (pdfDatei.AnzahlElementeInDieserDatei == 0)
                    continue;
                pdfDatei.PdfDateiVerarbeiten(Students, configuration);
            }
            catch (Exception ex)
            {
                Console.WriteLine(dateiName + ": " + ex.Message);
            }       
        }
    }

    internal void NotenlistenAnlegen(IConfiguration configuration,
        Lehrers lehrers,
        string zieldateiname,        
        List<Action<Datei>> funktionen)
    {   
        var leistungsdaten = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
        if (leistungsdaten == null || !leistungsdaten.Any()) return;
        
        var faecher = Quelldateien.GetMatchingList(configuration, "faecher", IStudents, Klassen);
        if (faecher == null || !faecher.Any()) return;

        var zieldatei = new Datei(zieldateiname, funktionen, configuration);

        var beteiligteLul = "";

  // Gib alle Fächer distinct als List<string> zurück, sortiert nach Sortierung S2
  var alleFaecherDistinct = faecher
   .Select(rec => (IDictionary<string, object>)rec)
   .OrderBy(dict => dict["Sortierung S2"].ToString())
   .Select(dict => dict["InternKrz"].ToString()) // Nur Teil vor Leerzeichen
   .Distinct()   
   .ToList();

  var faecherInDieserKlasseAusLeistungsdaten = alleFaecherDistinct
   .Where(fach =>
    leistungsdaten.Any(rec =>
    {
     var dict = (IDictionary<string, object>)rec;
     // Nur Teil vor Leerzeichen vergleichen
     return dict.ContainsKey("Fach") && dict["Fach"].ToString() == fach;
    }))
   .ToList();

  var verschiedeneKlassenDerStudents = IStudents.Select(s => s.Klasse).Distinct().OrderBy(k => k).ToList();

  foreach (var klasse in verschiedeneKlassenDerStudents)
  {
   zieldatei.Name = $"notenlisten:{klasse}-{configuration["InteressierendesSchuljahr"]}-{configuration["Abschnitt"]}";
   zieldatei.Add($"====== {klasse} Notenliste {configuration["InteressierendesSchuljahr"]}/{configuration["Abschnitt"]} ======");
   zieldatei.Add("");

   var kopfzeile = "^  Name  ^ ";
   var kopfzeile2 = "^  ^";

   // Prüfe für alle IStudents dieser Klasse, ob die Initialen (z.B. A.B.) mehrfach vorkommen.
   // Wenn ja, dann muss für diese Klasse die Initialen soweit verlängert werden, bis der Wert eindeutig ist.
   int anzahlZeichen = 1;
   for (int i = 1; i <= 5; i++) // Maximal bis 5 Zeichen
   {
    var initialenGruppiert = IStudents
     .Where(s => s.Klasse == klasse)
     .Select(s => new
     {
      Student = s,
      Initialen = s.Nachname.Substring(0, i) + "." + s.Vorname.Substring(0, i) + "."
     })
     .GroupBy(x => x.Initialen)
     .Where(g => g.Count() > 1)
     .ToList();

    if (initialenGruppiert.Count == 0)
    {
     anzahlZeichen = i;
     break; // Alle Initialen sind eindeutig
    }
   }

    foreach (var fach in faecherInDieserKlasseAusLeistungsdaten)
    {
     // Alle verschiedenen Lehrerkürzel, die das Fach (ohne Ziffer und nur vor dem Leerzeichen) unterrichten
     var alleLehrerDiesesFachs = leistungsdaten
      .Where(rec =>
      {
       var dict = (IDictionary<string, object>)rec;
       return dict["Fach"].ToString() == fach;
      })
      .Select(rec => (IDictionary<string, object>)rec)
      .GroupBy(dict => dict["Fachlehrer"].ToString().Split(' ')[0])
      .Select(g => g.First())
      .Select(dict => new
      {
       Kürzel = dict["Fachlehrer"].ToString().Split(' ')[0]
      })
      .OrderBy(l => l.Kürzel)
      .ToList();

     foreach (var lehrer in alleLehrerDiesesFachs)
     {
      kopfzeile2 += lehrer.Kürzel + " ";
      beteiligteLul += lehrers.Where(l => l.Kürzel == lehrer.Kürzel).Select(l => l.Mail).FirstOrDefault() + ";";
     }
     kopfzeile += fach + "  ^  ";
     kopfzeile2 += "  ^  ";
    }
      
   zieldatei.Add(kopfzeile);
   zieldatei.Add(kopfzeile2);

   foreach (var student in IStudents.Where(s => s.Klasse == klasse).OrderBy(s => s.Nachname).ThenBy(s => s.Vorname))
   {
    var nameKurz = student.Nachname.Substring(0, anzahlZeichen) + "." + student.Vorname.Substring(0, anzahlZeichen) + ".";

    var zeile = "| " + nameKurz;

    // Wenn Leistungsdaten genutzt werden, dann auch die Inhalte in die Zellen übergeben
     foreach (var fach in faecherInDieserKlasseAusLeistungsdaten)
     {
      // Suche die Kursart aus den Leistungsdaten für diesen Schüler
      var note = leistungsdaten
       .Where(rec =>
       {
        var dict = (IDictionary<string, object>)rec;
        return dict["Fach"].ToString() == fach &&
            dict["Nachname"].ToString() == student.Nachname &&
            dict["Vorname"].ToString() == student.Vorname &&
            dict["Geburtsdatum"].ToString() == student.Geburtsdatum && 
            dict["Abschnitt"].ToString() == configuration["Abschnitt"];
       })
       .Select(rec => (IDictionary<string, object>)rec)
       .Select(dict => dict["Note"].ToString())
       .FirstOrDefault();
      
      // Wenn die Note leer ist, dann wird ein Platzhalter eingesetzt, damit die Zelle in der Wiki-Tabelle nicht verschwindet.
      if(note == null)
       note = " ";      
      else if(note == "")
       note = "FIXME";

      zeile += "| " + note;
     }
        
    zeile += "| ";
    zieldatei.Add(zeile);
   }    
   zieldatei.Add("Hier können keine Änderungen vorgenommen werden. Tragen Sie die Noten in Webuntis ein!");

   lehrers.GetTeamsUrl(beteiligteLul.Split(';'), String.Join(';', klasse));            
   
   foreach (var aktion in zieldatei.Funktionen)
    aktion(zieldatei);}
 }

 internal void SvwsServerStarten(IConfiguration configuration)
 {
  throw new NotImplementedException();
 }

 internal void SvwsServerStoppen(IConfiguration configuration)
 {
  throw new NotImplementedException();
 }

 internal void SvwsServerBackupErstellen(IConfiguration configuration)
 {
  throw new NotImplementedException();
 }

 internal void SvwsServerStatusAnzeigen(IConfiguration configuration)
 {
  throw new NotImplementedException();
 }


 /*internal Datei? Leistungsdaten(IConfiguration configuration, string zieldateiname, Unterrichte kurse, Global.Zweck art)
 {
     var zieldatei = new Datei(zieldateiname);

     List<dynamic> marksPerLs = new List<dynamic>();

     List<dynamic> schLeistus = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
     if ((schLeistus == null || schLeistus.Count == 0)) return [];

     if (Global.Zweck.Statistik != art)
     {
         marksPerLs = Quelldateien.GetMatchingList(configuration, "marksperlesson", IStudents, Klassen);
         if (marksPerLs == null || marksPerLs.Count == 0) return [];
     }

     var records = new List<dynamic>();

     if (art == Global.Zweck.Mahnung)
     {
         marksPerLs = marksPerLs.Where(rec =>
         {
             var dict = (IDictionary<string, object>)rec;
             return dict["Prüfungsart"].ToString().Contains("Mahnung");
         }).ToList();

         // Reduziere die IStudents-Liste basierend auf den gefilterten marksPerLs
         var x = IStudents.Where(x => x.Status == "2").Where(student =>
             marksPerLs.Any(mark =>
             {
                 var dict = (IDictionary<string, object>)mark;
                 return dict["Name"].ToString().Contains(student.Vorname) &&
                     dict["Name"].ToString().Contains(student.Nachname) &&
                     dict["Klasse"].ToString() == student.Klasse;
             })
         ).ToList();

         IStudents.Clear();
         IStudents.AddRange(x);
     }

     AnsiConsole.Status()
         .Spinner(Spinner.Known.Dots)
         .Start("SchuelerLeistungsdaten.dat ...", ctx =>
     {
         foreach (var klasse in IStudents.OrderBy(x => x.Klasse).Select(x => x.Klasse).Distinct())
         {
             var isFirstRun = true;

             var verschiedeneFaecherDerKlasse = VerschiedeneFaecher(klasse, expLessons);

             var religionWurdeUnterrichtet = verschiedeneFaecherDerKlasse
                 .Any(fach => new List<string>() { "rel", "kr", "er", "reli" }.Contains(fach.ToLower()));



             foreach (var student in IStudents.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname).Where(x => x.Klasse == klasse))
             {
                 var istReliabmelder = schBasisds.Any(rec =>
                 {
                     var dict = (IDictionary<string, object>)rec;
                     return dict["Nachname"].ToString() == student.Nachname
                         && dict["Vorname"].ToString() == student.Vorname
                         && dict["Geburtsdatum"].ToString() == student.Geburtsdatum
                         && !string.IsNullOrEmpty(dict["Abmeldedatum Religionsunterricht"].ToString());
                 });

                 foreach (var fach in verschiedeneFaecherDerKlasse)
                 {
                     // Normalerweise gibt es nur einen Unterricht. 
                     var unterrichteMitDiesemFach = GetUnterrichteMitDiesemFach(fach, klasse, expLessons);

                     var dictExp = (IDictionary<string, object>)unterrichteMitDiesemFach[0];

                     var zusatzlehrkraft = "";
                     var zusatzlehrkraftWochenstunden = "";

                     // In der Statistikzählen allen Fächer mit, auch wenn sie nicht relevant sind.
                     if (art != Global.Zweck.Statistik)
                     {
                         if (!student.UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(dictExp, configuration)) continue;
                     }

                     // Wenn dieses Fach mit diesem Lehrer bereits in den records existiert,
                     // dann wird es nicht erneut hinzugefügt.

                     var gibtDasFachMitDemLehrerSchon = records.Any(rec =>
                     {
                         var dict = (IDictionary<string, object>)rec;
                         return dict["Fach"].ToString() == dictExp["subject"].ToString() &&
                             dict["Fachlehrer"].ToString() == dictExp["teacher"].ToString() &&
                             dict["Vorname"].ToString() == student.Vorname &&
                             dict["Nachname"].ToString() == student.Nachname &&
                             dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                     });

                     if (!gibtDasFachMitDemLehrerSchon)
                     {
                         string jahrgang = student.GetJahrgang(schBasisds);
                         string note = student.GetNote(jahrgang, marksPerLs, dictExp["subject"].ToString()!, art);

                         // Wenn Reli unterricht wurde und der Schüler abgemeldet ist, dann wird NT eingesetzt.
                         if (
                             new List<string>() { "rel", "kr", "er", "reli" }.Contains(fach.ToLower())
                             && religionWurdeUnterrichtet)
                         {
                             if (istReliabmelder)
                             {
                                 note = "NT";
                             }
                         }

                         string kursart = GetKursart(configuration, jahrgang, fach);
                         bool mahnung = student.GetMahnung(marksPerLs, dictExp["subject"].ToString()!);

                         // Die Kursart 
                         var kursartBisher = schLeistus
                             .Where(record =>
                             {
                                 var dict = (IDictionary<string, object>)record;
                                 return dict["Vorname"].ToString() == student.Vorname &&
                                     dict["Nachname"].ToString() == student.Nachname &&
                                     dict["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                                     dictExp["subject"] != null &&
                                     dict["Fach"].ToString() == dictExp["subject"].ToString();
                             })
                             .Select(record =>
                             {
                                 var dict = (IDictionary<string, object>)record;
                                 return dict["Kursart"].ToString();
                             })
                             .FirstOrDefault()
                             ?.ToString();

                         if (!string.IsNullOrEmpty(kursartBisher))
                             kursart = kursartBisher;

                         // Klassenunterrichte und Religion wird immer hinzugefügt
                         if (dictExp["studentgroup"].ToString() == "" || new List<string>() { "rel", "kr", "er", "reli", "religion", "rel1" }.Contains(fach.ToLower()))
                         {
                             dynamic record = new ExpandoObject();
                             record.Nachname = $"{student.Nachname}#{klasse}";
                             record.Vorname = student.Vorname;
                             record.Geburtsdatum = student.Geburtsdatum;
                             record.Jahr = Global.AktSj[0];
                             record.Abschnitt = configuration["Abschnitt"];
                             record.Fach = dictExp["subject"].ToString();
                             record.Fachlehrer = dictExp["teacher"].ToString();
                             record.Kursart = kursart;
                             record.Kurs = "";
                             record.Note = art == Global.Zweck.Statistik ? "" : note;
                             record.Abiturfach = "";
                             record.WochenstdPUNKT = dictExp["periods"];
                             record.ExterneLEERZEICHENSchulnrPUNKT = "";
                             record.Zusatzkraft = zusatzlehrkraft;
                             record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                             record.Jahrgang = "";
                             record.Jahrgänge = "";
                             record.FehlstdPUNKT = ""; // Fehlzeiten werden über die Abschnittsdaten importiert.
                             record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                             if (art == Global.Zweck.Mahnung)
                             {
                                 record.Mahnung = "J";
                                 record.Sortierung = "";
                                 record.Mahndatum = "";//DateTime.Now.ToShortDateString();
                             }
                             if ((mahnung && art == Global.Zweck.Mahnung) || art != Global.Zweck.Mahnung)
                             {
                                 records.Add(record);
                             }
                         }
                         else // Bei Kursunterrichten wird geschaut, ob der Schüler den Kurs belegt hat. 
                         {
                             var id = student.Id;
                             var studentZeile = stdgroupSs
                                 .Where(record =>
                                 {
                                     var dict = (IDictionary<string, object>)record;
                                     return dict["studentId"].ToString() == id &&
                                         dict["studentgroup.name"].ToString() ==
                                         dictExp["studentgroup"].ToString();
                                 })
                                 .FirstOrDefault();
                             var dictStudentgroup = (IDictionary<string, object>)studentZeile!;

                             if (dictStudentgroup != null)
                             {
                                 if (!student.UnterrichtIstRelevantFürZeugnisInDiesemAbschnitt(dictStudentgroup, configuration))
                                     continue;
                                 dynamic record = new ExpandoObject();
                                 record.Nachname = $"{student.Nachname}#{klasse}";
                                 record.Vorname = student.Vorname;
                                 record.Geburtsdatum = student.Geburtsdatum;
                                 record.Jahr = Global.AktSj[0];
                                 record.Abschnitt = configuration["Abschnitt"];
                                 record.Fach = dictStudentgroup["subject"].ToString();
                                 record.Fachlehrer = dictExp["teacher"].ToString();
                                 record.Kursart = kursart;
                                 record.Kurs = dictStudentgroup["studentgroup.name"].ToString()!.Substring(0,
                                     Math.Min(dictStudentgroup["studentgroup.name"].ToString()!.Length, 20));
                                 record.Note = note;
                                 record.Abiturfach = "";
                                 record.WochenstdPUNKT = dictExp["periods"];
                                 record.ExterneLEERZEICHENSchulnrPUNKT = "";
                                 record.Zusatzkraft = zusatzlehrkraft;
                                 record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                                 record.Jahrgang = student.Jahrgang;
                                 record.Jahrgänge = "";
                                 record.FehlstdPUNKT = "";
                                 record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                                 if (art == Global.Zweck.Mahnung)
                                 {
                                     record.Mahnung = "";
                                     record.Sortierung = "";
                                     record.Mahndatum = "";//DateTime.Now.ToShortDateString();
                                 }
                                 if ((mahnung && art == Global.Zweck.Mahnung) || art != Global.Zweck.Mahnung)
                                 {
                                     records.Add(record);
                                 }
                             }
                         }
                     }
                 }
             }
         }
         Global.ZeileSchreiben("SchuelerLeistungsdaten.dat", records.Count().ToString());
     });

     zieldatei.AddRange(records);
     return zieldatei;
 }*/

 // Comparer for (string? Vorname, string? Nachname) tuples
 private class TupleStringComparer : IEqualityComparer<(string? Vorname, string? Nachname)>
    {
        public bool Equals((string? Vorname, string? Nachname) x, (string? Vorname, string? Nachname) y)
        {
            return string.Equals(x.Vorname, y.Vorname, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.Nachname, y.Nachname, StringComparison.OrdinalIgnoreCase);
        }

        public int GetHashCode((string? Vorname, string? Nachname) obj)
        {
            int hashVorname = obj.Vorname?.ToLowerInvariant().GetHashCode() ?? 0;
            int hashNachname = obj.Nachname?.ToLowerInvariant().GetHashCode() ?? 0;
            return hashVorname ^ hashNachname;
        }
    }
}

public class CountColumn : ProgressColumn
{
    public override IRenderable Render(RenderOptions options, ProgressTask task, TimeSpan deltaTime)
    {
        return new Markup($"[yellow]{task.Value:N0}/{task.MaxValue:N0}[/]");
    }
}

[Serializable]
public class RestartException : Exception
{
    public RestartException()
    {
    }

    public RestartException(string? message) : base(message)
    {
    }

    public RestartException(string? message, Exception? innerException) : base(message, innerException)
    {
    }
}