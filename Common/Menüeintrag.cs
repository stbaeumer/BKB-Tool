using System.Dynamic;
using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using CookComputing.XmlRpc;
//using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using Color = Spectre.Console.Color;

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


public class Menüeintrag
{
    public Gruppen Gruppen { get; set; }
    public List<string> Beschreibung { get; set; }
    public Dateien Quelldateien { get; set; }
    public Anrechnungen Anrechnungen { get; set; }
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

    public Menüeintrag(string titel, Anrechnungen anrechnungen, Dateien quelldateien, Students students, Klassen klassen, List<string> beschreibung, Action<Menüeintrag> funktion, Global.Rubrik rubrik = Global.Rubrik.Allgemein, Global.NurBeiDiesenSchulnummern nurbeiDiesenSchulnummern = Global.NurBeiDiesenSchulnummern.Alle)
    {
        if (titel == null) throw new ArgumentNullException(nameof(titel));
        if (anrechnungen == null) throw new ArgumentNullException(nameof(anrechnungen));
        if (quelldateien == null) throw new ArgumentNullException(nameof(quelldateien));
        if (students == null) throw new ArgumentNullException(nameof(students));
        if (klassen == null) throw new ArgumentNullException(nameof(klassen));
        if (beschreibung == null) throw new ArgumentNullException(nameof(beschreibung));
        {            
            Titel = titel;
            Anrechnungen = anrechnungen;
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

        var quit = Global.Konfig("MaxDateiAlter", Global.Modus.Update, configuration, "Maximales Alter der eingelesenen Dateien", "", Global.Datentyp.Int);

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
    public void FilterInteressierendeStudentsUndKlassen(IConfiguration configuration)
    {
        var interessierendeStudents = new Students();

        Global.Konfig("Klassen", Global.Modus.Update, configuration, "Interessierende Klasse(n)", $"Geben Sie die interessierende(n) Klasse(n) an. Mehrere Klassen sind mit Komma zu trennen. Es können auch Namensteile von Klassen angegeben werden, wordurch alle Klassen gewählt werden, deren Klassenname den Namensteil enthält. Alle Klassen werden mit dem Wort [bold springGreen2]alle[/] gewählt.", Global.Datentyp.Klassen, "", this.Students);

        var interessierendeKlassen = configuration["Klassen"].ToString().Split(",").ToList();

        IStudents = new Students();
        IStudents.AddRange(from t in Students where interessierendeKlassen.Contains(t.Klasse) select t);
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

                    var kursart = GetKursart(jahrgang, fach);
                    var note = dict["Field8"].ToString();
                    var tendenz = dict["Field10"].ToString();
                    var punkte = dict["Field11"].ToString();
                    var noteOderPunkte = dict["Field13"].ToString();

                    dynamic record = new ExpandoObject();
                    record.Nachname = student.Nachname;
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

    private static string GetKursart(string jahrgang, string fach)
    {
        if (!jahrgang.StartsWith("GY")) return "PUK";
        if (!jahrgang.EndsWith("02") && !jahrgang.EndsWith("03") && !jahrgang.EndsWith("12") &&
            !jahrgang.EndsWith("13")) return "PUK";
        if (!fach.Contains(" L")) return fach.Contains(" G") ? "GKS" : "PUK";
        var linkerTeil = fach.Split(' ')[0].TrimEnd();
        return new List<string>() { "D", "M", "E", "BI" }.Contains(linkerTeil) ? "LK1" : "LK2";
    }

    public Datei Lernabschnittsdaten(IConfiguration configuration, Global.Art art, string zieldateiname)
    {
        var schuelerLernab = Quelldateien.GetMatchingList(configuration, "lernabschnittsdat", IStudents, Klassen);
        if (schuelerLernab == null || !schuelerLernab.Any()) return [];

        var schuelerBasisd = Quelldateien.GetMatchingList(configuration, "schuelerbasisdate", IStudents, Klassen);
        if (schuelerBasisd == null || !schuelerBasisd.Any()) return [];

        var absencePerStud = new List<dynamic>();

        var konferenzdatum = DateTime.Now;
        var zeugnisdatum = DateTime.Now;

        var zielDatei = new Datei(zieldateiname, new Datei(schuelerLernab));

        if (art != Global.Art.Statistik)
        {
            absencePerStud = Quelldateien.GetMatchingList(configuration, "absenceperstudent", IStudents, Klassen);
            if (absencePerStud == null || !absencePerStud.Any()) return [];

            configuration = Global.Konfig("Abschnitt", Global.Modus.Update, configuration, "Abschnitt", $"Geben Sie den Lernabschnitt an. Das Schuljahr beginnt immer mit Abschnitt [bold aqua]1[/]. I.d.R. wechselt der Abschnitt im Halbjahr auf Abschnitt [bold aqua]2[/]", Global.Datentyp.Abschnitt);

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
            
            configuration = Global.Konfig($"{konferenzart}konferenzdatum", Global.Modus.Update, configuration, $"{konferenzart}konferenzdatum", $"Geben Sie das {konferenzart}konferenzdatum an. Das kann später in SchILD (mit einem Gruppenprozess) erneut geändert werden.", Global.Datentyp.DateTime);
            konferenzdatum = DateTime.Parse(configuration[$"{konferenzart}konferenzdatum"]);
            configuration = Global.Konfig($"{konferenzart}zeugnisdatum", Global.Modus.Update, configuration, $"{konferenzart}zeugnisdatum", $"Geben Sie das {konferenzart}zeugnisdatum an. Das kann später in SchILD (mit einem Gruppenprozess) erneut geändert werden.", Global.Datentyp.DateTime);
            zeugnisdatum = DateTime.Parse(configuration[$"{konferenzart}zeugnisdatum"]);
            
            configuration = Global.Konfig("MaximaleAnzahlFehlstundenProTag", Global.Modus.Update, configuration, "Maximale Anzahl Fehlstunden pro Tag", "Geben Sie die maximale Anzahl zählender Fehlstunden pro Tag an. Wenn der Unterricht spätestens nach 8 Stunden endet, ist 8 ein guter Wert. Sollte die Anzahl der Fehlstunden in Webuntis diesen Wert übersteigen, dann deutet das auf Fehlzeiten im Praktikum oder ein Fehlen bei einer ganztägigen Veranstaltung hin. Es werden keine Fehlstunden an diesem Tag auf dem Zeugnis berücksichtigt.", Global.Datentyp.Int);
            configuration = Global.Konfig("FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt", Global.Modus.Update, configuration, "Unberücksichtigte Fehltage", "Geben Sie die Anzahl Tage vor der Zeugniskonferenz an, an denen Fehlzeiten unberücksichtigt bleiben. Wenn dieser Wert z.B. auf [bold springGreen2]3[/] gesetzt wird, wird verhindert, dass Schüler*innen eine Entschuldigung zwar unverzüglich, aber gleichzeitig erst nach der Zeugniskonferenz einreichen. Alternativ kann man den Wert auf 0 setzen und den Zeitraum des Exports der [bold springGreen2]AbsencePerStudent[/] entsprechend einschränken.", Global.Datentyp.Int);
        }

        var records = new List<dynamic>();

        try
        {
            foreach (var student in IStudents)
            {
                var halbjahreszeugnisdatum = DateTime.Parse(configuration["Halbjahreszeugnisdatum"]);
                var halbjahreswechelInderZukunft = halbjahreszeugnisdatum > DateTime.Now;

                var dictBasisdaten = schuelerBasisd
                    .Where(recBasis =>
                    {
                        var dictBasis = (IDictionary<string, object>)recBasis;
                        return dictBasis["Nachname"].ToString() == student.Nachname &&
                               dictBasis["Vorname"].ToString() == student.Vorname &&
                               dictBasis["Geburtsdatum"].ToString() == student.Geburtsdatum &&
                               dictBasis["Jahr"].ToString() == Global.AktSj[0] &&
                               ((art == Global.Art.Statistik || halbjahreswechelInderZukunft) ? true : dictBasis["Abschnitt"].ToString() == configuration["Abschnitt"]);
                    }).FirstOrDefault() as IDictionary<string, object>;

                if (dictBasisdaten != null)
                {
                    var versetzung = "";
                    var abschluss = "";
                    var klassenlehrer = Klassen.Where(rec => rec.Name == student.Klasse)
                        .Select(rec => rec.Klassenleitungen[0].Kürzel).FirstOrDefault();
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
                    record.Nachname = student.Nachname;
                    record.Vorname = student.Vorname;
                    record.Geburtsdatum = student.Geburtsdatum;
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = art == Global.Art.Statistik ? "1" : configuration["Abschnitt"];
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
                    record.Konferenzdatum = art == Global.Art.Statistik ? "" : konferenzdatum.ToShortDateString();
                    record.Zeugnisdatum = art == Global.Art.Statistik ? "" : zeugnisdatum.ToShortDateString();
                    record.SummeFehlstd = art == Global.Art.Statistik ? "" : student.GetFehlstd(absencePerStud, configuration);
                    record.SummeFehlstdUNTERSTRICHunentschuldigt = art == Global.Art.Statistik ? "" : student.GetUnentFehlstd(absencePerStud, configuration);
                    record.allgPUNKTMINUSbildenderLEERZEICHENAbschluss = "";
                    record.berufsbezPUNKTLEERZEICHENAbschluss = "";
                    record.Zeugnisart = zeugnisart;
                    record.FehlstundenMINUSGrenzwert = "";
                    record.DatumLEERZEICHENvon = "";
                    record.DatumLEERZEICHENbis = "";
                    records.Add(record);
                }
            }

            zielDatei.AddRange(records);
            return zielDatei;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Console.ReadKey();
            return null;
        }
    }

    public Datei Leistungsdaten(IConfiguration configuration, string zieldateiname, Global.Art art)
    {
        var zieldatei = new Datei(zieldateiname);

        List<dynamic> marksPerLs = new List<dynamic>();

        var expLessons = Quelldateien.GetMatchingList(configuration, "exportlessons", IStudents, Klassen);
        if (expLessons == null || expLessons.Count == 0) return [];

        var stdgroupSs = Quelldateien.GetMatchingList(configuration, "studentgroupstudents", IStudents, Klassen);
        if (art == null && (stdgroupSs == null || stdgroupSs.Count == 0)) return [];

        List<dynamic> schBasisds = Quelldateien.GetMatchingList(configuration, "schuelerbasisdaten", IStudents, Klassen);
        if (art == null && (schBasisds == null || schBasisds.Count == 0)) return [];

        List<dynamic> schLeistus = Quelldateien.GetMatchingList(configuration, "schuelerleistungsdaten", IStudents, Klassen);
        if (art == null && (schLeistus == null || schLeistus.Count == 0)) return [];

        if (Global.Art.Statistik != art)
        {
            marksPerLs = Quelldateien.GetMatchingList(configuration, "marksperlesson", IStudents, Klassen);
            if (marksPerLs == null || marksPerLs.Count == 0) return [];
        }
        else
        {
            configuration["Abschnitt"] = "1";
        }

        var records = new List<dynamic>();

        if (art == Global.Art.Mahnung)
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

        foreach (var klasse in IStudents.OrderBy(x => x.Klasse).Select(x => x.Klasse).Distinct())
        {
            var isFirstRun = true;

            var verschiedeneFaecherDerKlasse = VerschiedeneFaecher(klasse, expLessons);

            var religionWurdeUnterrichtet = verschiedeneFaecherDerKlasse
                .Any(fach => new List<string>() { "rel", "kr", "er", "reli" }.Contains(fach.ToLower()));

            foreach (var student in IStudents.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname).Where(x => x.Klasse == klasse))
            {
                if(student.Nachname == "Alp" && student.Vorname.StartsWith("U"))
                {
                    string a = "";
                }

                var istReliabmelder = schBasisds.Any(rec =>
                {
                    var dict = (IDictionary<string, object>)rec;
                    return dict["Nachname"].ToString() == student.Nachname
                        && dict["Vorname"].ToString() == student.Vorname
                        && dict["Geburtsdatum"].ToString() == student.Geburtsdatum
                        && !string.IsNullOrEmpty(dict["Abmeldedatum Religionsunterricht"].ToString())
                    //&& (DateTime.Parse(dict["AbmeldedatumLEERZEICHENReligionsunterricht"].ToString()) < DateTime.Parse(dict["AnmeldedatumLEERZEICHENReligionsunterricht"].ToString()))
                    ;
                });

                foreach (var fach in verschiedeneFaecherDerKlasse)
                {
                    // Normalerweise gibt es nur einen Unterricht. 
                    var unterrichteMitDiesemFach = GetUnterrichteMitDiesemFach(fach, klasse, expLessons);

                    var dictExp = (IDictionary<string, object>)unterrichteMitDiesemFach[0];

                    var zusatzlehrkraft = "";
                    var zusatzlehrkraftWochenstunden = "";

                    // In der Statistikzählen allen Fächer mit, auch wenn sie nicht relevant sind.
                    if (art != Global.Art.Statistik)
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

                        string kursart = GetKursart(jahrgang, fach);
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
                            record.Nachname = student.Nachname;
                            record.Vorname = student.Vorname;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Jahr = Global.AktSj[0];
                            record.Abschnitt = configuration["Abschnitt"];
                            record.Fach = dictExp["subject"].ToString();
                            record.Fachlehrer = dictExp["teacher"].ToString();
                            record.Kursart = kursart;
                            record.Kurs = "";
                            record.Note = art == Global.Art.Statistik ? "" : note;
                            record.Abiturfach = "";
                            record.WochenstdPUNKT = dictExp["periods"];
                            record.ExterneLEERZEICHENSchulnrPUNKT = "";
                            record.Zusatzkraft = zusatzlehrkraft;
                            record.WochenstdPUNKTLEERZEICHENZK = zusatzlehrkraftWochenstunden;
                            record.Jahrgang = "";
                            record.Jahrgänge = "";
                            record.FehlstdPUNKT = ""; // Fehlzeiten werden über die Abschnittsdaten importiert.
                            record.unentschPUNKTLEERZEICHENFehlstdPUNKT = "";
                            if (art == Global.Art.Mahnung)
                            {
                                record.Mahnung = "J";
                                record.Sortierung = "";
                                record.Mahndatum = "";//DateTime.Now.ToShortDateString();
                            }
                            if ((mahnung && art == Global.Art.Mahnung) || art != Global.Art.Mahnung)
                            {
                                records.Add(record);
                            }
                        }
                        else // Bei Kursunterrichten wird geschaut, ob der Schüler den Kurs belegt hat. 
                        {
                            var id = string.IsNullOrEmpty(student.ExterneIdNummer)
                                ? student.IdSchild
                                : student.ExterneIdNummer;
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
                                record.Nachname = student.Nachname;
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
                                if (art == Global.Art.Mahnung)
                                {
                                    record.Mahnung = "";
                                    record.Sortierung = "";
                                    record.Mahndatum = "";//DateTime.Now.ToShortDateString();
                                }
                                if ((mahnung && art == Global.Art.Mahnung) || art != Global.Art.Mahnung)
                                {
                                    records.Add(record);
                                }
                            }
                        }
                    }
                }
            }
        }

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

    public Datei Kurse(IConfiguration configuration, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        var records = new List<dynamic>();

        var exportLessons = Quelldateien.GetMatchingList(configuration, "exportlesson", IStudents, Klassen);
        if (exportLessons == null || exportLessons.Count == 0) return [];

        var klassen = Quelldateien.GetMatchingList(configuration, "Klassen", IStudents, Klassen);
        if (klassen == null || klassen.Count == 0) return [];

        foreach (var recExp in exportLessons)
        {
            var dictExp = (IDictionary<string, object>)recExp;

            if (string.IsNullOrEmpty(dictExp["studentgroup"].ToString())) continue;
            foreach (var klasse in dictExp["klassen"].ToString()!.Split('~'))
            {
                var zeileKlasse = klassen.Where(record =>
                {
                    var dict = (IDictionary<string, object>)record;
                    return dict["InternBez"].ToString() == klasse;
                }).FirstOrDefault();

                if (!IStudents.Select(x => x.Klasse).ToList().Contains(klasse)) continue;
                {
                    dynamic record = new ExpandoObject();
                    record.KursBez = dictExp["studentgroup"].ToString()!.Substring(0,
                        Math.Min(dictExp["studentgroup"].ToString()!.Length, 20));
                    record.Klasse = klasse;
                    record.Jahr = Global.AktSj[0];
                    record.Abschnitt = configuration["Abschnitt"];
                    record.Jahrgang = zeileKlasse?.Jahrgang;
                    record.Fach = Global.PrüfeAufNullOrEmpty(dictExp, "subject");
                    record.Kursart = GetKursart(record.Jahrgang, dictExp["subject"].ToString());
                    record.WochenstdPUNKT = dictExp["periods"];
                    record.WochenstdPUNKTLEERZEICHENKL = "";
                    record.Kursleiter = dictExp["teacher"];
                    record.Epochenunterricht = "";
                    record.Schulnr = "177659";
                    record.WochenstdPUNKTLEERZEICHENZK = "";
                    record.Zusatzkraft = "";
                    record.WochenstdLEERZEICHENZK = "";
                    record.WeitereLEERZEICHENZusatzkraft = "";
                    zieldatei.AddRange(record);
                }
            }
        }

        return zieldatei;
    }

    static bool IsValidHttpUrl(string url)
    {
        return Uri.TryCreate(url, UriKind.Absolute, out Uri? uriResult) &&
               (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
    }


    public Datei? Faecher(IConfiguration configuration, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);

        var records = new List<dynamic>();
        var exportLessons = Quelldateien.GetMatchingList(configuration, "exportlessons", IStudents, Klassen);
        if (!exportLessons.Any())
        {
            return [];
        }

        var schildFaecher = Quelldateien.GetMatchingList(configuration, "faecher", IStudents, Klassen);
        if (!schildFaecher.Any())
        {
            return [];
        }

        foreach (var recExp in exportLessons)
        {
            var dictExp = (IDictionary<string, object>)recExp;

            var schildFach = schildFaecher.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["InternKrz"].ToString() == dictExp["subject"].ToString();
            }).FirstOrDefault();

            // Wenn es das Fach in SchILD nicht gibt, ...

            if (schildFach != null) continue;
            {
                // ... wird bei Fächern mit Suffix geprüft, ob es bereits ein Schildfach ohne Suffix gibt.

                var subject = dictExp["subject"].ToString();
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
                        if (records.Any(x => x.InternKrz == dictExp["subject"].ToString())) continue;
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

                        var gibtEsSchon = records.Any(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return dict["InternKrz"].ToString() == subject;
                        });

                        if (!gibtEsSchon)
                        {
                            records.Add(record);
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

                        var gibtEsSchon = records.Any(rec =>
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

                    var gibtEsSchon = records.Any(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["InternKrz"].ToString() == subject;
                    });

                    if (!gibtEsSchon)
                    {
                        records.Add(record);
                    }
                }
            }
        }

        return zieldatei;
    }

    public void LuLAnEintragungDerZeugnisnotenErinnern(IConfiguration configuration, Lehrers lehrers)
    {
        var leistungsdaten = Quelldateien.GetMatchingList(configuration, "leistungsdaten", IStudents, Klassen);
        var betreff = "";
        var adressen = "";
        var anrede = "";
        var lul = new List<string?>();
        var eintaege = new List<string>();

        int i = 1;
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

                if (!lul.Contains(dict["Fachlehrer"].ToString()))
                {
                    lul.Add(dict["Fachlehrer"].ToString());
                }
            }
        }

        if (eintaege.Count > 0)
        {
            Global.DisplayHeader(configuration);
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

            Global.DisplayHeader(configuration);
            Console.WriteLine("   " + adressen.TrimEnd(','));
        }
        else
        {
            Console.WriteLine("  Es fehlen keine Noten. Gut so.");
        }
    }

    public void ChatErzeugen(IConfiguration configuration, String mitgliederMail)
    {
        var lehrers = new Lehrers();
        lehrers.GetTeamsUrl(mitgliederMail.Split(';'), String.Join(';', IKlassen));
    }

    public Datei? WebuntisOderNetmanOderLitteraCsv(IConfiguration configuration, string zieldateiname)
    {
        try
        {
            var zieldatei = new Datei(zieldateiname);

            List<dynamic>? webuntisStudents = Quelldateien.GetMatchingList(configuration, "student_", Students, Klassen);
            if (webuntisStudents == null || webuntisStudents.Count == 0) return [];
            var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", Students, Klassen);
            if (schuelerZusatzdaten == null || schuelerZusatzdaten.Count == 0) return [];
            var schuelerLernabschnittsdaten = Quelldateien.GetMatchingList(configuration, "schuelerlernabschnittsdaten", Students, Klassen);
            if (schuelerLernabschnittsdaten == null || schuelerLernabschnittsdaten.Count == 0) return [];            
            var schuelerErzieher = Quelldateien.GetMatchingList(configuration, "schuelererzieher", Students, Klassen);
            if (schuelerErzieher == null || schuelerErzieher.Count == 0) return [];
            var schuelerAdressen = Quelldateien.GetMatchingList(configuration, "schueleradressen", Students, Klassen);
            if (schuelerAdressen == null || schuelerAdressen.Count == 0) return [];
            var lehrkraefte = Quelldateien.GetMatchingList(configuration, "lehrkraefte", Students, Klassen);
            if (lehrkraefte == null || lehrkraefte.Count == 0) return [];
            var klassen = Quelldateien.GetMatchingList(configuration, "klassen", Students, Klassen);
            if (klassen == null || klassen.Count == 0) return [];
            
            AnsiConsole.Write(new Rule("[aqua]Anstehende Änderungen/Neuanlagen:[/]").RuleStyle("aqua").LeftJustified());

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

            foreach (var rec in webuntisStudents)
            {
                if (rec is not IDictionary<string, object> webuntisStudent) continue;

                var schildStudent = Students
                .Where(x =>
                    x.Nachname == webuntisStudent["longName"].ToString() &&
                    x.Vorname == webuntisStudent["foreName"].ToString() &&
                    x.Geburtsdatum == webuntisStudent["birthDate"].ToString())
                .OrderByDescending(x => int.TryParse(x.IdSchild, out var id) ? id : 0) // IdSchild in int umwandeln, Standardwert 0 bei Fehler
                .FirstOrDefault(); // Gibt den Schüler mit der höchsten IdSchild zurück

                schildStudent.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);

                if (schildStudent.Nachname == "Hermann" && schildStudent.Vorname.StartsWith("Rebecca"))
                {
                    string a = "a";
                }
                if (schildStudent == null) continue;

                var id = string.IsNullOrEmpty(schildStudent.ExterneIdNummer)
                    ? schildStudent.IdSchild
                    : schildStudent.ExterneIdNummer;

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

                                if(schildStudent.ZeugnisdatumLetztesZeugnisInDieserKlasse >= DateTime.Now)
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
                        }
                    }
                }
            }

            // Ab hier die Neuanlagen
            // Damit Schüler nicht doppelt angelegt werden, wir zuerst

            var uniqueStudents = Students
                .DistinctBy(s => new { s.Vorname, s.Nachname, s.Geburtsdatum })
                .OrderBy(s => s.Klasse)
                .ThenBy(s => s.Nachname)
                .ThenBy(s => s.Vorname);

            foreach (var studen in uniqueStudents)
            {
                if (studen.Nachname == "Hermann" && studen.Vorname.StartsWith("R"))
                {
                    string a = "a";
                }

                // Es kann sein, dass Schüler nach Abschluss als Gast bleiben. 
                // Es wird angenommen, dass der letzte in der Importliste der aktuelle ist.
                var student = Students.OrderByDescending(x => int.TryParse(x.IdSchild, out var id) ? id : 0)
                .FirstOrDefault(x => x.Nachname == studen.Nachname && x.Vorname == studen.Vorname && x.Geburtsdatum == studen.Geburtsdatum);

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
                        var id = string.IsNullOrEmpty(student.ExterneIdNummer) ? student.IdSchild : student.ExterneIdNummer;
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

                /*var sb = schildschuelerexport
                    ?.Where(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["Nachname"].ToString() == student.Nachname &&
                            dict["Vorname"].ToString() == student.Vorname &&
                            dict["Geburtsdatum"].ToString() == student.Geburtsdatum;
                    }).LastOrDefault() as IDictionary<string, object>;*/

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

                if (Path.GetFileName(zieldateiname).ToLower().Contains("webuntis"))
                {
                    record.Schlüssel = string.IsNullOrEmpty(student.ExterneIdNummer)
                        ? student.IdSchild
                        : student.ExterneIdNummer.ToString().PadLeft(6, '0');

                    if (studen.Nachname == "Hermann" && studen.Vorname.StartsWith("R"))
                    {
                        string a = "a";
                    }

                    record.EMINUSMail = student.MailSchulisch;
                    if (!student.MailSchulisch.Contains(record.Schlüssel))
                    {
                        string a = "";
                    }
                    record.Familienname = student.Nachname;
                    record.Vorname = student.Vorname;
                    record.Klasse = student.Klasse;
                    record.Kurzname = student.MailSchulisch.Split('@')[0];
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
                    record.O365Identität = student.MailSchulisch;
                    record.Benutzername = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");

                    // Es werden nur diejenigen Schüler exportiert die aktiv oder Gast sind, 
                    // und alle anderen, deren Entlassdatum heute oder in den letzten sechs Wochen war.                                         
                    if (student.Status == "2" || student.Status == "6" || student.JüngstEntlassen())
                    {
                        zieldatei.Add(record);
                    }
                }
                else if (Path.GetFileName(zieldateiname).ToLower().Contains("netman"))
                {
                    // Netman
                    // ed123456	Dagobert	Eggemann	ed123456@students.berufskolleg-borken.de	E01.07.1992	BZ22A	Stappert, Markus
                    record.Schlüssel = string.IsNullOrEmpty(student.ExterneIdNummer)
                        ? student.IdSchild
                        : student.ExterneIdNummer.ToString().PadLeft(6, '0');
                    record.Kurzname = student.MailSchulisch.Replace("@students.berufskolleg-borken.de", "");
                    record.Vorname = student.Vorname;
                    record.Nachname = student.Nachname;
                    record.Mail = student.MailSchulisch;
                    record.Passwort = student.Nachname.Substring(0, 1).ToUpper() + student.Geburtsdatum;
                    record.Klasse = student.Klasse;
                    record.Klassenleitung = klassenleitung;

                    student.GetLetztesZeugnisdatumInDerKlasse(schuelerLernabschnittsdaten);

                    // Aktive SuS oder Schüler mit Abschluss/Abgang, deren letztes Zeugnis noch keine 42 Tage zurückliegt.
                    if (new List<string>() { "2", "6" }.Contains(student.Status) || (new List<string>(){"8", "9"}.Contains(student.Status) && student.ZeugnisdatumLetztesZeugnisInDieserKlasse.AddDays(42) >= DateTime.Now))
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
                    record.Anmeldedatum = student.BeginnDerBildungsganges;
                    record.Telefon = sz?["Telefon-Nr."].ToString();
                    record.Mobiltelefon = sz?["Fax/Mobilnr"].ToString();
                    record.email = student.MailSchulisch;
                    record.ZusatzInfo = "";
                    record.Bemerkung = "";
                    record.Geschlecht = student.Geschlecht.ToString().ToUpper();
                                        
                    if (student.Status is "2" or "6")
                    {
                        zieldatei.Add(record);
                    }
                }
            }

            if (susMitÄnderung.Count() > 2)
                AnsiConsole.Write(table);
            //    Global.DisplayCenteredBox(susMitÄnderung, 97);

            zieldatei.ZippeBilder("PfadZuAtlantisFotos");

            return zieldatei;
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
            Console.ReadKey();
            return null;
        }
    }

    public Datei KlassenErstellen(IConfiguration configuration, string absoluterPfad)
    {
        var zieldatei = new Datei(absoluterPfad);

        var schildKlassen = Quelldateien.GetMatchingList(configuration, "klassen", Students, Klassen);
        if (schildKlassen.Count == 0) return [];
        var untisKlassen = Quelldateien.GetMatchingList(configuration, "GPU003", Students, Klassen);
        if (untisKlassen.Count == 0) return [];

        var records = new List<dynamic>();

        foreach (var untisKlasse in untisKlassen)
        {
            var dictUntis = (IDictionary<string, object>)untisKlasse;

            if (dictUntis["Field1"].ToString() == "AGG25C")
            {
                string aa = "";
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
                // Wenn es diese Klasse in SchILD nicht gibt, wird sie angelegt
                dynamic record = new ExpandoObject();
                record.InternBez = dictUntis["Field1"].ToString();
                record.StatistikBez = dictUntis["Field1"].ToString();
                record.SonstigeBez = "";
                record.Jahrgang = s["Jahrgang"].ToString();
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

        return zieldatei;
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


    public void Sprechtag(Lehrers lehrers, Raums raums, IConfiguration configuration, string hinweis)
    {
        var dokuwikiZugriff = new DokuwikiZugriff(configuration);

        Global.Konfig("WikiSprechtagKleineAenderung", Global.Modus.Update, configuration,
            "Handelt es sich um eine kleine Änderung? Kleine Änderungen erzeugen keine neue Version (j/n)", "",
            Global.Datentyp.JaNein);

        dokuwikiZugriff.Options = new XmlRpcStruct
        {
            { "sum", "Automatische Aktualisierung" },
            { "minor", Global.WikiSprechtagKleineAenderung } // Kein Minor-Edit
        };

        var content = new List<string>();

        var exportLessons = Quelldateien.GetMatchingList(configuration, "exportlessons", IStudents, Klassen);
        if (exportLessons == null || !exportLessons.Any()) return;

        Global.Konfig("Sprechtagsdatum", Global.Modus.Update, configuration, "Datum des Sprechtags angeben (tt.mm.jjjj)", "", Global.Datentyp.DateTime);
        //Global.Konfig("wikiSprechtagSeite", true, "Seite eingeben, die manipuliert werden soll.");

        hinweis = hinweis.Replace(" nach der allgemeinen Zeugnisausgabe", ", " + Global.Sprechtagsdatum + ",");

        var alleLehrerImUnterrichtKürzel = exportLessons.Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["teacher"].ToString();
        }).Distinct().ToList();

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

        content.Add("====== Sprechtag ======");
        content.Add("");

        content.Add(hinweis);

        var i = 1;
        content.Add("");
        content.Add("<WRAP column 15em>");
        content.Add("");
        content.Add("^Name^Raum^");

        var lehrerProSpalteAufSeite2 = ((alleLehrerImUnterricht.Count - 60) / 3) + 1;

        foreach (var l in alleLehrerImUnterricht.OrderBy(x => x.Nachname))
        {
            var raum = string.IsNullOrEmpty(l.Raum) ? "" : l.Raum;

            // Wenn ein KuK außer Haus ist, wird der Raum aus Untis unterdrückt, bleibt aber in Untis für das nächste Jahr stehen. 
            if (!string.IsNullOrEmpty(l.Text2) && l.Text2.ToLower().Contains("außer haus"))
            {
                raum = "";
            }

            content.Add(
                "|" + (l.Geschlecht == "m" ? "Herr " : "Frau ") + (l.Titel == "" ? "" : l.Titel + " ") +
                l.Nachname + (l.Text2 == "" ? "" : " ((" + l.Text2 + "))") + "|" + raum + "|");

            if (i == 20 || i == 40 || i == 60 || i == 60 + lehrerProSpalteAufSeite2 ||
                i == 60 + lehrerProSpalteAufSeite2 * 2)
            {
                content.Add("</WRAP>");
                content.Add("");

                if (i == 60)
                {
                    content.Add("<WRAP pagebreak>");
                }

                content.Add("<WRAP column 15em>");
                content.Add("");
                content.Add("^Name^Raum^");
            }

            i++;
        }

        content.Add("</WRAP>");

        content.Add(
            "Klassenleitungen finden die Einladung als Kopiervorlage im [[sharepoint>:f:/s/Kollegium2/EjakJvXmitdCkm_iQcqOTLwB-9EWV5uqXE8j3BrRzKQQAw?e=OwxG0N|Sharepoint]].\r\n" +
            Environment.NewLine);

        content.Add("");

        var freieR = raums.OrderBy(x => x.Raumnummer)
            .Where(raum => !(from v in vergebeneRäume where v.Raumnummer == raum.Raumnummer select v).Any()).Aggregate(
                @"Sprechtag: Gewünschte Räume müssen in Untis in den Lehrer-Stammdaten eingetragen werden: ",
                (current, raum) => current + (raum.Raumnummer + " "));

        var panel = new Panel(freieR)
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Red);

                AnsiConsole.Write(panel);



        dokuwikiZugriff.PutPage("oeffentlich:sprechtag", string.Join("\n", content));

        Global.OpenWebseite("https://bkb.wiki/oeffentlich:sprechtag");
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

    public Datei GetFaecher(IConfiguration configuration, string zieldateiname)
    {

        var faecher = Quelldateien.GetMatchingList(configuration, "faecher", IStudents, Klassen);
        var zieldatei = new Datei(zieldateiname);

        var verschiedeneFaecher = faecher.Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Bezeichnung"];
        }).ToList().Distinct();

        foreach (var langname in verschiedeneFaecher)
        {
            var fach = faecher.FirstOrDefault(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Bezeichnung"].ToString() == langname.ToString();
            });

            if (langname == "") continue;
            dynamic record = new ExpandoObject();
            record.name = langname;
            record.Kuerzel = fach["InternKrz"].ToString();
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public Datei GetLehrer(IConfiguration configuration, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);
        var lehrkraefte = Quelldateien.GetMatchingList(configuration, "lehrkraefte", IStudents, Klassen);
        if (lehrkraefte.Count == 0)
        {
            return [];
        }

        foreach (var lehrer in lehrkraefte)
        {
            var dict = (IDictionary<string, object>)lehrer;

            dynamic record = new ExpandoObject();
            record.Kürzel = dict["InternKrz"].ToString();
            record.Vorname = dict["Vorname"].ToString();
            record.Nachname = dict["Nachname"].ToString();
            record.Name = (dict["Titel"].ToString() == "" ? "" : dict["Titel"] + " ") + dict["Vorname"] + " " +
                          dict["Nachname"];
            record.Mail = dict["E-Mail"].ToString();
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public Datei? Praktikanten(List<string> interessierendeKlassenUndJg, string zieldateiname)
    {

        var records = new List<dynamic>();
        var zieldatei = new Datei(zieldateiname);

        var praktikanten = new List<Student>();

        foreach (var item in interessierendeKlassenUndJg)
        {
            praktikanten.AddRange((from s in Students
                                   where s.Klasse.StartsWith(item.Split(',')[0])
                                   select s).ToList());
        }

        foreach (var praktikant in praktikanten)
        {
            if (praktikant == null) continue;
            dynamic record = new ExpandoObject();
            record.Name = praktikant.Nachname + ", " + praktikant.Vorname;
            record.Klasse = praktikant.Klasse;
            record.Jahrgang = praktikant.Jahrgang;
            record.Betrieb = "";
            record.Betreuung = "";
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public Datei? KlassenAnlegen(IConfiguration configuration, string zieldateiname)
    {
        var std = Students;
        var kl = Klassen;

        var klassen = Quelldateien.GetMatchingList(configuration, "klassen", Students, new Klassen());

        if (klassen.Count == 0)
        {
            return [];
        }

        var records = new List<dynamic>();

        var zieldatei = new Datei(zieldateiname);

        foreach (var klasse in klassen.OrderBy(x =>
                 {
                     var dictKlasse = (Dictionary<string, dynamic>)x;
                     return x["InternBez"].ToString();
                 }))
        {
            var dictKlasse = (Dictionary<string, dynamic>)klasse;

            dynamic record = new ExpandoObject();
            record.Name = dictKlasse["InternBez"].ToString();
            record.Klassenleitung = dictKlasse["Klassenlehrer"].ToString();
            record.Klassensprecher = "";
            record.Klassensprecher2 = "";
            zieldatei.Add(record);
        }

        return zieldatei;
    }

    public void Schulpflichtüberwachung(IConfiguration configuration)
    {
        var schuelerMitAbwesenheitenUndMaßnahmen = GetMaßnahmenUndFehlzeiten(configuration,
            [
                "Ordnungsmaßnahme",
                "Anhörung",
                "Bußgeldverfahren",
                "Attestpflicht",
                "Mahnung",
                "Familienkasse",
                "Versäumnisanzeige",
                "Suspendierung",
                "ausschluss"
            ]
        );

        schuelerMitAbwesenheitenUndMaßnahmen.SchulpflichtüberwachungTxt(
            configuration,
            "ImportNachWiki/schulpflichtueberwachung.txt",
            10, // Schonfrist: So viele Tage hat die Klassenleitung Zeit offene Stunden
                // zu bearbeiten, bevor eine Warnung ausgelöst wird.
            20, // Nach so vielen unent. Stunden ohne Maßnahme wird eine Warnung ausgelöst.
            30, // Nach so vielen Tagen verjähren unentschuldigte Fehlstunden für Unbescholtene.
            90, // Nach so vielen Tagen verjähren unentschuldigte Fehlstunden für SuS mit Maßnahme
            Klassen,
            Quelldateien
        );
    }

    public Students GetMaßnahmenUndFehlzeiten(IConfiguration configuration, List<string> maßnahmenString)
    {
        var sMitAbwesenheiten = new Students();

        var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", IStudents, Klassen);
        //.Where(datei => datei != null && !string.IsNullOrEmpty(datei!.UnterordnerUndDateiname))
        //.FirstOrDefault(datei => datei!.UnterordnerUndDateiname.Contains("SchuelerZusatzdaten"));

        foreach (Student student in Students)
        {
            var id = student.GetId(schuelerZusatzdaten);
            student.GetMaßnahmen(Quelldateien, maßnahmenString);
            student.GetAbwesenheiten(Quelldateien, id);

            if (student.Abwesenheiten.Count != 0)
                sMitAbwesenheiten.Add(student);
        }

        return sMitAbwesenheiten;
    }

    public Datei GetGruppen(IConfiguration configuration, string zieldateiname, Anrechnungen anrechnungen, Lehrers lehrers)
    {
        var zieldatei = new Datei(zieldateiname);
        var exportlessons = Quelldateien.GetMatchingList(configuration, "exportlesson", IStudents, Klassen);
        if (exportlessons == null || exportlessons.Count == 0) return [];

        Gruppen = new Gruppen();
        Gruppen.AddRange(new Gruppen().GetBildungsgaenge(exportlessons, Klassen, anrechnungen, lehrers));
        Gruppen.AddRange(new Gruppen().GetSchulformen(exportlessons, Klassen, anrechnungen, lehrers));
        Gruppen.Add(new Gruppe().Get(exportlessons, Klassen, anrechnungen, lehrers,
            "versetzung:blaue_briefe",
            new List<string>() { "BS", "HBG", "HBT", "HBW", "FS" },
            new List<int>() { 1 }));
        Gruppen.Add(new Gruppe().Get(exportlessons, Klassen, anrechnungen, lehrers,
            "termine:fhr:start",
            new List<string>() { "BS", "HBG", "HBT", "HBW", "FS", "FM" },
            new List<int>() { 2 }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:deutsch_kommunikation",
            new List<string>() { "D", "D FU", "D1", "D2", "D G1", "D G2", "D L1", "D L2", "D L", "DL", "DL1", "DL2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:englisch",
            new List<string>() { "E", "E FU", "E1", "E2", "E G1", "E G2", "E L1", "E L2", "E L", "EL", "EL1", "EL2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:religionslehre",
            new List<string>() { "KR", "KR FU", "KR1", "KR2", "KR G1", "KR G2", "ER", "ER G1" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:mathematik_physik",
            new List<string>() { "M", "M FU", "M1", "M2", "M G1", "M G2", "M L1", "M L2", "M L", "ML", "ML1", "ML2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:politik_gesellschaftslehre",
            new List<string>() { "PK", "PK FU", "PK1", "PK2", "GG G1", "GG G2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:wirtschaftslehre_in_nicht_kaufm_klassen",
            new List<string>() { "WL", "WBL" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:sport",
            new List<string>() { "SP", "SP G1", "SP G2" }));
        Gruppen.Add(new Gruppe().GetFachschaft(exportlessons, Klassen, anrechnungen, lehrers,
            ":fachschaften:biologie",
            new List<string>() { "BI", "Bi", "Bi FU", "Bi1", "Bi G1", "Bi G2", "BI G1", "BI L1" }));
        Gruppen.Add(new Gruppe().GetKollegium(exportlessons, Klassen, anrechnungen, lehrers,
            ":kollegium:start"));
        Gruppen.Add(new Gruppe().GetLehrerinnen(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:lehrerinnen"));
        Gruppen.Add(new Gruppe().GetRefs(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:referendar_innen"));
        Gruppen.Add(new Gruppe().GetKlassenleitungen(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:klassenleitungen"));

        Gruppen.Add(new Gruppe().GetBildungsgangleitungen(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:bildungsgangleitungen"));
        Gruppen.Add(new Gruppe().GetByWikilink(exportlessons, Klassen, anrechnungen, lehrers,
            "kollegium:schulleitung:erweiterte:start"));

        foreach (var gruppe in Gruppen)
        {
            zieldatei.Add(gruppe.Record);
        }
        return zieldatei;
    }


    public Datei MailadressenErgänzen(IConfiguration configuration, string zieldateiname)
    {
        var pfadSchilddatenaustausch = configuration["pfadSchilddatenaustausch"] ?? throw new ArgumentNullException(nameof(configuration), "'pfadSchilddatenaustausch' cannot be null.");
        var absoluterPfad = Path.Combine(pfadSchilddatenaustausch, zieldateiname);

        //Konfig("MailDomain", true, "Mail-Domain der Schüler*innen angeben. Bsp: '@students.berufskolleg-borken.de'");
        configuration = Global.Konfig("MailDomain", Global.Modus.Update, configuration);

        var zieldatei = new Datei(absoluterPfad);
        var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", Students, Klassen);
        if (schuelerZusatzdaten.Count == 0) return [];

        // Die Schild-ID muss aus dem SchildSchuelerExport ausgelesen werden
        var students = new Students(configuration, @"SchildSchuelerExport", "*.txt");
        if (Students.Count == 0)
        {
            return [];
        }

        foreach (var recZusatz in schuelerZusatzdaten)
        {
            var dictB = (IDictionary<string, object>)recZusatz;
            var vorname = dictB["Vorname"].ToString();
            var nachname = dictB["Nachname"].ToString();
            var geburtsdatum = dictB["Geburtsdatum"].ToString();

            var verschiedeneSuSmitDiesenDaten = schuelerZusatzdaten.Where(rec =>
            {
                var x = (IDictionary<string, object>)rec;
                return vorname == x["Vorname"].ToString() && nachname == x["Nachname"].ToString() &&
                       geburtsdatum == x["Geburtsdatum"].ToString();
            }).ToList();

            var verschiedenestudnts = students.Where(rec =>
                rec.Geburtsdatum == geburtsdatum && rec.Nachname == nachname && rec.Vorname == vorname).ToList();


            for (int i = 0; i < verschiedeneSuSmitDiesenDaten.Count; i++)
            {
                var x = verschiedeneSuSmitDiesenDaten[i];
            }


            var Zusatz = (IDictionary<string, object>)recZusatz;


            var student = students.Where(x =>
                x.Nachname == Zusatz["Nachname"].ToString() && x.Vorname == Zusatz["Vorname"].ToString() &&
                x.Geburtsdatum == Zusatz["Geburtsdatum"].ToString());

            dynamic record = new ExpandoObject();

            record.Nachname = Zusatz["Nachname"].ToString();
            record.Vorname = Zusatz["Vorname"].ToString();
            record.Geburtsdatum = Zusatz["Geburtsdatum"].ToString();
            record.Namenszusatz = Zusatz["Namenszusatz"].ToString();
            record.Geburtsname = Zusatz["Geburtsname"].ToString();
            record.Geburtsort = Zusatz["Geburtsort"].ToString();
            record.Ortsteil = Zusatz["Ortsteil"].ToString();
            record.TelefonMINUSNrPUNKT = Zusatz["Telefon-Nr."].ToString();
            record.EMINUSMail = Zusatz["E-Mail"].ToString();
            record.ZWEIPUNKTLEERZEICHENStaatsangPUNKT = Zusatz["2. Staatsang."].ToString();
            record.ExterneLEERZEICHENIDMINUSNr = Zusatz["Externe ID-Nr"].ToString();
            record.Sportbefreiung = Zusatz["Sportbefreiung"].ToString();
            record.Fahrschülerart = Zusatz["Fahrschülerart"].ToString();
            record.Haltestelle = Zusatz["Haltestelle"].ToString();
            record.Einschulungsart = Zusatz["Einschulungsart"].ToString();
            record.Entlassdatum = Zusatz["Entlassdatum"].ToString();
            record.Entlassjahrgang = Zusatz["Entlassjahrgang"].ToString();
            record.DatumLEERZEICHENSchulwechsel = Zusatz["Datum Schulwechsel"].ToString();
            record.Bemerkungen = Zusatz["Bemerkungen"].ToString();
            record.BKAZVO = Zusatz["BKAZVO"].ToString();
            record.BeginnBildungsgang = Zusatz["BeginnBildungsgang"].ToString();
            record.Anmeldedatum = Zusatz["Anmeldedatum"].ToString();
            record.Bafög = Zusatz["Bafög"].ToString();
            record.EPMINUSJahre = Zusatz["EP-Jahre"].ToString();
            record.FaxSCHRÄGSTRICHMobilnr = Zusatz["FaxSCHRÄGSTRICHMobilnr"].ToString();
            record.Ausweisnummer = Zusatz["Ausweisnummer"].ToString();
            record.schulischeLEERZEICHENEMINUSMail = "";
            throw new Exception();
            //record.MasernMINUSImpfnachweis = Zusatz["Masern-Impfnachweis"].ToString();
            //record.bisherigeLEERZEICHENID = "";
            //zieldatei.Add(record);
        }

        return zieldatei;
    }

    public enum Datentyp
    {
        String,
        Int,
        DateTime,
        Pfad,
        Url,
        JaNein,
        Klassen
    }

    public Datei Kalender2Wiki(IConfiguration configuration, string kalender, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname + ".csv");
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

                    dynamic record = new ExpandoObject();
                    record.Betreff = dict["Betreff"].ToString()!.Trim();
                    record.Seite = dict["Kategorien"].ToString().Split(';')[0];
                    record.Hinweise = "";
                    record.Datum = dat + zeit;
                    record.Kategorien = GetKategorien(dict["Kategorien"].ToString());
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
            return zieldatei;
        }

        return [];
    }

    private string GetKategorien(string? toString)
    {
        if (string.IsNullOrEmpty(toString))
        {
            return string.Empty;
        }

        var kategorien = toString.Split(';').Aggregate("", (current, str) => current + (str.Trim() + ","));

        return kategorien.TrimEnd(',');
    }

    public Datei Teilleistungen(IConfiguration configuration, string zieldateiname)
    {
        configuration = Global.Konfig("Teilleistungsarten", Global.Modus.Update, configuration, "Welche Teilleistungsarten (kommagetrennt) sollen gezogen werden?", "Wenn die von Ihnen eingegeben Teilleistungsart(en) ... \nin Webuntis anders heißen, werden keine Teilleistungen exportiert.\nin SchILD anders heißen, werden keine Teilleistungen nach SchILD importiert. Wenn die  Schauen Sie am besten zuerst in SchILD und Webuntis, welche Teilleistungsarten es in beiden Programmen gibt. Die können Sie dann hier angeben.", Global.Datentyp.String, "Vornote,Abschluss-Schriftl.,Abschluss-Mündl.");
        configuration = Global.Konfig("Abschnitt", Global.Modus.Update, configuration, "Abschnitt angeben", "Für welchen Abschnitt sollen die Teilleistungen ausgegeben werden?\nGeben die Ziffer [tan]1[/] (1. Halbjahr) oder [tan]2[/] (2. Halbjahr) an. ", Global.Datentyp.String, "");

        var zieldatei = new Datei(zieldateiname);
        var records = new List<dynamic>();

        var marksPerLs = Quelldateien.GetMatchingList(configuration, "marksperlesson", IStudents, Klassen);
        if (marksPerLs == null || marksPerLs.Count == 0) return [];

        try
        {
            foreach (var student in IStudents)
            {
                foreach (var recMar in marksPerLs)
                {
                    var dictMar = (IDictionary<string, object>)recMar;
                    if (!(dictMar["Name"].ToString() == student.Nachname + " " + student.Vorname && dictMar["Klasse"].ToString() == student.Klasse)) continue;
                    if (!(string.IsNullOrEmpty(configuration["Teilleistungsarten"].ToString()) || configuration["Teilleistungsarten"].ToString().Trim().Split(',').Contains(dictMar["Prüfungsart"].ToString()))) continue;
                    dynamic record = new ExpandoObject();
                    record.Nachname = student.Nachname;
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
                var panel = new Panel($"Keine Teilleistungen gefunden. Haben Sie die Teilleistungsart(en) exakt so eingegeben, wie sie in Webuntis mit Langname heißt?")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Red);

                AnsiConsole.Write(panel);
            }

            return zieldatei;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Console.ReadKey();
        }

        return null;
    }

    internal Datei GetLehrerDerKlassen(IConfiguration configuration, Lehrers lehrers)
    {
        var datei = new Datei();
        var mitgliederMail = "";

        var recExp = Quelldateien.GetMatchingList(configuration, "ExportLesson", IStudents, Klassen);
        if (recExp.Count == 0) return datei;

        var verschiedeneLulKuerzel = recExp
            .Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                var klassenString = dict["klassen"].ToString();
                var klassenListe = klassenString.Split('~'); // Zerlegt den String in eine Liste
                return IKlassen.Any(klasse => klassenListe.Contains(klasse)) &&
                       !string.IsNullOrEmpty(dict["teacher"].ToString());
            }).Select(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["teacher"].ToString();
            }).Distinct();

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
    internal Datei? LehrkraefteSonderzeiten(IConfiguration configuration, string zieldateiname, string nurDieseGründe = "")
    {
        var akt = int.Parse(Global.AktSj[0]);

        var zieldatei = new Datei(zieldateiname);

        var lehrkraefte = Quelldateien.GetMatchingList(configuration, "lehrkraefte", IStudents, Klassen);
        if (lehrkraefte == null || lehrkraefte.Count == 0) return [];

        var lehrkraefteSonderzeiten = Quelldateien.GetMatchingList(configuration, "lehrkraefteSonderzeiten", IStudents, Klassen);
        if (lehrkraefteSonderzeiten == null) return [];

        var gpu020 = Quelldateien.GetMatchingList(configuration, "gpu020", IStudents, Klassen); // Anrechnungen
        if (gpu020 == null) return [];

        var gpu004 = Quelldateien.GetMatchingList(configuration, "gpu004", IStudents, Klassen); // Lehrkraefte
        if (gpu004 == null) return [];

        if (nurDieseGründe == "")
            configuration = Global.Konfig("LehrkraefteSonderzeiten", Global.Modus.Update, configuration, "Welche Anrechnungsgründe sollen ignoriert bzw. auf 0 gesetzt werden?", "", Global.Datentyp.ListInt, "098,099,007,360,160");

        var lehrers = lehrkraefte
            .Where(rec => ((IDictionary<string, object>)rec)["statistik-relevant"].ToString() == "J")
            .ToList();

        var verschiedeneGründe = gpu020
            .Where(rec => ((IDictionary<string, object>)rec)["Field13"].ToString() != "0")
            .Select(rec => ((IDictionary<string, object>)rec)["Field5"].ToString())
            .Distinct()
            .OrderBy(x => x)
            .ToList();

        if (nurDieseGründe == "")
            verschiedeneGründe = gpu020
                .Where(rec => ((IDictionary<string, object>)rec)["Field13"].ToString() != "0" &&
                !configuration["LehrkraefteSonderzeiten"].ToString().Trim().Split(',').ToList().Contains(((IDictionary<string, object>)rec)["Field5"].ToString()) && ((IDictionary<string, object>)rec)["Field5"].ToString() != "")
                .Select(rec => ((IDictionary<string, object>)rec)["Field5"].ToString())
                .Distinct()
                .OrderBy(x => x)
                .ToList();

        if (nurDieseGründe != "")
            verschiedeneGründe = nurDieseGründe.Split(',').ToList();

        if (nurDieseGründe != "200")
            configuration["LehrkraefteSonderzeiten"] = "200";

            var panel = new Panel($"Die Anrechnungen aus der Untis-Datei [aqua]GPU020.txt[/] werden mit der SchILD-Datei [aqua]LehrkraefteSonderzeiten.dat[/] abgeglichen." +
                            $"\nDie Datei [aqua]{zieldateiname}[/] wird neu erstellt und kann nach ScHILD importiert werden." +
                            $"\nAnrechnungsgründe [aqua]...........................[/] werden ignoriert und auf 0 gesetzt.")
                            .Header($" [bold springGreen2] Anrechnungen [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Spectre.Console.Color.SpringGreen2);

        if (nurDieseGründe == "200")
            panel = new Panel($"Im Folgenden werden alle Lehrkräfte mit Altersermäßigung und alle mit fehlendem Deputat bzw. fehlendem Geburtsdatum angezeigt. " +
                                "Ohne Angabe des Deputats und Geburtsdatums in SchILD, findet keine Berechnung statt. " +
                                "Sch/Unt zeigt (möglicherweise) abweichende Werte in SchILD und in Untis. " +                                
                                $"Die Datei [aqua]{zieldateiname}[/] wird neu erstellt und kann nach ScHILD importiert werden. " +
                                "Auch wenn der Wert eines Anrechnungsgrunds bei einer Lehrkraft 0 ist, wird er in die Datei übernommen, um möglicherweise veraltete Werte zu überschreiben. " +
                                $"Da die Datei nur zusammen mit einer (leeren) lehrkraefte.dat importiert werden kann, wird eine leere lehrkraefte.dat ebenfalls erzeugt. " +
                                $"Wenn in der Spalte [bold springGreen2]Alter am 31.7. {akt + 1}[/] die Zahl 55 oder 60 steht, dann ändern sich die Werte im kommenden Jahr. ")
                    .HeaderAlignment(Justify.Left)
                    .SquareBorder()
                    .Expand()
                    .BorderColor(Color.Grey);

        AnsiConsole.Write(panel);
        
        var table = new Table();

        if (nurDieseGründe == "")
        {
            //table.AddColumn("Nr.");
            table.AddColumn("Lehrkraft");
            table.AddColumn("Grund");
            table.AddColumn("bisher");
            table.AddColumn("neu");
        }

        if (nurDieseGründe == "200")
        {
            //table.AddColumn(new TableColumn("Nr."));
            table.AddColumn(new TableColumn("LuL"));
            table.AddColumn(new TableColumn("Geb.dat."));
            table.AddColumn(new TableColumn("Deput. lt. Sch/Unt"));            
            table.AddColumn(new TableColumn("Stelle in %"));
            table.AddColumn(new TableColumn($"Alter am 31.7.{akt}"));
            table.AddColumn(new TableColumn($"Alter am 31.7.{akt + 1}"));            
            table.AddColumn(new TableColumn($"IST laut Sch/Unt"));
            table.AddColumn(new TableColumn($"SOLL {akt}/ {akt + 1} lt. SchILD"));
            table.AddColumn(new TableColumn($"SOLL {akt + 1}/ {akt + 2} lt.SchILD"));
        }

        if (nurDieseGründe == "200")
            configuration = Global.Konfig("VolleStelle", Global.Modus.Update, configuration, "Wie viele Stunden entsprechen einer vollen Stelle?", "", Global.Datentyp.Float, "25,5");

        foreach (var lehrerDyn in lehrers)
        {
            var lehrer = new Lehrer();
            // Spalte 2
            lehrer.Kürzel = ((IDictionary<string, object>)lehrerDyn)["InternKrz"].ToString();
            // Spalte 3
            lehrer.Geburtsdatum = lehrer.GetGeburtsdatum(((IDictionary<string, object>)lehrerDyn)["Geburtsdatum"].ToString());
            // Spalte 4
            lehrer.DeputatLautSchild = lehrer.GetDeputatSchild(((IDictionary<string, object>)lehrerDyn)["Pflichtstunden-Soll"].ToString());

            foreach (var grund in verschiedeneGründe)
            {
                var istWertSonderzeitLautSchild = lehrer.GetWertSonderzeiten(grund, lehrkraefteSonderzeiten);
                var istWertGpu020LautUntis = lehrer.GetAnrechnungswertGPU020Soll(gpu020, grund);
                lehrer.DeputatLautUntis = lehrer.GetDeputatLautUntis(gpu004);

                if (nurDieseGründe == "200")
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

                        table.AddRow(
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
                if (nurDieseGründe == "")
                {
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

                    if (wertIst != wertDerAnrechnung.ToString().Replace(',', '.'))
                    {
                        table.AddRow(
                            lehrer.Kürzel,
                            grund,
                            wertIst.ToString().ToString().Replace(",00", ""),
                            wertDerAnrechnung.ToString("0.##", CultureInfo.InvariantCulture));
                    }
                }

                dynamic record = new ExpandoObject();
                record.Lehrkraft = lehrer.Kürzel;
                record.Zeitart = grund.ToString().StartsWith("1") ? "MEHRLEISTUNG" : grund.ToString().StartsWith("2") ? "MINDERLEISTUNG" : "ANRECHNUNG";
                record.Grund = grund;
                record.AnzahlLEERZEICHENStunden = nurDieseGründe == "" ? wertDerAnrechnung.ToString().Replace('.', ',') : lehrer.AltersermäßigungSoll.ToString("F2", CultureInfo.InvariantCulture).Replace('.', ',');
                zieldatei.Add(record);
            }
        }
        AnsiConsole.Write(table);

        return zieldatei;
    }

    internal Datei? Lehrkraefte(IConfiguration configuration, string zieldateiname)
    {
        var zieldatei = new Datei(zieldateiname);

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
        return zieldatei;
    }

    public void RenderAuswahlÜberschrift(IConfiguration configuration)
    {
        var panel3 = new Panel($"[grey85]{string.Join("\n", Beschreibung)}[/]")
            .Header($" [bold yellow3_1]  {Titel}  [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Color.Yellow3_1);
        
        AnsiConsole.Write(panel3);
    }
}