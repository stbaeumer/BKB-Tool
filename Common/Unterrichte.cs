using Common;
using Microsoft.Extensions.Configuration;
using System.Text.RegularExpressions;
using Spectre.Console;

public class Unterrichte : List<Unterricht>
{
    public Unterrichte(IConfiguration configuration, Menüeintrag m, Global.Zweck zweck, Global.Art art)
    {
        // Aus der Kurse.dat wird die Kursart ermittelt. Wenn die Kursart in SchILD einmal gesetzt ist, wird sie nicht mehr geändert.
        List<dynamic> kurseDat = m.Quelldateien.GetMatchingList(configuration, "kurse.", m.IStudents, m.Klassen);
        if (kurseDat == null)        
            throw new FileNotFoundException($"[grey]Keine Kurse.dat gefunden. Bitte exportieren Sie die Datei erneut.[/]");

        List<dynamic> gpu002 = m.Quelldateien.GetMatchingList(configuration, "gpu002", m.IStudents, m.Klassen);
        if (gpu002 == null)
            throw new FileNotFoundException($"[grey]Keine GPU002-Daten gefunden. Bitte exportieren Sie die Datei erneut.[/]");
                if (gpu002.Count == 0)
                    throw new FileNotFoundException($"[grey]Die Datei GPU002 ist leer. UTF8, Textbegrenzer \" und Delimiter | müssen gesetzt sein.[/]");

        List<dynamic> studentgroupStudents = m.Quelldateien.GetMatchingList(configuration, "studentgroupstudents", m.IStudents, m.Klassen);
        if (studentgroupStudents == null)
            throw new FileNotFoundException($"[grey]Keine Kurse instudentgroupStudents gefunden. Bitte exportieren Sie die Datei erneut.[/]");
        if (studentgroupStudents.Count == 0)
            AnsiConsole.MarkupLine($"[grey] Keine Zeilen in studentgroupStudents gefunden. Ist das korrekt?.[/]");
        
        var zulässigeAuswahlOptionen = GetZulässigeUnterrichtsgruppen(configuration, gpu002);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"{art} aus GPU002.TXT einlesen ...", ctx =>
        {
            // Ordne die GPU002 aufsteigend nach Field6. Dadurch wird erreicht, dass die erste Lehrkraft im Alphabet Kursleiter wird.
            gpu002 = gpu002.OrderBy(record => { var dict = (IDictionary<string, object>)record; return dict.ContainsKey("Field6") ? dict["Field6"]?.ToString() : string.Empty; }).ToList();

#region Hilfsfunktionen (Data-Parsing)

/// <summary>
/// Prüft, ob ein Untis/GPU-Datensatz eine Wochenstundenzahl größer als 0 aufweist (Field11).
/// </summary>
bool HasHours(IDictionary<string, object> r)
{
    if (r.ContainsKey("Field11") && r["Field11"] != null)
    {
        if (double.TryParse(r["Field11"]?.ToString()?.Trim(), 
            System.Globalization.NumberStyles.Any, 
            System.Globalization.CultureInfo.InvariantCulture, out var parsed))
        {
            // Untis skaliert manche Stundenwerte intern mit Faktor 100.000
            var hours = parsed >= 1000 ? parsed / 100000.0 : parsed;
            return hours > 0;
        }
    }
    return false;
}

/// <summary>
/// Liest das Lehrerkürzel aus den bevorzugten Feldern (Field6, Fallback: Field5) aus.
/// </summary>
string GetTeacherFromRecord(IDictionary<string, object> r)
{
    string[] candidateFields = { "Field6", "Field5" };

    foreach (var field in candidateFields)
    {
        if (r.ContainsKey(field) && r[field] != null)
        {
            var val = r[field]?.ToString()?.Trim()?.Trim('"');
            if (!string.IsNullOrEmpty(val) && val != "0")
            {
                return val;
            }
        }
    }
    return string.Empty;
}

/// <summary>
/// Liest die Schülergruppen-Bezeichnung aus Field42 (bzw. Field43) aus und ignoriert Ungültige/Zahlenwerte.
/// </summary>
string GetStudentGroupFromRecord(IDictionary<string, object> r)
{
    string[] candidateFields = { "Field42", "Field43" };

    foreach (var field in candidateFields)
    {
        if (r.ContainsKey(field) && r[field] != null)
        {
            var val = r[field]?.ToString()?.Trim()?.Trim('"');

            if (string.IsNullOrEmpty(val) || val == "0" || val == "n") 
                continue;

            // Reine Zahlen ignorieren (keine gültigen Schülergruppen-Namen)
            if (double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
                continue;

            return val;
        }
    }

    return string.Empty;
}

/// <summary>
/// Liest die Unterrichtsgruppe (U-Gruppe) aus Field12 (bzw. Field13) aus.
/// </summary>
string GetUGroupFromRecord(IDictionary<string, object> r)
{
    string[] candidateFields = { "Field12", "Field13" };

    foreach (var field in candidateFields)
    {
        if (r.ContainsKey(field) && r[field] != null)
        {
            var val = r[field]?.ToString()?.Trim()?.Trim('"');
            if (!string.IsNullOrEmpty(val) && val != "0")
            {
                return val;
            }
        }
    }

    return string.Empty;
}

#endregion

#region 1. SCHRITT: Bereinigung der Rohdaten

// Exakte Duplikate aus gpu002 entfernen und Datensätze ohne Stunden (0 h) herausfiltern
var records = gpu002
    .Cast<IDictionary<string, object>>()
    .GroupBy(r => string.Join("|", r.Select(kv => $"{kv.Key}={kv.Value}")))
    .Select(g => g.First())
    .Where(HasHours)
    .ToList();

#endregion

#region 2. SCHRITT: Ermittlung von Klassenunterricht vs. Kurse (Voranalyse)

// Gruppierung nach Stammfach + Klassenkombination zur Identifikation redundanter Kurse
var rawGroups = records
    .GroupBy(r => 
    {
        var rawFach = r.ContainsKey("Field7") ? r["Field7"]?.ToString() ?? "" : "";
        var baseFach = GetBaseSubjectWithCourseType(rawFach);
        var klassen = r.ContainsKey("Field5") ? r["Field5"]?.ToString() ?? "" : "";
        
        return $"{baseFach}_{klassen}";
    })
    .ToList();

// Metadaten pro Rohgruppe analysieren
var groupInfos = rawGroups.Select(g =>
{
    var recordsList = g.ToList();
    var sample = recordsList.First();
    var rawSampleFach = sample.ContainsKey("Field7") ? sample["Field7"]?.ToString() ?? "" : "";
    var baseFach = GetBaseSubjectWithCourseType(rawSampleFach);

    var klassenList = recordsList
        .Select(r => r.ContainsKey("Field5") ? r["Field5"]?.ToString() ?? "" : "")
        .Where(k => !string.IsNullOrEmpty(k))
        .Distinct()
        .ToList();

    var studentGroupsList = recordsList
        .Select(r => GetStudentGroupFromRecord(r))
        .Where(sg => !string.IsNullOrEmpty(sg))
        .Distinct()
        .ToList();

    var teachersList = recordsList
        .Select(r => GetTeacherFromRecord(r))
        .Where(t => !string.IsNullOrEmpty(t))
        .Distinct()
        .ToList();

    // Ein Unterricht gilt als Kurs, wenn Schülergruppen vorhanden sind oder mehrere Klassen/Lehrer beteiligt sind
    bool isKurs = studentGroupsList.Any() || klassenList.Count > 1 || teachersList.Count > 1;

    return new
    {
        Group = g,
        BaseFach = baseFach,
        Klassen = klassenList,
        IsKurs = isKurs
    };
}).ToList();

// Set aller Kombinationen aus (Klasse + BaseFach) bilden, die echten KLASSENUNTERRICHT darstellen
var klassenUnterrichtSet = groupInfos
    .Where(gi => !gi.IsKurs)
    .SelectMany(gi => gi.Klassen.Select(k => $"{k}_{gi.BaseFach}"))
    .ToHashSet();

// Strikte Gruppierung nach Stammfach für die finale Aggregation
var finalGroups = records
    .GroupBy(r => 
    {
        var rawFach = r.ContainsKey("Field7") ? r["Field7"]?.ToString() ?? "" : "";
        return GetBaseSubjectWithCourseType(rawFach);
    })
    .ToList();

#endregion

#region 3. SCHRITT: UI-Vorbereitung (Spectre.Console)

var targetFormList = new List<object>();

var table = new Table();
table.Border(TableBorder.Rounded);
table.AddColumn(new TableColumn("[bold yellow]Fach[/]").Centered());
table.AddColumn(new TableColumn("[bold cyan]Lehrer (Stunden)[/]"));
table.AddColumn(new TableColumn("[bold orange1]Schülergruppen[/]"));
table.AddColumn(new TableColumn("[bold teal]U-Gruppe[/]").Centered());
table.AddColumn(new TableColumn("[bold white]IDs[/]").Centered());
table.AddColumn(new TableColumn("[bold green]Klassen[/]"));
table.AddColumn(new TableColumn("[bold dim]Roh-Fächer[/]"));
table.AddColumn(new TableColumn("[bold magenta]Unterrichtstyp[/]").Centered());
table.AddColumn(new TableColumn("[bold blue]Kursbezeichnung[/]").Centered());
table.AddColumn(new TableColumn("[bold grey]Gültig von-bis[/]").Centered());

#endregion

#region 4. SCHRITT: Aggregation & Erzeugung der Unterricht-Objekte

DateTime statistikDatum = tt_mm_jjjjNachDateTime(configuration["StatistikDatum"]);

foreach (var group in finalGroups)
{
    var baseFach = group.Key;
    var rawRecordsInGroup = group.ToList();

    // -------------------------------------------------------------------------------------
    // AUFGABE 2: Datensätze anhand von Stunden, Lehrer, Konfiguration & Statistikdatum filtern
    // -------------------------------------------------------------------------------------
    var allRecordsInGroup = rawRecordsInGroup.Where(dict =>
    {
        // 1. Stunden aus Field2 lesen & auf 0 prüfen
        double wochentundenLehrkraft = 0;
        if (dict.ContainsKey("Field2") && dict["Field2"] != null)
        {
            double.TryParse(dict["Field2"]?.ToString()?.Trim(), 
                System.Globalization.NumberStyles.Any, 
                System.Globalization.CultureInfo.InvariantCulture, out wochentundenLehrkraft);
        }
        if (wochentundenLehrkraft == 0) return false;

        // 2. Ohne gewiesenen Lehrer überspringen
        if (!dict.ContainsKey("Field6") || string.IsNullOrEmpty(dict["Field6"]?.ToString()))
            return false;

        // 3. Konfigurationsbasierte Filter anwenden
        if (nichtInteressierdendeUnterrichtsgruppe(configuration, dict))
            return false;

        if (nichtInteressierdendeFächer(configuration, dict))
            return false;

        // 4. Statistikdatum-Prüfung gegen Zeitraum (Field15 = Von, Field16 = Bis)
        if (zweck == Global.Zweck.Statistik)
        {
            var von = dict.ContainsKey("Field15") ? jjjjmmddNachDateTime(dict["Field15"]) : DateTime.MinValue;
            var bis = dict.ContainsKey("Field16") ? jjjjmmddNachDateTime(dict["Field16"]) : DateTime.MaxValue;

            if (von > statistikDatum || statistikDatum > bis)
                return false;
        }

        return true;
    }).ToList();

    // Falls nach der Filterung keine Datensätze übrig sind, abbrechen
    if (!allRecordsInGroup.Any()) continue;

    // -------------------------------------------------------------------------------------
    // AUFGABE 1: Gültigkeitszeitraum (Von / Bis) bestimmen
    // -------------------------------------------------------------------------------------
    var vonDates = allRecordsInGroup
        .Where(r => r.ContainsKey("Field15") && !string.IsNullOrEmpty(r["Field15"]?.ToString()))
        .Select(r => jjjjmmddNachDateTime(r["Field15"]))
        .Where(d => d != DateTime.MinValue)
        .ToList();

    var bisDates = allRecordsInGroup
        .Where(r => r.ContainsKey("Field16") && !string.IsNullOrEmpty(r["Field16"]?.ToString()))
        .Select(r => jjjjmmddNachDateTime(r["Field16"]))
        .Where(d => d != DateTime.MinValue)
        .ToList();

    DateTime? minVon = vonDates.Any() ? vonDates.Min() : null;
    DateTime? maxBis = bisDates.Any() ? bisDates.Max() : null;

    string mehrausgabeDatum = (minVon.HasValue && maxBis.HasValue)
        ? $"{minVon.Value:dd.MM.yyyy} - {maxBis.Value:dd.MM.yyyy}"
        : "-";

    // -------------------------------------------------------------------------------------
    // Aggregierte Metadaten der Gruppe sammeln
    // -------------------------------------------------------------------------------------
    var ids = allRecordsInGroup
        .Select(r => r.ContainsKey("Field1") ? r["Field1"]?.ToString() ?? "" : "")
        .Where(id => !string.IsNullOrEmpty(id))
        .Distinct()
        .OrderBy(id => id)
        .ToList();

    var klassen = allRecordsInGroup
        .Select(r => r.ContainsKey("Field5") ? r["Field5"]?.ToString() ?? "" : "")
        .Where(k => !string.IsNullOrEmpty(k))
        .Distinct()
        .OrderBy(k => k)
        .ToList();

    var rawFaecher = allRecordsInGroup
        .Select(r => r.ContainsKey("Field7") ? r["Field7"]?.ToString() ?? "" : "")
        .Where(f => !string.IsNullOrEmpty(f))
        .Distinct()
        .OrderBy(f => f)
        .ToList();

    var studentGroupsList = allRecordsInGroup
        .Select(r => GetStudentGroupFromRecord(r))
        .Where(sg => !string.IsNullOrEmpty(sg))
        .Distinct()
        .OrderBy(sg => sg)
        .ToList();

    var uGroupsList = allRecordsInGroup
        .Select(r => GetUGroupFromRecord(r))
        .Where(ug => !string.IsNullOrEmpty(ug))
        .Distinct()
        .OrderBy(ug => ug)
        .ToList();

    // -------------------------------------------------------------------------------------
    // Lehrkräfte & Stunden aggregieren (Spalte 2 / Field2)
    // -------------------------------------------------------------------------------------
    var teacherHoursMap = new Dictionary<string, double>();

    var uniqueLessonRecords = allRecordsInGroup
        .GroupBy(r => new 
        { 
            Id = r.ContainsKey("Field1") ? r["Field1"]?.ToString() ?? "" : "",
            Lehrer = GetTeacherFromRecord(r)
        })
        .Select(g => g.First())
        .ToList();

    foreach (var r in uniqueLessonRecords)
    {
        var lehrer = GetTeacherFromRecord(r);
        if (string.IsNullOrEmpty(lehrer)) continue;

        double hours = 0;
        if (r.ContainsKey("Field2") && r["Field2"] != null)
        {
            var rawVal = r["Field2"]?.ToString()?.Trim();
            if (double.TryParse(rawVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var parsed))
            {
                if (parsed >= 100000)
                    hours = parsed / 100000.0;
                else if (parsed >= 100)
                    hours = parsed / 100.0;
                else
                    hours = parsed;
            }
        }

        if (hours > 0)
        {
            if (!teacherHoursMap.ContainsKey(lehrer))
                teacherHoursMap[lehrer] = 0;

            teacherHoursMap[lehrer] += hours;
        }
    }

    if (!teacherHoursMap.Any()) continue;

    // -------------------------------------------------------------------------------------
    // Logik zur Bestimmung des Unterrichtstyps (Kurs vs. Klassenunterricht)
    // -------------------------------------------------------------------------------------
    bool mehereLehrer = teacherHoursMap.Count > 1;
    bool hatKlassenunterrichtDatensatz = allRecordsInGroup.Any(r => string.IsNullOrEmpty(GetStudentGroupFromRecord(r)));

    string unterrichtsTyp;
    List<string> effektiveSchuelerGruppen;

    if (mehereLehrer)
    {
        unterrichtsTyp = "Kurs";
        effektiveSchuelerGruppen = new List<string>();
    }
    else if (hatKlassenunterrichtDatensatz)
    {
        unterrichtsTyp = "Klassenunterricht";
        effektiveSchuelerGruppen = new List<string>();
    }
    else
    {
        bool hasStudentGroupInGroup = studentGroupsList.Any();
        unterrichtsTyp = hasStudentGroupInGroup || klassen.Count > 1
            ? "Kurs" 
            : "Klassenunterricht";
        effektiveSchuelerGruppen = studentGroupsList;
    }

    // -------------------------------------------------------------------------------------
    // AUFGABE 3: Schüleranzahl pro Schülergruppe ermitteln & formatieren
    // -------------------------------------------------------------------------------------
    var effektiveSchuelerGruppenMitAnzahl = effektiveSchuelerGruppen.Select(schuelergruppe =>
    {
        int schuelerInKurs = studentgroupStudents
            .Select(record => (IDictionary<string, object>)record)
            .Where(dict => dict.ContainsKey("studentgroup.name") && 
                           !string.IsNullOrEmpty(dict["studentgroup.name"]?.ToString()))
            .Where(dict => dict["studentgroup.name"].ToString() == schuelergruppe)  
            .Count();

        return $"{schuelergruppe} ({schuelerInKurs})";
    }).ToList();

    // -------------------------------------------------------------------------------------
    // AUFGABE 4: Kursbezeichnung für SchILD generieren (max. 20 Zeichen)
    // -------------------------------------------------------------------------------------
    string kursBez = "-";
    if (unterrichtsTyp == "Kurs")
    {
        var ersterLehrer = teacherHoursMap.Keys.FirstOrDefault() ?? "";
        kursBez = $"{ersterLehrer}-{string.Join('-', ids)}";

        if (kursBez.Length > 20)
        {
            kursBez = kursBez.Substring(0, 20);
        }
    }

    // -------------------------------------------------------------------------------------
    // Konsolenausgabe & UI-Tabelle befüllen
    // -------------------------------------------------------------------------------------
    var lehrerMitStunden = string.Join(", ", teacherHoursMap.Select(kv => $"{kv.Key} ({kv.Value:0.##})"));
    var schuelerGruppenStr = string.Join(", ", effektiveSchuelerGruppenMitAnzahl);
    var uGruppenStr = string.Join(", ", uGroupsList);
    var idListStr = string.Join(", ", ids);
    var klassenStr = string.Join(", ", klassen);
    var faecherStr = string.Join(", ", rawFaecher);

    var typMarkup = unterrichtsTyp == "Kurs" 
        ? $"[magenta]{unterrichtsTyp}[/]" 
        : $"[blue]{unterrichtsTyp}[/]";

    table.AddRow(
        Markup.Escape(baseFach),
        Markup.Escape(lehrerMitStunden),
        Markup.Escape(string.IsNullOrEmpty(schuelerGruppenStr) ? "-" : schuelerGruppenStr),
        Markup.Escape(string.IsNullOrEmpty(uGruppenStr) ? "-" : uGruppenStr),
        Markup.Escape(idListStr),
        Markup.Escape(klassenStr),
        Markup.Escape(faecherStr),
        typMarkup,
        Markup.Escape(kursBez),
        Markup.Escape(mehrausgabeDatum)
    );

    targetFormList.Add(new
    {
        BaseFach = baseFach,
        Lehrer = lehrerMitStunden,
        Schuelergruppe = schuelerGruppenStr,
        UGruppe = uGruppenStr,
        IDs = idListStr,
        Klassen = klassenStr,
        OriginalFaecher = faecherStr,
        Typ = unterrichtsTyp,
        Kursbezeichnung = kursBez,
        VonBis = mehrausgabeDatum,
        RawRecords = allRecordsInGroup
    });

    // -------------------------------------------------------------------------------------
    // Domain-Objekt 'Unterricht' instanziieren & zur Sammlung hinzufügen
    // -------------------------------------------------------------------------------------
    var primaryTeacherPair = teacherHoursMap.FirstOrDefault();
    string primaryTeacher = primaryTeacherPair.Key ?? "";
    double primaryTeacherHours = primaryTeacherPair.Value;

    var coTeachers = teacherHoursMap.Skip(1).ToList();

    // -------------------------------------------------------------------------------------
// JAHRGÄNGE BERECHNEN (aus Klassennamen wie HBW25A, FS26B etc.)
// -------------------------------------------------------------------------------------
int aktuellesSchuljahrStart = DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1; // Sept. 2026 -> 2026

var jahrgaengeList = klassen
    .Select(k =>
    {
        // Sucht nach den ersten zwei aufeinanderfolgenden Ziffern im Klassennamen
        var match = System.Text.RegularExpressions.Regex.Match(k, @"\d{2}");
        if (match.Success && int.TryParse(match.Value, out int einschulungsJahrShort))
        {
            int einschulungsJahr = 2000 + einschulungsJahrShort;
            int jahrgangNummer = aktuellesSchuljahrStart - einschulungsJahr + 1;

            if (jahrgangNummer > 0)
            {
                return jahrgangNummer.ToString("D2"); // Formatiert z. B. 2 zu "02"
            }
        }
        return null;
    })
    .Where(j => j != null)
    .Distinct()
    .OrderBy(j => j)
    .Cast<string>()
    .ToList();


    var unterricht = new Unterricht
    {        
        Fach = Bereinigen(baseFach),
        KursBez = kursBez == "-" ? "" : kursBez,
        Kursart = unterrichtsTyp == "Klassenunterricht" ? "PUK" : "",
        
        // Erstlehrer / Kursleiter
        Kursleiter = primaryTeacher,
        KursleiterWochenstunden = Convert.ToInt32(primaryTeacherHours),
        Wochenstunden = Convert.ToInt32(primaryTeacherHours),

        // Weitere Lehrkräfte (Teamteaching)
        Lehrkraefte = coTeachers.Select(ct => ct.Key).ToList(),
        LehrkraefteWochenstunden = coTeachers.Select(ct => Convert.ToInt32(ct.Value)).ToList(),

        // IDs & Klassenzuordnung
        UnterrichtsIds = ids.Select(idStr => int.TryParse(idStr, out var parsedId) ? parsedId : 0)
                            .Where(id => id > 0)
                            .ToList(),
        Klassen = klassen,
        Jahrgaenge = jahrgaengeList,

        // Schülergruppen & gefilterte Schülerobjekte
        Schülergruppe = schuelerGruppenStr,
        Students = m.IStudents.Filter(configuration, zweck, klassen.FirstOrDefault() ?? "", schuelerGruppenStr, studentgroupStudents)
    };
    if(!string.IsNullOrEmpty(unterricht.KursBez))
        unterricht.Kursart = unterricht.GetKursart(configuration, kurseDat, unterricht.Fach, primaryTeacher, ids[0]);

    // Unterricht zum aktuellen Container/Liste hinzufügen
    this.Add(unterricht);
    
}

// Rendern der Spectre.Console Tabelle auf der Konsole
AnsiConsole.Write(table);

#endregion

            
});}

 string GetBaseSubjectWithCourseType(string subject)
{
    if (string.IsNullOrWhiteSpace(subject)) return string.Empty;

    var trimmed = subject.Trim();

    // Schneidet bei z. B. "EW L1" die hintere Ziffer ab -> "EW L"
    if (trimmed.Length > 4)
    {
        trimmed = trimmed.Substring(0, 4);
    }
    else if (trimmed.Length == 4 && char.IsDigit(trimmed[3]))
    {
        trimmed = trimmed.Substring(0, 3);
    }

    return trimmed.TrimEnd();
}












 /*




             foreach (var record in gpu002)
             {
                 var Klassen = new Klassen();
                 var dict = (IDictionary<string, object>)record;

                 // ohne eingetragenen Lehrer wird die Zeile übersprungen
                 if (!dict.ContainsKey("Field6") || string.IsNullOrEmpty(dict["Field6"]?.ToString()))
                     continue; // Ohne Lehrer wird die Zeile übersprungen.

                 if (nichtInteressierdendeUnterrichtsgruppe(configuration, dict))
                     continue; // Diese Gruppe ist nicht interessant, also überspringen.

                 if (nichtInteressierdendeFächer(configuration, dict))
                     continue; // Diese Fächer sind nicht interessant, also überspringen.

                 if (zweck == Global.Zweck.Statistik && !dict.ContainsKey("Field12") && außerhalbDesStatistikdatums(configuration, dict))
                     continue; // Statistikdatum liegt außerhalb des Zeitraums.

                 DateTime von = jjjjmmddNachDateTime(dict["Field15"]);
                 DateTime bis = jjjjmmddNachDateTime(dict["Field16"]);
                 DateTime statistikDatum = tt_mm_jjjjNachDateTime(configuration["StatistikDatum"]);

                 // Wenn das Statistikdatum außerhalb des Zeitraums von Field15 und Field16 liegt, wird die Zeile übersprungen.
                 if (zweck == Global.Zweck.Statistik && (von > statistikDatum || statistikDatum > bis))
                     continue; // Statistikdatum liegt außerhalb des Zeitraums.

                 var unterrichtsId = dict["Field1"]?.ToString();
                 var fach = Bereinigen(dict["Field7"]?.ToString());
                 var schuelergruppe = dict["Field42"]?.ToString();
                 var klasse = dict["Field5"]?.ToString();
                 var lehrer = dict["Field6"]?.ToString();
                 int wochentundenLehrkraft = int.TryParse(dict["Field2"]?.ToString(), out int ws) ? ws : 0;
                 int wochenstundenKurs = 0;
                 int schuelerInKurs = studentgroupStudents
                     .Select(record => (IDictionary<string, object>)record)
                     .Where(dict => dict.ContainsKey("studentgroup.name") && 
                     !string.IsNullOrEmpty(dict["studentgroup.name"]?.ToString()))
                     .Where(dict => dict["studentgroup.name"].ToString() == schuelergruppe)  
                     .Count();

                 if(unterrichtsId == "1545" || unterrichtsId == "2152")
                 {
                     string a = "";
                 }

                 if (wochentundenLehrkraft == 0) continue; // Unterrichte mit 0 Wochenstunden werden nicht berücksichtigt.

                 // Wenn es um Kurse geht und die Zeile ein Kurs ist oder zu einem Kurs gehört
                 if (art == Global.Art.KursUnterrichte && zeileIstKursOderGehörtZuKurs(gpu002, dict, m.IStudents, schuelerInKurs))
                 {
                     // Fall 1
                     // Zwei oder mehr Lehrer unterrichten dasselbe Fach in einer Kopplung ohne Schülergruppen
                     // Hinzu kommt, dass einer der Lehrer allein weitere Stunden im Fach unterrichtet








                     // Suche nach einem bestehenden Kurs mit identischer UntisID & Lehrkraft
                     var kurs = this.FirstOrDefault(k => k.UnterrichtsIds.Contains(Convert.ToInt32(unterrichtsId)) && k.Lehrkraefte.Contains(lehrer));

                     // Wenn es zu dieser UntisID bereits einen Kurs gibt, ...
                     if (kurs != null)
                     {
                         // dann wird der Kurs bei identischer Schülergruppe (oder wenn beide leer sind) upgedatet (Fall 1)
                         if(kurs.Schülergruppe == schuelergruppe || (string.IsNullOrEmpty(kurs.Schülergruppe) && string.IsNullOrEmpty(schuelergruppe)))                        
                         {                         
                             // Wochenstunden werden bei Bedarf erhöht usw.
                             var kursBez = kurs.Updaten(zweck, m, configuration, fach, lehrer, unterrichtsId, schuelergruppe, klasse, wochentundenLehrkraft, studentgroupStudents);                        
                         }
                         else // oder bei unterschiedlicher Schülergruppe (Fall2:) ein neuer Kurs erstellt.
                         {
                             var kursBez = $"{lehrer}-{unterrichtsId}";
                             Add(new Unterricht(kursBez, zweck, m, unterrichtsId, fach, schuelergruppe, klasse, lehrer, wochentundenLehrkraft, kurseDat, configuration, studentgroupStudents));
                         }
                     }

                     // Wenn kein passender Kurs gefunden wurde:
                     if (kurs == null)
                     {
                         kurs = this.FirstOrDefault(k =>
                         Bereinigen(k.Fach) == Bereinigen(fach) &&
                         k.Schülergruppe == schuelergruppe &&
                         (k.Klassen.Contains(klasse) ||
                             (k.KursBez.StartsWith(lehrer) && k.KursBez.Contains(unterrichtsId) // Kopplung von zwei Klassen in Untis
                         )));

                         // Wochenstunden werden bei Bedarf erhöht usw.
                         if(kurs != null)
                         {
                             var kursBez = kurs.Updaten(zweck, m, configuration, fach, lehrer, unterrichtsId, schuelergruppe, klasse, wochentundenLehrkraft, studentgroupStudents);                            
                         }
                     }

                     if (kurs == null)
                     {
                         var kursBez = $"{lehrer}-{unterrichtsId}";
                         Add(new Unterricht(kursBez, zweck, m, unterrichtsId, fach, schuelergruppe, klasse, lehrer, wochentundenLehrkraft, kurseDat, configuration, studentgroupStudents));
                     }                        
                 }
                 else if (art == Global.Art.NichtKursUnterrichte && !zeileIstKursOderGehörtZuKurs(gpu002, dict, m.IStudents, schuelerInKurs))
                 {
                     var nichtKursUnterricht = this.FirstOrDefault(u =>
                         Bereinigen(u.Fach) == Bereinigen(fach) &&
                         u.Kursleiter == lehrer &&
                         u.Schülergruppe == schuelergruppe &&
                         u.Klassen.Contains(klasse));

                     if (nichtKursUnterricht == null)
                     {
                         Add(new Unterricht(zweck, m, configuration, unterrichtsId, fach, schuelergruppe, klasse, lehrer, wochentundenLehrkraft, studentgroupStudents));
                     }
                     else
                     {
                         // Wochenstunden werden bei Bedarf erhöht usw.                        
                         nichtKursUnterricht.Updaten(zweck, m, configuration, fach, lehrer, unterrichtsId, schuelergruppe, klasse, wochentundenLehrkraft, studentgroupStudents);
                     }
                 }
             }
         });

         // Pro Klasse wird eine Tabelle ausgegeben
         // distincte Liste aller Klassen aus this:
         var alleKlassen = this.SelectMany(u => u.Klassen).Distinct().ToList();

         foreach (var klasse in alleKlassen)
         {
             var table = new Table();
             table.Border = TableBorder.Rounded;
             table.Expand();
             table.Title = new TableTitle($"{art} der Klasse {klasse} aus GPU002", new Style(foreground: Color.Grey, decoration: Decoration.Bold));
             if (art == Global.Art.KursUnterrichte)
                 table.AddColumn("Kursbezeichnung");
             if (art == Global.Art.NichtKursUnterrichte)
                 table.AddColumn("Unterrichtsnummern");            
             table.AddColumn("Fach");
             table.AddColumn("WStd");
             table.AddColumn("Kursleiter(Wstd)");
             if (art == Global.Art.KursUnterrichte)
                 table.AddColumn($"LuL(WStd)");
             table.Columns[2].Alignment = Justify.Right;

             foreach (var unterricht in this.Where(u => u.Klassen.Contains(klasse)).OrderBy(u => u.Fach).ThenBy(u => u.Kursleiter))
             {
                 var kursleiter = $"{unterricht.Kursleiter}({unterricht.KursleiterWochenstunden})";                

                 if (art == Global.Art.KursUnterrichte)
                     table.AddRow(
                         new Markup($"{unterricht.Kursleiter}-{string.Join('-', unterricht.UnterrichtsIds)}"),                        
                         new Markup(unterricht.Fach),
                         new Markup(unterricht.Wochenstunden.ToString()),
                         new Markup(kursleiter),
                         new Markup($"{string.Join(" ", unterricht.Lehrkraefte.Select((lk, index) => $"{lk}({unterricht.LehrkraefteWochenstunden[index]})"))}")
                     );

                 if(art == Global.Art.NichtKursUnterrichte)
                     table.AddRow(
                         new Markup($"{string.Join(' ', unterricht.UnterrichtsIds)}"),
                         new Markup(unterricht.Fach),
                         new Markup(unterricht.KursleiterWochenstunden.ToString()),
                         new Markup(kursleiter)
                     );
             }
             int wochenstunden = this.Sum(u => u.Wochenstunden);

             table.AddRow("", "", $"[{Global.GetColor(Global.ColorZahlen)}]{wochenstunden}[/]", "");
             if(configuration["Klassen"]?.Split(',').Select(s => s.Trim()).Contains(klasse) == true)
                 AnsiConsole.Write(table);
         }   
     }*/

 private bool nichtInteressierdendeFächer(IConfiguration configuration, IDictionary<string, object> dict)
 {
        // Wenn Field7 leer ist oder configuration["InteressierendeFächer"].Split(',').Contains(Field7) dann wird die Zeile weiterverarbeitet.
        if (dict.ContainsKey("Field7") && !string.IsNullOrEmpty(dict["Field7"]?.ToString()))
        {
            var nichtInteressierendeFächer = configuration["NichtInteressierendeFächer"]?.Split(',').Select(s => s.Trim()).ToList();
            if (nichtInteressierendeFächer != null && nichtInteressierendeFächer.Contains(dict["Field7"].ToString()))
            {
                return true; // Dieses Fach ist nicht interessant, also überspringen.
            }
        }
        return false;
 }

 private DateTime tt_mm_jjjjNachDateTime(string? v)
    {
        if (string.IsNullOrEmpty(v) || v.Length != 10)
            return DateTime.MinValue; // Ungültiges Datum

        if (DateTime.TryParseExact(v, "dd.MM.yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime dateTime))
        {
            return dateTime;
        }
        return DateTime.MinValue; // Ungültiges Datum
    }

    public static DateTime? jjjjmmddNachDateTime(object input)
{
    if (input == null) return null;
    
    var str = input.ToString()?.Trim();
    if (string.IsNullOrEmpty(str) || str.Length < 8) return null;

    if (DateTime.TryParseExact(str, "yyyyMMdd", 
        System.Globalization.CultureInfo.InvariantCulture, 
        System.Globalization.DateTimeStyles.None, 
        out var dt))
    {
        return dt;
    }

    return null;
}

    private string GetZulässigeUnterrichtsgruppen(IConfiguration configuration, List<dynamic> gpu002)
    {
        var x = gpu002
            .Select(record => (IDictionary<string, object>)record)
            .Where(dict => dict.ContainsKey("Field12") && !string.IsNullOrEmpty(dict["Field12"]?.ToString()))
            .Select(dict => dict["Field12"].ToString())
            .Distinct()
            .OrderBy(s => s)
            .ToList();
        if (x.Count == 0)
        {
            AnsiConsole.MarkupLine($"[grey]      Keine Unterrichtsgruppen gefunden.[/]");
            return string.Empty;
        }
        return string.Join(",", x);
    }

    public Unterrichte()
    {
    }

    private string GetJahrgang(List<dynamic> klassen, IDictionary<string, object> dict)
    {
        var jahrgangRec = klassen
                    .FirstOrDefault(record =>
                    {
                        var dictKlasse = (IDictionary<string, object>)record;
                        return dictKlasse["InternBez"]?.ToString() == dict["Field5"]?.ToString();
                    }) as IDictionary<string, object>;

                    return jahrgangRec != null ? jahrgangRec["Jahrgang"]?.ToString() : null;
    }

    private bool nichtInteressierdendeUnterrichtsgruppe(IConfiguration configuration, IDictionary<string, object> dict)
    {
        // Wenn Field12 leer ist oder configuration["InteressierendeUnterrichtsgruppen"].Split(',').Contains(Field12) dann wird die Zeile weiterverarbeitet.
        if (dict.ContainsKey("Field12") && !string.IsNullOrEmpty(dict["Field12"]?.ToString()))
        {
            var nichtInteressierendeGruppen = configuration["NichtInteressierendeUnterrichtsgruppen"]?.Split(',').Select(s => s.Trim()).ToList();
            if (nichtInteressierendeGruppen != null && nichtInteressierendeGruppen.Contains(dict["Field12"].ToString()))
            {
                return true; // Diese Gruppe ist nicht interessant, also überspringen.
            }
        }
        return false;
    }

    private bool außerhalbDesStatistikdatums(IConfiguration configuration, IDictionary<string, object> dict)
    {
        // Wenn das Statistikdatum außerhalb der Zeitspanne von Field15 (yyyymmdd) und Field16 (yyyymmdd) liegt, wird die Zeile übersprungen.
        if (dict.ContainsKey("Field15") && dict.ContainsKey("Field16"))
        {
            if (DateTime.TryParseExact(dict["Field15"].ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime startDate) &&
                DateTime.TryParseExact(dict["Field16"].ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime endDate))
            {
                var statistikDatum = configuration["StatistikDatum"];
                if (!DateTime.TryParse(statistikDatum, out DateTime statistikDate) || statistikDate < startDate || statistikDate > endDate)
                {
                    return true; // Statistikdatum liegt außerhalb des Zeitraums.
                }
            }
        }
        return false; // Statistikdatum liegt innerhalb des Zeitraums.
    }

    private bool zeileIstKursOderGehörtZuKurs(List<dynamic> gpu002, IDictionary<string, object> dict, Students students = null, int schuelerInKurs = 0)
    {
        if (dict["Field1"]?.ToString() == "2676" || dict["Field1"]?.ToString() == "2263")
        {
            string aa = "Test"; // Debugging purpose
        }

        // Ein Kurs ist definiert, wenn Field42 (Schülergruppe) nicht leer ist.
        if (dict.ContainsKey("Field42") && !string.IsNullOrEmpty(dict["Field42"]?.ToString()))
        {   
            // ... es sei denn, dass alle SuS einer Klasse in dem Unterricht sind, dann ist es kein Kurs, auch wenn Field42 nicht leer ist. Das wird hier überprüft, indem geschaut wird, ob die Schülergruppe (Field42) alle Schüler einer Klasse enthält.
            if (students != null && dict.ContainsKey("Field5") && !string.IsNullOrEmpty(dict["Field5"]?.ToString()))
            {
                var klasse = dict["Field5"].ToString();
                var schuelerInKlasse = students.Where(s => s.Klasse == klasse).ToList();
                
                if (schuelerInKurs > 0 && schuelerInKurs == schuelerInKlasse.Count)
                {
                    return false; // Es ist kein Kurs, wenn alle Schüler einer Klasse in der Schülergruppe sind.
                }
            }
            return true; // Es ist ein Kurs, wenn Field42 nicht leer ist.
        }

        // Ein Kurs ist auch definiert, wenn Field1 mehrfach vorkommt und Field42 leer ist.
        if (dict.ContainsKey("Field1") && !string.IsNullOrEmpty(dict["Field1"]?.ToString()))
        {
            // Überprüfen, ob Field1 mehrfach vorkommt
            var field1 = dict["Field1"].ToString();
            var count = gpu002.Count(record =>
            {
                var recDict = (IDictionary<string, object>)record;
                return recDict.ContainsKey("Field1") && recDict["Field1"].ToString() == field1;
            });
            if (count > 1)
            {
                return true; // Es ist ein Kurs, wenn Field1 mehrfach vorkommt.    
            }
        }

        // Ein Kurs ist definiert, wenn dasselbe Fach (Field7) und dieselbe Klasse (Field5) mehrfach vorkommen, auch wenn Field42 leer ist.
        // Wenn allerdings der Lehrer (Field6) immer identisch ist, dann ist es kein Kurs.
        // Ein Zähler am Ende des Fachs wird nicht berücksichtigt. Bsp.: M1 wird zu M., M G1 bleibt M G1, weil M das Fach ist und G1 der Grundkurs.

        var count1 = gpu002.FirstOrDefault(record =>
        {
            var recDict = (IDictionary<string, object>)record;

            if(recDict["Field1"].ToString() == "1410")
            {
                string aa = "Test"; // Debugging purpose
            }

            return Bereinigen(recDict["Field7"].ToString()) == Bereinigen(dict["Field7"].ToString()) &&
            recDict["Field5"].ToString() == dict["Field5"].ToString() &&
            recDict["Field42"].ToString() == dict["Field42"].ToString() &&
            recDict["Field6"].ToString() != dict["Field6"].ToString();
        });

        // Es ist ein Kurs, wenn dasselbe Fach und dieselbe Klasse mit einem anderen Lehrer nochmal vorkommen.
        if (count1 != null)
        {
            return true;
        }

        return false;
    }

    private string Bereinigen(string? fach)
    {
        string fachBereinigt = fach;
        if (!string.IsNullOrEmpty(fach) && fach.Length > 1 && char.IsDigit(fach[^1]) && fach.Count(c => c == ' ') == 0)
        {
            // Wenn das Fach nur aus Buchstaben und einer Ziffer am Ende besteht (z.B. "M1"), entferne die Ziffer
            fachBereinigt = fach.Substring(0, fach.Length - 1);
        }
        return fachBereinigt;
    }
}