using Common;
using Microsoft.Extensions.Configuration;
using System.Text.RegularExpressions;
using Spectre.Console;
using System.Globalization;
using System.Data.SqlTypes;
using System.Xml.Serialization;

public class Unterrichte : List<Unterricht>
{
    public Unterrichte(IConfiguration configuration, Menüeintrag m, Global.Zweck zweck, Global.Art art)
    {
        // Aus der Kurse.dat wird die Kursart ermittelt. Wenn die Kursart in SchILD einmal gesetzt ist, wird sie nicht mehr geändert.
        List<dynamic> kurseDat = m.Quelldateien.GetMatchingList(configuration, "kurse.", m.IStudents, m.Klassen);
        if (kurseDat == null)        
            throw new FileNotFoundException($"[grey]Keine Kurse.dat gefunden. Bitte exportieren Sie die Datei erneut.[/]");

        List<dynamic> faecherDat = m.Quelldateien.GetMatchingList(configuration, "faecher", m.Students, m.Klassen);
        if (faecherDat == null)        
            throw new FileNotFoundException($"[grey]Keine Faecher.dat gefunden. Bitte exportieren Sie die Datei erneut.[/]");

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

            List<string> klassen = GetKlassen(gpu002);

            foreach(var klasse in m.IKlassen)
            {
                var table = new Table();
                table.Border(TableBorder.Rounded);
                table.AddColumn(new TableColumn("[bold yellow]Fach[/]"));
                table.AddColumn(new TableColumn("[bold cyan]Lehrer(Stunden)[/]"));
                table.AddColumn(new TableColumn("[bold orange1]Schülergruppen[/]"));
                table.AddColumn(new TableColumn("[bold white]IDs (U-Gruppe)[/]"));
                table.AddColumn(new TableColumn("[bold green]Klassen[/]"));
                table.AddColumn(new TableColumn("[bold dim]Roh-Fächer[/]"));
                table.AddColumn(new TableColumn("[bold magenta]Unterrichtstyp[/]").Centered());
                table.AddColumn(new TableColumn("[bold blue]Kursbezeichnung[/]"));
                table.AddColumn(new TableColumn("[bold grey]Gültig von-bis[/]").Centered());
                table.AddColumn(new TableColumn("[bold grey]SuS[/]").Centered());

                IEnumerable<object> klassenZeilen = GetZeilenZuKlasse(klasse, gpu002);

                List<string> faecherRoh = GetFaecherRoh(klassenZeilen);
                
                if(faecherRoh.Contains("EVAB"))
                {
                    string aa = "";
                }
                
                List<string> faecherBereinigt = GetFaecherBereinigt(faecherRoh, faecherDat);
           
                foreach (var fach in faecherBereinigt)
                {
                    var fachZeilen = GetZeilenZuKlassenFaechern(klasse, fach, klassenZeilen, faecherDat);
                    List<int> idsBeteiligteSortiert = GetBeteiligteIds(fachZeilen);
                    var lehrerDesFachs = GetLehrerDesFachs(fach, fachZeilen, faecherDat);
                    var lehrerFachs = GetLehrerFachs(fach, fachZeilen, faecherDat);
                    List<string> faecherRohFach = GetFaecherRoh(fachZeilen);
                    bool fachMitSchülergruppe = istFachMitSchülergruppe(fachZeilen);
                   
                    var susDerKlasse = m.IStudents.Where(x=>x.Klasse == klasse).ToList();

                    foreach(var s in susDerKlasse)
                    {                        
                        // Gib die Schülergruppen aus dieses S in diesem Fach
                        var schuelergruppenDiesesSchuelersinDiesemKlassenfach = GetSchuelergruppenDiesesSusInDiesemFach(s.Id, studentgroupStudents, fachZeilen);
                        
                        foreach(var sg in schuelergruppenDiesesSchuelersinDiesemKlassenfach)
                        {                            
                            // Erstelle den Kurs wie folgt als string: "Kurs|Fach|Leh1,Leh2|Id1,Id2|2,1"
                            // wobei Kurs als Wort stehenbleibt. Danach werden alle beteiligten Lehrer (also weitere Zeilen mit derselben ID und demselben fach) am Fachunterricht bei diesem Schüler aufgeführt.
                            //Console.WriteLine(GetUnterrichteString(fachZeilen, sg));
                            var unterricht = GetUnterricht(fachZeilen, faecherRohFach, faecherDat, sg);

                            if (!this.Any(u => u.KursBez == unterricht.KursBez))
                            {
                                unterricht.Students.Add(s);
                                this.Add(unterricht);
                            }
                            else
                            {
                                var u = this.FirstOrDefault(u => u.KursBez == unterricht.KursBez);                                
                                u.UpdateUnterricht(s, unterricht);
                            }                                
                        }

                        // Wenn es kein Schülergruppenfach ist, dann bekommt jeder Schüler diesen Unterricht zugewiesen:

                        if(!fachMitSchülergruppe)
                        {
                            if(fach == "SP G")
                            {
                                string a = "";
                            }
                            var unterricht = GetUnterricht(fachZeilen, faecherRohFach, faecherDat);
                            
                            if (!this.Any(u => u.KursBez == unterricht.KursBez))
                            {
                                unterricht.Students.Add(s);
                                this.Add(unterricht);                     
                            }
                            else
                            {
                                var u = this.FirstOrDefault(u => u.KursBez == unterricht.KursBez);
                                u.UpdateUnterricht(s, unterricht);
                            }
                        }
                    }
                }

                foreach(var u in this)
                {
                    if (u.Klassen.Contains(klasse))
                    {
                        Spectre.Console.TableExtensions.AddRow(
                            table,
                            Markup.Escape(u.Fach),
                            Markup.Escape($"{u.Kursleiter}({u.KursleiterWochenstunden}) " + string.Join(" ", u.Lehrkraefte.Zip(u.LehrkraefteWochenstunden, (lk, std) => $"{lk}({std})"))).Trim(),
                            Markup.Escape(string.Join(',',u.Schülergruppe)),
                            Markup.Escape(string.Join(',',u.UnterrichtsIds)),
                            Markup.Escape(string.Join(',',u.Klassen)),
                            Markup.Escape(string.Join(',',u.FaecherRoh)),
                            Markup.Escape(u.Kursart),
                            Markup.Escape(u.KursBez),
                            Markup.Escape($"{u.Von:dd.MM.yyyy} - {u.Bis:dd.MM.yyyy}"),
                            Markup.Escape(u.Students.Count()+ "")
                        );
                    }

                    // Bei PUK wird die Kursbezeichnung entfernt, weil eine existierende die Kursbezeichnung 
                    // zu einem Eintrag in Kurse.dat führen wird.
                    if(u.Kursart == "PUK")
                    {
                        u.KursBez = "";
                        u.KursBezUngekürzt = "";
                    }
                }

                var reversed = this.AsEnumerable().Reverse().ToList();
                this.Clear();
                this.AddRange(reversed);

                // Option 1b: Eine stylische Trennlinie (Rule) als Überschrift
                AnsiConsole.Write(new Rule("[bold blue]" + klasse + "[/]").LeftJustified());
                AnsiConsole.Write(table);
            }            
        }
    );
}


 private bool istFachMitSchülergruppe(IEnumerable<object> fachZeilen)
{
    if (fachZeilen == null)
        return false;

    return fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "")
        .Any(sg => !string.IsNullOrWhiteSpace(sg) && sg != "?");
}

 private Unterricht GetUnterricht(IEnumerable<object> fachZeilen, List<string> faecherRohFach, IEnumerable<dynamic> faecherDat, string schuelergruppe = null)
{
    if (fachZeilen == null)
        return null;

    // Optionalen Filterwert vorab trimmen
    var schuelergruppeFilter = schuelergruppe?.Trim().Trim('"');

    var relevanteZeilen = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => new
        {
            Lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "",
            RawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Klasse = d.ContainsKey("Field5") ? d["Field5"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            VonStr = d.ContainsKey("Field15") ? d["Field15"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            BisStr = d.ContainsKey("Field16") ? d["Field16"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            IdStr = ExtrahiereUnterrichtId(d)?.Trim().Trim('"') ?? "",
            Stunden = ExtrahiereStundenWert(d)
        })
        .Where(x => (string.IsNullOrWhiteSpace(schuelergruppeFilter) || 
                     x.Schuelergruppe.Equals(schuelergruppeFilter, StringComparison.OrdinalIgnoreCase)) &&
                    !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.RawFach) && x.RawFach != "?")
        .ToList();

    if (!relevanteZeilen.Any())
        return null;

        if (relevanteZeilen.First().RawFach.EndsWith("1"))
        {
            string aaa = "";
        }

    // 1. Fachbereinigung
    var erstesFach = relevanteZeilen.First().RawFach;
    var fachbereinigt = GetBaseSubjectWithCourseType(erstesFach, faecherDat);

    // 2. Lehrer verarbeiten & Stunden pro Lehrer aufsummieren (alphabetisch sortiert)
    var lehrerGruppen = relevanteZeilen
        .GroupBy(x => x.Lehrer)
        .OrderBy(g => g.Key)
        .Select(g => new
        {
            Lehrer = g.Key,
            Stunden = g.Sum(x => x.Stunden)
        })
        .ToList();

    var ersterLehrer = lehrerGruppen.First();
    var weitereLehrer = lehrerGruppen.Skip(1).ToList();

    // 3. Unterrichts-IDs verarbeiten
    var unterrichtsIds = relevanteZeilen
        .Select(x => int.TryParse(x.IdStr, out var id) ? (int?)id : null)
        .Where(id => id.HasValue)
        .Select(id => id.Value)
        .Distinct()
        .OrderBy(id => id)
        .ToList();

    // 4. KursBezeichnung (Ungekürzt & Auf 20 Zeichen gekürzt)
    var idsJoined = string.Join("-", unterrichtsIds);
    var kursBezUngekuerzt = string.IsNullOrEmpty(idsJoined) 
        ? ersterLehrer.Lehrer 
        : $"{ersterLehrer.Lehrer}-{idsJoined}";

    var kursBezGekuerzt = kursBezUngekuerzt.Length > 20 
        ? kursBezUngekuerzt.Substring(0, 20) 
        : kursBezUngekuerzt;

    // 5. Maximale Wochenstunden ermitteln
    var maxWochenstunden = lehrerGruppen.Any() 
        ? lehrerGruppen.Max(x => x.Stunden) 
        : 0;

    // 6. Klassen und Datums-Parsing
    var klassenList = relevanteZeilen
        .Select(x => x.Klasse)
        .Where(k => !string.IsNullOrWhiteSpace(k) && k != "?")
        .Distinct()
        .OrderBy(k => k)
        .ToList();

    DateTime.TryParseExact(relevanteZeilen.FirstOrDefault()?.VonStr, "yyyyMMdd", 
        System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime vonDatum);

    DateTime.TryParseExact(relevanteZeilen.FirstOrDefault()?.BisStr, "yyyyMMdd", 
        System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime bisDatum);

    // 7. Unterrichts-Objekt befüllen

    var kursart = "PUK";

    if(!string.IsNullOrEmpty(schuelergruppeFilter))
    {
        kursart = "Kurs";
    }
    if(weitereLehrer.Count() > 0)
    {
        kursart = "Kurs";
    }

    // 8. Unterrichts-Objekt befüllen
    return new Unterricht
    {
        Fach = fachbereinigt,
        KursBez = kursBezGekuerzt,
        KursBezUngekürzt = kursBezUngekuerzt,
        Kursart = kursart,
        Kursleiter = ersterLehrer.Lehrer,
        KursleiterWochenstunden = (int)ersterLehrer.Stunden,
        Lehrkraefte = weitereLehrer.Select(x => x.Lehrer).ToList(),
        LehrkraefteWochenstunden = weitereLehrer.Select(x => (int)x.Stunden).ToList(),
        Wochenstunden = (int)maxWochenstunden,
        UnterrichtsIds = unterrichtsIds,
        Klassen = klassenList,
        Jahrgaenge = new List<string>(),
        Schülergruppe = schuelergruppeFilter ?? "", // Falls kein Wert übergeben wurde, wird ein leerer String gesetzt
        Students = new Students(),
        Von = vonDatum,
        Bis = bisDatum,
        FaecherRoh = faecherRohFach
    };
}

 private string GetUnterrichteString(IEnumerable<object> fachZeilen, string schuelergruppe, IEnumerable<object> faecherDat)
{
    if (fachZeilen == null || string.IsNullOrWhiteSpace(schuelergruppe))
        return string.Empty;

    var relevanteZeilen = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => new
        {
            Lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "",
            RawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Unterrichtsgruppe = d.ContainsKey("Field12") ? d["Field12"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Von = d.ContainsKey("Field15") ? d["Field15"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Bis = d.ContainsKey("Field16") ? d["Field16"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Id = ExtrahiereUnterrichtId(d)?.Trim().Trim('"') ?? "",
            Stunden = ExtrahiereStundenWert(d)
        })
        .Where(x => x.Schuelergruppe.Equals(schuelergruppe.Trim().Trim('"'), StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.RawFach) && x.RawFach != "?")
        .ToList();

    if (!relevanteZeilen.Any())
        return string.Empty;

    // Bereinigtes Fach bestimmen (BaseFach)
    var erstesFach = relevanteZeilen.First().RawFach;
    var baseFach = GetBaseSubjectWithCourseType(erstesFach, faecherDat);

    // Lehrer alphabetisch sortieren
    var lehrerGruppen = relevanteZeilen
        .GroupBy(x => x.Lehrer)
        .OrderBy(g => g.Key)
        .ToList();

    var lehrerList = lehrerGruppen.Select(g => g.Key).ToList();
    var stundenList = lehrerGruppen.Select(g => g.Sum(x => x.Stunden).ToString()).ToList();

    // IDs pro Lehrer sammeln (intern mit Leerzeichen, Lehrer-Blöcke per Komma getrennt)
    var idBlöckeProLehrer = lehrerGruppen
        .Select(g => 
        {
            var ids = g.Select(x => x.Id)
                       .Where(id => !string.IsNullOrWhiteSpace(id) && id != "?")
                       .Distinct()
                       .OrderBy(id => id);

            return string.Join(" ", ids);
        });

    // Schülergruppen absolut DISTINCT über alle Zeilen sammeln
    var alleSchuelergruppenDistinct = relevanteZeilen
        .Select(x => x.Schuelergruppe)
        .Where(sg => !string.IsNullOrWhiteSpace(sg) && sg != "?")
        .Distinct()
        .OrderBy(sg => sg);

    // Unterrichtsgruppe aus Field12 pro Lehrer sammeln
    var uGruppenProLehrer = lehrerGruppen
        .Select(g => 
        {
            var uGruppen = g.Select(x => x.Unterrichtsgruppe)
                            .Where(ug => !string.IsNullOrWhiteSpace(ug) && ug != "?")
                            .Distinct()
                            .OrderBy(ug => ug);

            return string.Join(" ", uGruppen);
        });

    // Zeitraum (Von/Bis aus Field15/Field16) ermitteln
    var vonDatum = relevanteZeilen
        .Select(x => x.Von)
        .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v) && v != "?") ?? "";

    var bisDatum = relevanteZeilen
        .Select(x => x.Bis)
        .FirstOrDefault(b => !string.IsNullOrWhiteSpace(b) && b != "?") ?? "";

    var lehrerJoined = string.Join(",", lehrerList);
    var idsJoined = string.Join(",", idBlöckeProLehrer);
    var stundenJoined = string.Join(",", stundenList);
    var gruppenJoined = string.Join(", ", alleSchuelergruppenDistinct);
    var uGruppenJoined = string.Join(",", uGruppenProLehrer);
    var zeitraumPart = (!string.IsNullOrEmpty(vonDatum) || !string.IsNullOrEmpty(bisDatum)) 
        ? $"{vonDatum} - {bisDatum}" 
        : "";

    return $"Kurs|{baseFach}|{lehrerJoined}|{idsJoined}|{stundenJoined}| {gruppenJoined}|{uGruppenJoined}|{zeitraumPart}";
}

 private List<string> GetSchuelergruppenDiesesSusInDiesemFach(
    string studentId,
    IEnumerable<object> studentgroupStudents, 
    IEnumerable<object> fachZeilen)
{
    if (string.IsNullOrWhiteSpace(studentId) || studentgroupStudents == null || fachZeilen == null)
        return new List<string>();

    var studentIdClean = CleanString(studentId);

    // 1. Gültige Schülergruppen aus fachZeilen (z. B. "DVG1_GG26A")
    var gueltigeSchuelergruppen = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => CleanString(GetDictValue(d, "Field42")))
        .Where(sg => !string.IsNullOrWhiteSpace(sg) && sg != "?")
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    if (!gueltigeSchuelergruppen.Any())
        return new List<string>();

    // 2. studentgroupStudents filtern
    var erg = new List<string>();

    foreach (var obj in studentgroupStudents)
    {
        if (obj is not IDictionary<string, object> d)
            continue;

        // Versuche Schüler-ID über verschiedene Key-Möglichkeiten zu finden
        var currentStudentId = CleanString(
            GetDictValue(d, "student.name", "student.key", "student.id", "student", "schueler", "Field1", "Field2")
        );

        if (!currentStudentId.Equals(studentIdClean, StringComparison.OrdinalIgnoreCase))
            continue;

        // Versuche Schülergruppe über verschiedene Key-Möglichkeiten zu finden
        var sgName = CleanString(
            GetDictValue(d, "studentgroup.name", "studentgroup", "gruppe", "schuelergruppe", "Field4", "Field5")
        );

        if (!string.IsNullOrWhiteSpace(sgName) && gueltigeSchuelergruppen.Contains(sgName))
        {
            erg.Add(sgName);
        }
    }

    return erg.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x).ToList();
}

// Hilfsmethode: Durchsucht das Dictionary flexibel nach Keys (Case-Insensitive und Fallbacks)
private string GetDictValue(IDictionary<string, object> dict, params string[] possibleKeys)
{
    // 1. Exakter/Vergleichbarer Match aus den gewünschten Keys
    foreach (var key in possibleKeys)
    {
        var foundKey = dict.Keys.FirstOrDefault(k => k.Equals(key, StringComparison.OrdinalIgnoreCase));
        if (foundKey != null && dict[foundKey] != null)
            return dict[foundKey].ToString();
    }

    // 2. Fallback: Falls Keys z.B. "student_id" oder "studentName" heißen
    foreach (var key in possibleKeys)
    {
        var partialKey = dict.Keys.FirstOrDefault(k => k.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0);
        if (partialKey != null && dict[partialKey] != null)
            return dict[partialKey].ToString();
    }

    return "";
}

// Hilfsmethode: Bereinigt Strings gründlich von Anführungszeichen, Leerzeichen, Tabs und NBSP
private string CleanString(string input)
{
    if (string.IsNullOrWhiteSpace(input))
        return "";

    return input
        .Replace("\u00A0", " ") // Geschütztes Leerzeichen durch normales ersetzen
        .Replace("\t", " ")     // Tabulatoren ersetzen
        .Trim()
        .Trim('"')
        .Trim();
}

 private List<string> GetFachLehIdStundenGekoppelt(IEnumerable<object> fachZeilen, IEnumerable<object> faecherDat)
{
    if (fachZeilen == null)
        return new List<string>();

    var zeilenList = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => new
        {
            Lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "",
            RawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Id = ExtrahiereUnterrichtId(d)?.Trim().Trim('"') ?? "",
            Stunden = ExtrahiereStundenWert(d)
        })
        .Where(x => !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.RawFach) && x.RawFach != "?")
        .ToList();

    if (!zeilenList.Any())
        return new List<string>();

    var result = new List<string>();

    // Gruppieren nach bereinigtem Fach und Lehrer
    var lehrerGruppen = zeilenList
        .GroupBy(x => new 
        { 
            BaseFach = GetBaseSubjectWithCourseType(x.RawFach, faecherDat), 
            x.Lehrer 
        })
        .OrderBy(g => g.Key.BaseFach)
        .ThenBy(g => g.Key.Lehrer);

    foreach (var lg in lehrerGruppen)
    {
        var fach = lg.Key.BaseFach;
        var lehrer = lg.Key.Lehrer;

        // IDs sammeln und alphabetisch/numerisch kommagetrennt verknüpfen (z.B. "1675" oder "1678,1679")
        var ids = string.Join(",", lg
            .Select(x => x.Id)
            .Where(id => !string.IsNullOrWhiteSpace(id) && id != "?")
            .Distinct()
            .OrderBy(id => id));

        // Gesamtstunden des Lehrers in diesem gekoppelten Fach ermitteln
        var gesamtStunden = lg.Sum(x => x.Stunden);
        var stundenText = gesamtStunden == 1 ? "1 Stunde" : $"{gesamtStunden} Stunden";

        // Klammer-Inhalt aufbauen: ID erzwingen
        var idUndStundenPart = string.IsNullOrEmpty(ids)
            ? $"({stundenText})"
            : $"({ids}; {stundenText})";

        result.Add($"{fach}-{lehrer}{idUndStundenPart}");
    }

    return result;
}

 private bool gekoppeltMitAnderenFächernOhneSchülergruppe(IEnumerable<object> fachZeilen, IEnumerable<object> klassenZeilen, IEnumerable<dynamic> faecherDat)
{
    if (fachZeilen == null || klassenZeilen == null)
        return false;

    // 1. Alle IDs extrahieren, die im aktuellen Fach KEINE Schülergruppe haben
    var relevanteIds = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            var schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            return string.IsNullOrWhiteSpace(schuelergruppe) || schuelergruppe == "?";
        })
        .Select(d => ExtrahiereUnterrichtId(d)?.Trim().Trim('"') ?? "")
        .Where(id => !string.IsNullOrWhiteSpace(id) && id != "?")
        .Distinct()
        .ToList();

    if (!relevanteIds.Any())
        return false;

    // Das aktuelle Fach ermitteln (BaseFach)
    var aktuellesRawFach = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "")
        .FirstOrDefault(f => !string.IsNullOrWhiteSpace(f));

    var aktuellesBaseFach = GetBaseSubjectWithCourseType(aktuellesRawFach, faecherDat);

    // 2. Prüfen, ob mindestens eine dieser IDs in der Klasse mit einem ANDEREN Fach vorkommt
    return klassenZeilen
        .OfType<IDictionary<string, object>>()
        .Any(d => 
        {
            var id = ExtrahiereUnterrichtId(d)?.Trim().Trim('"') ?? "";
            if (!relevanteIds.Contains(id))
                return false;

            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

            // Gekoppelt = Selbe ID, aber anderes Fach
            return !baseFach.Equals(aktuellesBaseFach, StringComparison.OrdinalIgnoreCase);
        });
}

 private List<string> GetFachLehIdStunden(IEnumerable<object> fachZeilen, IEnumerable<object> faecherDat)
{
    if (fachZeilen == null)
        return new List<string>();

    var zeilenList = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => new
        {
            Lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "",
            RawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Id = ExtrahiereUnterrichtId(d)?.Trim().Trim('"') ?? "",
            Stunden = ExtrahiereStundenWert(d)
        })
        .Where(x => !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.RawFach) && x.RawFach != "?")
        .ToList();

    if (!zeilenList.Any())
        return new List<string>();

    var result = new List<string>();

    // Nach bereinigtem Fach und Lehrer gruppieren
    var lehrerGruppen = zeilenList
        .GroupBy(x => new 
        { 
            BaseFach = GetBaseSubjectWithCourseType(x.RawFach, faecherDat), 
            x.Lehrer 
        })
        .OrderBy(g => g.Key.BaseFach)
        .ThenBy(g => g.Key.Lehrer);

    foreach (var lg in lehrerGruppen)
    {
        var fach = lg.Key.BaseFach;
        var lehrer = lg.Key.Lehrer;

        // IDs konsolidieren (z. B. "1678,1679" oder einzeln "1675")
        var ids = string.Join(",", lg
            .Select(x => x.Id)
            .Where(id => !string.IsNullOrWhiteSpace(id) && id != "?")
            .Distinct()
            .OrderBy(id => id));

        // Gesamtstunden des Lehrers in diesem Fach berechnen
        var gesamtStunden = lg.Sum(x => x.Stunden);
        var stundenText = gesamtStunden == 1 ? "1 Stunde" : $"{gesamtStunden} Stunden";

        // Klammer-Inhalt aufbauen (id; X Stunden)
        var idUndStundenPart = string.IsNullOrEmpty(ids)
            ? $"({stundenText})"
            : $"({ids}; {stundenText})";

        // Alle Schülergruppen des Lehrers sammeln und alphabetisch verknüpfen
        var schuelergruppen = lg
            .Select(x => x.Schuelergruppe)
            .Where(sg => !string.IsNullOrWhiteSpace(sg) && sg != "?")
            .Distinct()
            .OrderBy(sg => sg)
            .ToList();

        var gruppenJoined = schuelergruppen.Any() 
            ? "-" + string.Join("-", schuelergruppen) 
            : "";

        result.Add($"{fach}-{lehrer}{idUndStundenPart}{gruppenJoined}");
    }

    return result;
}

 private bool alleZeilenZuDemFachHabenEineSchülergruppe(IEnumerable<object> fachZeilen)
{
    if (fachZeilen == null)
        return false;

    var zeilenList = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "";
            
            return !string.IsNullOrWhiteSpace(rawFach) && rawFach != "?" &&
                   !string.IsNullOrWhiteSpace(lehrer) && lehrer != "?";
        })
        .ToList();

    if (!zeilenList.Any())
        return false;

    return zeilenList.All(d => 
    {
        var schuelergruppe = d.ContainsKey("Field42") 
            ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" 
            : "";

        return !string.IsNullOrWhiteSpace(schuelergruppe) && schuelergruppe != "?";
    });
}

 private string GetFachLehStunden(IEnumerable<object> fachZeilen, IEnumerable<dynamic> faecherDat)
{
    if (fachZeilen == null)
        return string.Empty;

    var zeilenList = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => new
        {
            Lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "",
            RawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Stunden = ExtrahiereStundenWert(d)
        })
        .Where(x => !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.RawFach) && x.RawFach != "?")
        .ToList();

    if (!zeilenList.Any())
        return string.Empty;

    // Bereinigtes Fach ermitteln (einheitlich für alle enthaltenen Zeilen)
    var erstesFach = zeilenList.First().RawFach;
    var baseFach = GetBaseSubjectWithCourseType(erstesFach, faecherDat);

    // Stunden pro Lehrer summieren und alphabetisch ordnen
    var lehrerTeile = zeilenList
        .GroupBy(x => x.Lehrer)
        .OrderBy(g => g.Key)
        .Select(g => 
        {
            var stunden = g.Sum(x => x.Stunden);
            var stundenText = stunden == 1 ? "1 Stunde" : $"{stunden} Stunden";
            return $"{g.Key}({stundenText})";
        });

    return $"{baseFach}-{string.Join("-", lehrerTeile)}";
}

 private bool minEinFachHatKeineSchülergruppe(IEnumerable<object> fachZeilen)
{
    if (fachZeilen == null)
        return false;

    return fachZeilen
        .OfType<IDictionary<string, object>>()
        .Any(d => 
        {
            var schuelergruppe = d.ContainsKey("Field42") 
                ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" 
                : "";

            // Zeile hat keine Schülergruppe, wenn das Feld leer, null oder der Untis-Platzhalter "?" ist
            return string.IsNullOrWhiteSpace(schuelergruppe) || schuelergruppe == "?";
        });
}

 private List<string> GetLehrerFachs(string fach, IEnumerable<object> fachZeilen, IEnumerable<dynamic> faecherDat)
{
    if (string.IsNullOrWhiteSpace(fach) || fachZeilen == null)
        return new List<string>();

    var gesuchtesFach = GetBaseSubjectWithCourseType(fach.Trim().Trim('"'), faecherDat);

    return fachZeilen
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

            return baseFach.Equals(gesuchtesFach, StringComparison.OrdinalIgnoreCase);
        })
        .Select(d => ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "")
        .Where(l => !string.IsNullOrWhiteSpace(l) && l != "?")
        .Distinct()
        .OrderBy(l => l)
        .Select(l => $"{gesuchtesFach}-{l}")
        .ToList();
}

 private List<string> GetLehrerDesFachs(string fach, IEnumerable<object> fachZeilen, IEnumerable<dynamic> faecherDat)
{
    if (string.IsNullOrWhiteSpace(fach) || fachZeilen == null)
        return new List<string>();

    var gesuchtesFach = GetBaseSubjectWithCourseType(fach.Trim().Trim('"'), faecherDat);

    return fachZeilen
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

            return baseFach.Equals(gesuchtesFach, StringComparison.OrdinalIgnoreCase);
        })
        .Select(d => ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "")
        .Where(l => !string.IsNullOrWhiteSpace(l) && l != "?")
        .Distinct()
        .OrderBy(l => l)
        .ToList();
}

 private List<int> GetBeteiligteIds(IEnumerable<object> fachZeilen)
{
    if (fachZeilen == null)
        return new List<int>();

    return fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => ExtrahiereUnterrichtId(d))
        .Select(rawId => int.TryParse(rawId, out var id) ? (int?)id : null)
        .Where(id => id.HasValue)
        .Select(id => id.Value)
        .Distinct()
        .OrderBy(id => id)
        .ToList();
}

 private List<string> getFachUnterrichte(IEnumerable<object> fachZeilen, IEnumerable<dynamic> faecherDat)
{
    if (fachZeilen == null)
        return new List<string>();

    var zeilenList = fachZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => new
        {
            Record = d,
            Lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "",
            RawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Id = ExtrahiereUnterrichtId(d)?.Trim().Trim('"') ?? "",
            Stunden = ExtrahiereStundenWert(d)
        })
        .Where(x => !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.RawFach) && x.RawFach != "?")
        .ToList();

    if (!zeilenList.Any())
        return new List<string>();

    var result = new List<string>();

    // Nach bereinigtem Fach (BaseFach) gruppieren
    var faecherGruppen = zeilenList.GroupBy(x => GetBaseSubjectWithCourseType(x.RawFach, faecherDat));

    foreach (var fachGruppe in faecherGruppen)
    {
        var fach = fachGruppe.Key;

        // Prüfen, ob mindestens eine Zeile KEINE Schülergruppe hat (oder nur empty / "?")
        bool hatZeileOhneGruppe = fachGruppe.Any(x => string.IsNullOrWhiteSpace(x.Schuelergruppe) || x.Schuelergruppe == "?");

        if (hatZeileOhneGruppe)
        {
            // FALL 1: Mindestens eine Zeile ohne Schülergruppe
            var lehrerListe = fachGruppe.Select(x => x.Lehrer).Distinct().ToList();

            if (lehrerListe.Count == 1)
            {
                // Nur 1 Lehrer ohne Schülergruppe -> ID wird unterdrückt
                var lehrer = lehrerListe.First();
                var gesamtStunden = fachGruppe.Where(x => x.Lehrer == lehrer).Sum(x => x.Stunden);
                var stundenText = gesamtStunden == 1 ? "1 Stunde" : $"{gesamtStunden} Stunden";

                result.Add($"{fach}-{lehrer}({stundenText})");
            }
            else
            {
                // Mehrere Lehrer ohne Schülergruppe -> IDs unterdrückt, Lehrer alphabetisch
                var lehrerTeile = fachGruppe
                    .GroupBy(x => x.Lehrer)
                    .OrderBy(g => g.Key)
                    .Select(g => 
                    {
                        var stunden = g.Sum(x => x.Stunden);
                        var stundenText = stunden == 1 ? "1 Stunde" : $"{stunden} Stunden";
                        return $"{g.Key}({stundenText})";
                    });

                result.Add($"{fach}-{string.Join("-", lehrerTeile)}");
            }
        }
        else
        {
            // ALLE Zeilen haben Schülergruppen
            // Prüfen auf gemeinsame Schülergruppen (Fall 3)
            var gruppenMapping = fachGruppe
                .GroupBy(x => x.Schuelergruppe)
                .Select(g => new
                {
                    Schuelergruppe = g.Key,
                    LehrerListe = g.Select(x => x.Lehrer).Distinct().OrderBy(l => l).ToList()
                })
                .ToList();

            bool hatGemeinsameGruppe = gruppenMapping.Any(g => g.LehrerListe.Count > 1);

            if (hatGemeinsameGruppe)
            {
                // FALL 3: Mehrere Lehrer teilen sich dieselbe Schülergruppe
                foreach (var g in gruppenMapping.OrderBy(x => x.Schuelergruppe))
                {
                    var lehrerInfoTeile = fachGruppe
                        .Where(x => x.Schuelergruppe == g.Schuelergruppe)
                        .GroupBy(x => x.Lehrer)
                        .OrderBy(lg => lg.Key)
                        .Select(lg => 
                        {
                            var lehrer = lg.Key;
                            var details = lg.Select(x => 
                            {
                                var stundenText = x.Stunden == 1 ? "1 Stunde" : $"{x.Stunden} Stunden";
                                return string.IsNullOrEmpty(x.Id) ? stundenText : $"{x.Id}; {stundenText}";
                            });

                            return $"{lehrer}({string.Join(", ", details)})";
                        });

                    result.Add($"{fach}-{string.Join("-", lehrerInfoTeile)}-{g.Schuelergruppe}");
                }
            }
            else
            {
                // FALL 2: Konsolidierung PRO LEHRER (Maximal 1 Zeile pro Fach und Lehrer)
                var lehrerGruppen = fachGruppe
                    .GroupBy(x => x.Lehrer)
                    .OrderBy(g => g.Key);

                foreach (var lg in lehrerGruppen)
                {
                    var lehrer = lg.Key;

                    // Details aus allen Unterrichen des Lehrers zusammenfassen (z.B. "1678; 1 Stunde, 1679; 1 Stunde")
                    var detailsTeile = lg.Select(x => 
                    {
                        var stundenText = x.Stunden == 1 ? "1 Stunde" : $"{x.Stunden} Stunden";
                        return string.IsNullOrEmpty(x.Id) ? stundenText : $"{x.Id}; {stundenText}";
                    });
                    var detailsJoined = string.Join(", ", detailsTeile);

                    // Alle verschiedenen Schülergruppen des Lehrers sammeln und verknüpfen
                    var alleGruppen = lg
                        .Select(x => x.Schuelergruppe)
                        .Distinct()
                        .OrderBy(sg => sg)
                        .ToList();

                    var gruppenJoined = string.Join("-", alleGruppen);
                    result.Add($"{fach}-{lehrer}({detailsJoined})-{gruppenJoined}");
                }
            }
        }
    }

    return result.Distinct().OrderBy(r => r).ToList();
}

 private List<IDictionary<string, object>> GetZeilenZuKlassenFaechern(string klasse, string fach, IEnumerable<object> klassenZeilen, IEnumerable<dynamic> faecherDat)
{
    if (string.IsNullOrWhiteSpace(klasse) || string.IsNullOrWhiteSpace(fach) || klassenZeilen == null)
        return new List<IDictionary<string, object>>();

    var gesuchteKlasse = klasse.Trim().Trim('"');
    var gesuchtesFach = GetBaseSubjectWithCourseType(fach.Trim().Trim('"'), faecherDat);

    return klassenZeilen
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            var k = d.ContainsKey("Field5") ? d["Field5"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

            return k.Equals(gesuchteKlasse, StringComparison.OrdinalIgnoreCase) &&
                   baseFach.Equals(gesuchtesFach, StringComparison.OrdinalIgnoreCase);
        })
        .ToList();
}

 private List<IDictionary<string, object>> GetZeilenZuKlasse(string klasse, IEnumerable<object> alleGpu002)
{
    if (string.IsNullOrWhiteSpace(klasse) || alleGpu002 == null)
        return new List<IDictionary<string, object>>();

    var gesuchteKlasse = klasse.Trim().Trim('"');

    return alleGpu002
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            // 1. Prüfung auf Klasse
            var k = d.ContainsKey("Field5") ? d["Field5"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            if (!k.Equals(gesuchteKlasse, StringComparison.OrdinalIgnoreCase))
                return false;

            // 2. Prüfung auf Stundenwert > 0
            var stunden = ExtrahiereStundenWert(d);
            return stunden > 0;
        })
        .ToList();
}

 private List<string> getUnterrichte(IEnumerable<object> beteiligteZeilen, IEnumerable<dynamic> faecherDat)
{
    if (beteiligteZeilen == null)
        return new List<string>();

    var zeilenList = beteiligteZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => new
        {
            Lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "",
            RawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "",
            Schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : ""
        })
        .Where(x => !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.RawFach) && x.RawFach != "?")
        .ToList();

    if (!zeilenList.Any())
        return new List<string>();

    var result = new List<string>();

    // Gruppieren nach bereinigtem Fach (BaseFach)
    var faecherGruppen = zeilenList
        .GroupBy(x => GetBaseSubjectWithCourseType(x.RawFach,faecherDat));

    foreach (var fachGruppe in faecherGruppen)
    {
        var fach = fachGruppe.Key;

        // Prüfen, ob mindestens eine Zeile KEINE Schülergruppe hat (oder nur empty / "?")
        bool hatZeileOhneGruppe = fachGruppe.Any(x => string.IsNullOrWhiteSpace(x.Schuelergruppe) || x.Schuelergruppe == "?");

        if (hatZeileOhneGruppe)
        {
            // FALL 1: Mindestens eine Zeile ohne Schülergruppe
            // -> Fach-Leh1-Leh2 (alle Lehrer alphabetisch angehängt)
            var alleLehrer = fachGruppe
                .Select(x => x.Lehrer)
                .Distinct()
                .OrderBy(l => l)
                .ToList();

            if (alleLehrer.Any())
            {
                result.Add($"{fach}-{string.Join("-", alleLehrer)}");
            }
        }
        else
        {
            // FALL 2 & 3: Alle Zeilen haben Schülergruppen
            // Gruppieren nach Schülergruppe, um gemeinsame Gruppen mehrerer Lehrer zu finden (Fall 3)
            var gruppenMapping = fachGruppe
                .GroupBy(x => x.Schuelergruppe)
                .Select(g => new
                {
                    Schuelergruppe = g.Key,
                    LehrerListe = g.Select(x => x.Lehrer).Distinct().OrderBy(l => l).ToList()
                })
                .ToList();

            // Prüfen, ob es Schülergruppen gibt, die sich MEHRERE Lehrer teilen (Fall 3)
            bool hatGemeinsameGruppe = gruppenMapping.Any(g => g.LehrerListe.Count > 1);

            if (hatGemeinsameGruppe)
            {
                // FALL 3: Zwei/mehrere Lehrer teilen sich dieselbe Schülergruppe
                // -> Fach-Leh1-Leh2-Schülergruppe
                foreach (var g in gruppenMapping.OrderBy(x => x.Schuelergruppe))
                {
                    var lehrerJoined = string.Join("-", g.LehrerListe);
                    result.Add($"{fach}-{lehrerJoined}-{g.Schuelergruppe}");
                }
            }
            else
            {
                // FALL 2: Jeder Lehrer hat seine eigenen Schülergruppen
                // -> Fach-Leh1-Schülergruppe1VonLeh1-Schülergruppe2VonLeh1
                var lehrerGruppen = fachGruppe
                    .GroupBy(x => x.Lehrer)
                    .OrderBy(g => g.Key);

                foreach (var lg in lehrerGruppen)
                {
                    var lehrer = lg.Key;
                    var schuelergruppen = lg
                        .Select(x => x.Schuelergruppe)
                        .Distinct()
                        .OrderBy(sg => sg)
                        .ToList();

                    var gruppenJoined = string.Join("-", schuelergruppen);
                    result.Add($"{fach}-{lehrer}-{gruppenJoined}");
                }
            }
        }
    }

    return result.Distinct().OrderBy(r => r).ToList();
}

 private List<string> GetSchuelergruppen(IEnumerable<object> beteiligteZeilen)
{
    if (beteiligteZeilen == null)
        return new List<string>();

    return beteiligteZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "")
        .Where(sg => !string.IsNullOrWhiteSpace(sg) && sg != "?")
        .Distinct()
        .OrderBy(sg => sg)
        .ToList();
}

 private List<string> GetLehrerFachSchuelergruppen(IEnumerable<object> beteiligteZeilen, IEnumerable<dynamic> faecherDat)
{
    if (beteiligteZeilen == null)
        return new List<string>();

    return beteiligteZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => 
        {
            var lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "";
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var bereinigtesFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);
            var schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";

            return new 
            {
                Lehrer = lehrer,
                Fach = bereinigtesFach,
                Schuelergruppe = schuelergruppe
            };
        })
        // Gültige Datensätze filtern (keine leeren Lehrer/Fächer oder Untis-Platzhalter)
        .Where(x => !string.IsNullOrWhiteSpace(x.Lehrer) && x.Lehrer != "?" &&
                    !string.IsNullOrWhiteSpace(x.Fach) && x.Fach != "?")
        // Gruppieren nach Lehrer und bereinigtem Fach
        .GroupBy(x => new { x.Lehrer, x.Fach })
        .Select(g => 
        {
            // Eindeutige Schülergruppen sammeln
            var gruppen = g
                .Select(x => x.Schuelergruppe)
                .Where(sg => !string.IsNullOrWhiteSpace(sg) && sg != "?")
                .Distinct()
                .OrderBy(sg => sg)
                .ToList();

            string gruppenJoined = gruppen.Any() ? string.Join("-", gruppen) : "";

            return string.IsNullOrEmpty(gruppenJoined)
                ? $"{g.Key.Lehrer}-{g.Key.Fach}"
                : $"{g.Key.Lehrer}-{g.Key.Fach}-{gruppenJoined}";
        })
        .Distinct()
        .OrderBy(eintrag => eintrag)
        .ToList();
}

 private List<IDictionary<string, object>> GetZeilenZuKlasseUndFach(List<int> ids, IEnumerable<object> alleGpu002)
{
    if (ids == null || !ids.Any() || alleGpu002 == null)
        return new List<IDictionary<string, object>>();

    // Schnellere Abfrage via HashSet
    var targetIds = ids.ToHashSet();

    return alleGpu002
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            var rawId = ExtrahiereUnterrichtId(d);
            return int.TryParse(rawId, out var id) && targetIds.Contains(id);
        })
        .ToList();
}

 private List<string> GetFaecherBereinigt(List<string> faecherRoh, IEnumerable<dynamic> faecherDat)
 {
  List<string> faecherBereinigt = new List<string>();
  
  foreach(var f in faecherRoh)
        {
            var ff = GetBaseSubjectWithCourseType(f, faecherDat);

            if(!faecherBereinigt.Contains(ff))
            {
                faecherBereinigt.Add(ff);
            }
        }
        return faecherBereinigt;
 }

 private List<IDictionary<string, object>> lehrerZeilenDerKlasseInDiesemFach(string f, IEnumerable<dynamic> zeilen, IEnumerable<dynamic> faecherDat)
{
    if (string.IsNullOrWhiteSpace(f) || zeilen == null)
        return new List<IDictionary<string, object>>();

    // Such-Fach säubern und Basis-Fach ermitteln
    var gesuchtesBaseFach = GetBaseSubjectWithCourseType(f.Trim().Trim('"'), faecherDat);

    return zeilen
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
        {
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);
            var lehrer = ExtrahiereLehrer(d);

            return baseFach == gesuchtesBaseFach && 
                   !string.IsNullOrWhiteSpace(lehrer) && 
                   lehrer != "?";
        })
        .ToList();
}

 private List<int> GetBeteiligteIdsInDiesemFach(string f, IEnumerable<dynamic> zeilen, IEnumerable<dynamic> faecherDat)
{
    if (string.IsNullOrWhiteSpace(f) || zeilen == null)
        return new List<int>();

    // Fach-String säubern und Base-Fach ermitteln
    var gesuchtesBaseFach = GetBaseSubjectWithCourseType(f.Trim().Trim('"'), faecherDat);

    return zeilen
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
        {
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            return GetBaseSubjectWithCourseType(rawFach, faecherDat) == gesuchtesBaseFach;
        })
        .Select(d => ExtrahiereUnterrichtId(d))
        .Where(idStr => !string.IsNullOrWhiteSpace(idStr))
        .Select(idStr => int.TryParse(idStr, out var parsed) ? parsed : (int?)null)
        .Where(id => id.HasValue)
        .Select(id => id.Value)
        .Distinct()
        .OrderBy(id => id)
        .ToList();
}

 private List<string> getUnterschiedlicheLehrerSchülergruppen(List<int> beteiligteIds, IEnumerable<dynamic> alleGpu002, IEnumerable<dynamic> faecherDat)
{
    if (beteiligteIds == null || !beteiligteIds.Any())
        return new List<string>();

    var targetIds = beteiligteIds.ToHashSet();

    return alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
        {
            var rawId = ExtrahiereUnterrichtId(d);
            return int.TryParse(rawId, out var id) && targetIds.Contains(id);
        })
        .Select(d => 
        {
            var lehrer = d.ContainsKey("Field6") ? d["Field6"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var schildFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";

            // Gültigkeitsprüfung: Alle drei Werte müssen vorhanden sein und dürfen kein Untis-Platzhalter ("?") sein
            if (string.IsNullOrWhiteSpace(lehrer) || lehrer == "?" || 
                string.IsNullOrWhiteSpace(schildFach) || schildFach == "?" ||
                string.IsNullOrWhiteSpace(schuelergruppe))
            {
                return null;
            }

            return $"{lehrer}-{GetBaseSubjectWithCourseType(schildFach, faecherDat)}-{schuelergruppe}";
        })
        .Where(eintrag => eintrag != null)
        .Distinct()
        .OrderBy(eintrag => eintrag)
        .ToList();
}

 private List<string> getUnterschiedlicheFächerInRohfaechern(List<string> rohfaecher, IEnumerable<dynamic> faecherDat)
 {
    var schildfaecher = new List<string>();

        foreach(var r in rohfaecher)
        {
            var schildfach = GetBaseSubjectWithCourseType(r, faecherDat);

            if (!schildfaecher.Contains(schildfach))
            {
                schildfaecher.Add(schildfach);     
            }            
        }
        return schildfaecher;
 }

 private bool rohfaecherEnthältNichtSchildFächer(dynamic schildFachOhneSuffix, List<string> rohfaecher)
 {
  foreach(var r in rohfaecher)
        {
            if(!r.Contains(schildFachOhneSuffix))
                return false;
        }
        return true;
 }

 private string GetIdsMitUnterrichtsgruppenString(List<int> beteiligteIds, IEnumerable<dynamic> alleGpu002)
{
    if (beteiligteIds == null || !beteiligteIds.Any())
        return "-";

    // 1. GPU002-Datensätze typensicher indizieren (ID -> Unterrichtsgruppe aus Field12)
    var dictIdZuUGruppe = alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Select(d => new 
        { 
            RawId = ExtrahiereUnterrichtId(d), 
            UGruppe = GetUGroupFromRecord(d) 
        })
        .Where(x => int.TryParse(x.RawId, out _))
        .GroupBy(x => int.Parse(x.RawId))
        .ToDictionary(
            g => g.Key, 
            g => g.First().UGruppe
        );

    // 2. Formatierung für jede ID aufbauen: ID(U-Gruppe)
    var formatiert = beteiligteIds.Select(id => 
    {
        var uGruppe = dictIdZuUGruppe.TryGetValue(id, out var g) ? g : "";
        return $"{id}({uGruppe})";
    });

    return string.Join(" ", formatiert);
}

 private List<string> GetBeteiligteUnterrichtsgruppen(List<int> beteiligteIdsSortiert, IEnumerable<dynamic> alleGpu002)
{
    if (beteiligteIdsSortiert == null || !beteiligteIdsSortiert.Any())
        return new List<string>();

    // Typensicherer Abgleich über HashSet<int>
    var targetIds = beteiligteIdsSortiert.ToHashSet();

    return alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
        {
            var rawId = ExtrahiereUnterrichtId(d);
            return int.TryParse(rawId, out var id) && targetIds.Contains(id);
        })
        .Select(d => GetUGroupFromRecord(d)?.Trim() ?? "")
        .Where(uGruppe => !string.IsNullOrWhiteSpace(uGruppe) && uGruppe != "0")
        .Distinct()
        .OrderBy(uGruppe => uGruppe)
        .ToList();
}

private string GetUGroupFromRecord(IDictionary<string, object> dictRecord)
{
    if (dictRecord == null)
        return "";

    // Unterrichtsgruppe direkt aus Field12 auslesen
    if (dictRecord.ContainsKey("Field12") && dictRecord["Field12"] != null)
    {
        return dictRecord["Field12"]?.ToString()?.Trim()?.Trim('"') ?? "";
    }

    return "";
}

// Überladung für dynamic-Objekte



 /*
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
 // -------------------------------------------------------------------------------------
 // Gruppierung für die finale Aggregation:
 // Unterscheidet strikt zwischen Klassenunterricht (pro Klasse/Lehrer getrennt)
 // und Kursen (zusammengefasste Klassen/Schülergruppen).
 // -------------------------------------------------------------------------------------
 // Strikte Gruppierung für die finale Aggregation
 var finalGroups = records
     .GroupBy(r => 
     {
         var rawFach = r.ContainsKey("Field7") ? r["Field7"]?.ToString() ?? "" : "";
         var baseFach = GetBaseSubjectWithCourseType(rawFach);
         var lehrer = GetTeacherFromRecord(r);
         var klasse = r.ContainsKey("Field5") ? r["Field5"]?.ToString() ?? "" : "";
         var studentGroup = GetStudentGroupFromRecord(r);

         // 1. Hat der Datensatz eine explizite Schülergruppe? -> Echter Kurs
         if (!string.IsNullOrEmpty(studentGroup))
         {
             return $"{baseFach}_KURS_SG_{studentGroup}";
         }

         // Ermitteln, wie viele Lehrer für dieses Fach in dieser Klasse unterrichten
         var teachersInClassForSubject = records
             .Where(x => (x.ContainsKey("Field5") ? x["Field5"]?.ToString() ?? "" : "") == klasse &&
                         GetBaseSubjectWithCourseType(x.ContainsKey("Field7") ? x["Field7"]?.ToString() ?? "" : "") == baseFach)
             .Select(x => GetTeacherFromRecord(x))
             .Where(t => !string.IsNullOrEmpty(t))
             .Distinct()
             .Count();

         // 2. Mehrere Lehrer in derselben Klasse & Fach? -> Kurs (Teamteaching / Differenzierung)
         if (teachersInClassForSubject > 1)
         {
             return $"{baseFach}_KURS_{klasse}";
         }

         // 3. Reiner Klassenunterricht (Einzelner Lehrer, keine Schülergruppe, 1 Klasse)
         bool istKlassenunterricht = klassenUnterrichtSet.Contains($"{klasse}_{baseFach}");

         if (istKlassenunterricht)
         {
             return $"{baseFach}_KU_{lehrer}_{klasse}";
         }

         // 4. Fallback für sonstige Kurse/Koppeln pro Klasse
         return $"{baseFach}_KURS_{klasse}_{lehrer}";
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
     // Statt group.Key (welcher z. B. "SP_KU_HES_AGG26A" lautet)
     // extrahieren wir das echte BaseFach aus den Datensätzen der Gruppe
     var rawRecordsInGroup = group.ToList();
     var sampleRecord = rawRecordsInGroup.First();
     var sampleRawFach = sampleRecord.ContainsKey("Field7") ? sampleRecord["Field7"]?.ToString() ?? "" : "";
     var baseFach = GetBaseSubjectWithCourseType(sampleRawFach);

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
         Jahrgaenge = new List<string>(),//jahrgaengeList,

         // Schülergruppen & gefilterte Schülerobjekte
         Schülergruppe = schuelerGruppenStr,
         Students = m.IStudents.Filter(configuration, zweck, klassen.FirstOrDefault() ?? "", schuelerGruppenStr, studentgroupStudents)
     };
     if(!string.IsNullOrEmpty(unterricht.KursBez))
         unterricht.Kursart = unterricht.GetKursart(configuration, kurseDat, unterricht.Fach, primaryTeacher, ids[0]);

     // Unterricht zum aktuellen Container/Liste hinzufügen
     this.Add(unterricht);

 }
 */
 // Rendern der Spectre.Console Tabelle auf der Konsole







 private int? GetId(dynamic record)
{
    if (record == null)
        return null;

    // DLR-Sicherheit: Casten auf IDictionary<string, object> zur Vermeidung von RuntimeBinderExceptions
    var dictRecord = record as IDictionary<string, object>;
    if (dictRecord == null)
        return null;

    // Unterrichts-ID extrahieren (z. B. aus Field1 oder via Helper-Methode)
    string rawId = ExtrahiereUnterrichtId(dictRecord);

    if (string.IsNullOrWhiteSpace(rawId))
        return null;

    // Sicheres Parsen nach int
    if (int.TryParse(rawId.Trim(), out var parsedId))
    {
        return parsedId;
    }

    return null;
}

 private DateTime GetFrühestesDatum(List<int> beteiligteIds, IEnumerable<dynamic> alleGpu002)
{
    if (beteiligteIds == null || !beteiligteIds.Any())
        return DateTime.MinValue;

    var targetIds = beteiligteIds.ToHashSet();

    var daten = alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
        {
            var rawId = ExtrahiereUnterrichtId(d);
            return int.TryParse(rawId, out var id) && targetIds.Contains(id);
        })
        .Select(d => d.ContainsKey("Field15") ? jjjjmmddNachDateTime(d["Field15"]?.ToString()) : null)
        .Where(dt => dt.HasValue)
        .Select(dt => dt.Value)
        .ToList();

    return daten.Any() ? daten.Min() : DateTime.MinValue;
}

private DateTime GetSpätestesDatum(List<int> beteiligteIds, IEnumerable<dynamic> alleGpu002)
{
    if (beteiligteIds == null || !beteiligteIds.Any())
        return DateTime.MaxValue;

    var targetIds = beteiligteIds.ToHashSet();

    var daten = alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
        {
            var rawId = ExtrahiereUnterrichtId(d);
            return int.TryParse(rawId, out var id) && targetIds.Contains(id);
        })
        .Select(d => d.ContainsKey("Field16") ? jjjjmmddNachDateTime(d["Field16"]?.ToString()) : null)
        .Where(dt => dt.HasValue)
        .Select(dt => dt.Value)
        .ToList();

    return daten.Any() ? daten.Max() : DateTime.MaxValue;
}

 private List<string> GetFaecherRoh(IEnumerable<object> beteiligteZeilen)
{
    if (beteiligteZeilen == null)
        return new List<string>();

    return beteiligteZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "")
        .Where(rawFach => !string.IsNullOrWhiteSpace(rawFach) && rawFach != "?")
        .Distinct()
        .OrderBy(rawFach => rawFach)
        .ToList();
}

 private bool fälltInDieStatistik(List<int> beteiligteIds, DateTime statistikDatum, IEnumerable<dynamic> alleGpu002)
{
    if (beteiligteIds == null || !beteiligteIds.Any())
        return false;

    // IDs direkt als Integer im HashSet ablegen
    var targetIds = beteiligteIds.ToHashSet();

    // Alle Datensätze filtern, die zu den beteiligten IDs gehören
    var relevanteZeilen = alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
        {
            var rawId = ExtrahiereUnterrichtId(d);
            return int.TryParse(rawId, out var id) && targetIds.Contains(id);
        })
        .ToList();

    if (!relevanteZeilen.Any())
        return false;

    DateTime? fruehestesVon = null;
    DateTime? spaetestesBis = null;

    foreach (var dict in relevanteZeilen)
    {
        // Field15: Von-Datum (Format: YYYYMMDD)
        if (dict.ContainsKey("Field15") && dict["Field15"] != null)
        {
            var dtVon = jjjjmmddNachDateTime(dict["Field15"]?.ToString());
            if (dtVon.HasValue && (!fruehestesVon.HasValue || dtVon.Value < fruehestesVon.Value))
            {
                fruehestesVon = dtVon.Value;
            }
        }

        // Field16: Bis-Datum (Format: YYYYMMDD)
        if (dict.ContainsKey("Field16") && dict["Field16"] != null)
        {
            var dtBis = jjjjmmddNachDateTime(dict["Field16"]?.ToString());
            if (dtBis.HasValue && (!spaetestesBis.HasValue || dtBis.Value > spaetestesBis.Value))
            {
                spaetestesBis = dtBis.Value;
            }
        }
    }

    // Wenn keine gültigen Datumsangaben vorhanden sind, im Zweifel true
    if (!fruehestesVon.HasValue || !spaetestesBis.HasValue)
        return true;

    // Prüfen, ob das Statistikdatum im Gesamtraum liegt
    return statistikDatum >= fruehestesVon.Value && statistikDatum <= spaetestesBis.Value;
}

 private string GetKursbezeichnung(string kursleiter, List<int> beteiligteIds, bool istKurs)
{
    if(!istKurs) return "";

    var lehrerKuerzel = kursleiter?.Trim() ?? "";

    // 1. Ungültige/negative IDs herausfiltern (falls vorhanden)
    var gueltigeIds = (beteiligteIds ?? new List<int>())
        .Where(id => id > 0)
        .ToList();

    // 2. Format zusammenbauen: "LEHRER-ID1-ID2-..."
    string name = lehrerKuerzel;

    if (gueltigeIds.Any())
    {
        var idsString = string.Join("-", gueltigeIds);
        name = string.IsNullOrEmpty(lehrerKuerzel) ? idsString : $"{lehrerKuerzel}-{idsString}";
    }
    return name;
}

private string GetKursbezeichnungGekürzt(string kursleiter, List<int> beteiligteIds, bool istKurs)
{
    if(!istKurs) return "";

    var lehrerKuerzel = kursleiter?.Trim() ?? "";

    // 1. Ungültige/negative IDs herausfiltern (falls vorhanden)
    var gueltigeIds = (beteiligteIds ?? new List<int>())
        .Where(id => id > 0)
        .ToList();

    // 2. Format zusammenbauen: "LEHRER-ID1-ID2-..."
    string name = lehrerKuerzel;

    if (gueltigeIds.Any())
    {
        var idsString = string.Join("-", gueltigeIds);
        name = string.IsNullOrEmpty(lehrerKuerzel) ? idsString : $"{lehrerKuerzel}-{idsString}";
    }

    // 3. Auf maximal 20 Zeichen begrenzen
    if (name.Length > 20)
    {
        name = name.Substring(0, 20);
    }

    return name;
}

private List<string> GetKlassen(IEnumerable<object> beteiligteZeilen)
{
    if (beteiligteZeilen == null)
        return new List<string>();

    return beteiligteZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => d.ContainsKey("Field5") ? d["Field5"]?.ToString()?.Trim()?.Trim('"') ?? "" : "")
        .Where(k => !string.IsNullOrWhiteSpace(k) && k != "0" && k != "?")
        .Distinct()
        .OrderBy(k => k) // Alphabetisch aufsteigend sortieren
        .ToList();
}

 private string GetRawFach(dynamic record)
{
    var dict = (IDictionary<string, object>)record;
    
    // Feldkandidaten für das Fach in Untis GPU002 (Field7 = Fach-Kürzel)
    string[] candidateFields = { "Field7", "Field8" };

    foreach (var field in candidateFields)
    {
        if (dict.ContainsKey(field) && dict[field] != null)
        {
            var val = dict[field]?.ToString()?.Trim()?.Trim('"');
            if (!string.IsNullOrEmpty(val))
                return val;
        }
    }

    return string.Empty;
}

 private List<int> getWeitereLehrerWochenstunden(string kursleiter, List<string> weitereLehrer, dynamic record, IEnumerable<dynamic> alleGpu002, IEnumerable<dynamic> faecherDat)
{
    var stundenListe = new List<int>();

    if (weitereLehrer == null || !weitereLehrer.Any())
        return stundenListe;

    var dictRecord = (IDictionary<string, object>)record;
    var klasse = dictRecord.ContainsKey("Field5") ? dictRecord["Field5"]?.ToString() ?? "" : "";
    var rawFach = dictRecord.ContainsKey("Field7") ? dictRecord["Field7"]?.ToString() ?? "" : "";
    var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

    // Alle relevanten Datensätze für Klasse und Fach vorfiltern (IDictionary verhindert RuntimeBinderException)
    var relevanteZeilen = alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
            (d.ContainsKey("Field5") ? d["Field5"]?.ToString() ?? "" : "") == klasse &&
            GetBaseSubjectWithCourseType(d.ContainsKey("Field7") ? d["Field7"]?.ToString() ?? "" : "", faecherDat) == baseFach
        )
        .ToList();

    // Für jeden weiteren Lehrer die Wochenstunden als int aufsummieren
    foreach (var lehrer in weitereLehrer)
    {
        // Wandelt das Ergebnis der Summe explizit in int um
        int stunden = (int)relevanteZeilen
            .Where(d => ExtrahiereLehrer(d) == lehrer)
            .Sum(d => (double)ExtrahiereStundenWert(d));

        stundenListe.Add(stunden);
    }

    return stundenListe;
}

 private List<int> GetBeteiligteIds(dynamic record, IEnumerable<dynamic> alleGpu002, IEnumerable<dynamic> faecherDat)
{
    var dictRecord = (IDictionary<string, object>)record;

    var klasse = dictRecord.ContainsKey("Field5") ? dictRecord["Field5"]?.ToString() ?? "" : "";
    var rawFach = dictRecord.ContainsKey("Field7") ? dictRecord["Field7"]?.ToString() ?? "" : "";
    var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

    return alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
            (d.ContainsKey("Field5") ? d["Field5"]?.ToString() ?? "" : "") == klasse &&
            GetBaseSubjectWithCourseType(d.ContainsKey("Field7") ? d["Field7"]?.ToString() ?? "" : "", faecherDat) == baseFach
        )
        .Select(d => ExtrahiereUnterrichtId(d))
        .Where(id => !string.IsNullOrWhiteSpace(id))
        .Select(id => int.TryParse(id, out var parsed) ? parsed : (int?)null)
        .Where(id => id.HasValue)
        .Select(id => id.Value)
        .Distinct()
        .OrderBy(id => id)
        .ToList();
}

#region Hilfsmethode zur ID-Extraktion

private string ExtrahiereUnterrichtId(IDictionary<string, object> d)
{
    // Typische Untis-Felder für die Unterrichts-ID (Field1 oder Field0)
    string[] candidateFields = { "Field1", "Field0" };

    foreach (var field in candidateFields)
    {
        if (d.ContainsKey(field) && d[field] != null)
        {
            var val = d[field]?.ToString()?.Trim()?.Trim('"');
            if (!string.IsNullOrEmpty(val) && val != "0")
                return val;
        }
    }

    return string.Empty;
}

#endregion

 private double GetKursleiterwochenstunden(string kursleiter, IEnumerable<object> beteiligteZeilen)
{
    if (string.IsNullOrWhiteSpace(kursleiter) || beteiligteZeilen == null) 
        return 0.0;

    var gesuchterLehrer = kursleiter.Trim().Trim('"');

    return beteiligteZeilen
        .OfType<IDictionary<string, object>>()
        .Where(d => 
        {
            var lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "";
            return lehrer.Equals(gesuchterLehrer, StringComparison.OrdinalIgnoreCase);
        })
        .Sum(d => ExtrahiereStundenWert(d));
}

#region Hilfsmethode zur Stunden-Extraktion

private double ExtrahiereStundenWert(IDictionary<string, object> d)
{
    // Kandidatenfelder für Wochenstunden in Untis (Field2 / Field11)
    string[] candidateFields = { "Field2", "Field11" };

    foreach (var field in candidateFields)
    {
        if (d.ContainsKey(field) && d[field] != null)
        {
            var rawVal = d[field]?.ToString()?.Trim();
            if (double.TryParse(rawVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var parsed))
            {
                if (parsed <= 0) continue;

                // Untis-Skalierungsfaktoren behandeln (z.B. 100000 = 1.0 h, 200 = 2.0 h)
                if (parsed >= 100000) return parsed / 100000.0;
                if (parsed >= 100) return parsed / 100.0;
                
                return (int)parsed;
            }
        }
    }

    return 0;
}

#endregion

 private List<string> getLehrerKürzelSortiert(IEnumerable<object> beteiligteZeilen)
{
    if (beteiligteZeilen == null)
        return new List<string>();

    return beteiligteZeilen
        .OfType<IDictionary<string, object>>()
        .Select(d => ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "")
        .Where(l => !string.IsNullOrWhiteSpace(l) && l != "?")
        .Distinct()
        .OrderBy(l => l)
        .ToList();
}

 private List<IDictionary<string, object>> lehrerZeilenInDerKlasseInDiesemFach(dynamic record, IEnumerable<dynamic> alleGpu002, IEnumerable<dynamic> faecherDat)
{
    var dictRecord = (IDictionary<string, object>)record;
    
    // Klasse (Field5) und Fach (Field7) aus dem aktuellen Datensatz auslesen
    var klasse = dictRecord.ContainsKey("Field5") ? dictRecord["Field5"]?.ToString() ?? "" : "";
    var rawFach = dictRecord.ContainsKey("Field7") ? dictRecord["Field7"]?.ToString() ?? "" : "";
    var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

    // Alle Zeilen filtern, die dieselbe Klasse und dasselbe Fach haben
    return alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
            (d.ContainsKey("Field5") ? d["Field5"]?.ToString() ?? "" : "") == klasse &&
            GetBaseSubjectWithCourseType(d.ContainsKey("Field7") ? d["Field7"]?.ToString() ?? "" : "", faecherDat) == baseFach &&
            !string.IsNullOrWhiteSpace(ExtrahiereLehrer(d))
        )
        .ToList();
}

 private bool esGibtSchülergruppen(dynamic record, IEnumerable<dynamic> alleGpu002, IEnumerable<dynamic> faecherDat)
{
    var dictRecord = (IDictionary<string, object>)record;
    
    var klasse = dictRecord.ContainsKey("Field5") ? dictRecord["Field5"]?.ToString() ?? "" : "";
    var rawFach = dictRecord.ContainsKey("Field7") ? dictRecord["Field7"]?.ToString() ?? "" : "";
    var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

    // Prüft, ob mindestens ein Datensatz dieser Klasse und dieses Fachs eine Schülergruppe besitzt
    return alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Any(d => 
        {
            var k = d.ContainsKey("Field5") ? d["Field5"]?.ToString() ?? "" : "";
            var f = GetBaseSubjectWithCourseType(d.ContainsKey("Field7") ? d["Field7"]?.ToString() ?? "" : "", faecherDat);
            var sg = ExtrahiereSchuelergruppe(d);

            return k == klasse && f == baseFach && !string.IsNullOrEmpty(sg);
        });
}

#region Hilfsmethode



#endregion

 private bool diesesFachWirdVonMehrerenLehrernInDieserKlasseUnterrichtet(dynamic record, IEnumerable<dynamic> alleGpu002, IEnumerable<dynamic> faecherDat)
{
    var dictRecord = (IDictionary<string, object>)record;
    
    // Klasse (Field5) und Fach (Field7) aus dem aktuellen Datensatz ermitteln
    var klasse = dictRecord.ContainsKey("Field5") ? dictRecord["Field5"]?.ToString() ?? "" : "";
    var rawFach = dictRecord.ContainsKey("Field7") ? dictRecord["Field7"]?.ToString() ?? "" : "";
    var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

    // Alle eindeutigen Lehrer für genau diese Klasse und dieses Fach zählen
    var anzahlLehrer = alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Where(d => 
            (d.ContainsKey("Field5") ? d["Field5"]?.ToString() ?? "" : "") == klasse &&
            GetBaseSubjectWithCourseType(d.ContainsKey("Field7") ? d["Field7"]?.ToString() ?? "" : "", faecherDat) == baseFach
        )
        .Select(d => ExtrahiereLehrer(d))
        .Where(l => !string.IsNullOrWhiteSpace(l))
        .Distinct()
        .Count();

    return anzahlLehrer > 1;
}

#region Hilfsmethode

private string ExtrahiereLehrer(IDictionary<string, object> r)
{
    string[] candidateFields = { "Field6", "Field5" };
    foreach (var field in candidateFields)
    {
        if (r.ContainsKey(field) && r[field] != null)
        {
            var val = r[field]?.ToString()?.Trim()?.Trim('"');
            if (!string.IsNullOrEmpty(val) && val != "0")
                return val;
        }
    }
    return string.Empty;
}

#endregion

 private bool fachMin1xImKlassenverbandInDerKlasse(dynamic record, IEnumerable<dynamic> alleGpu002, IEnumerable<dynamic> faecherDat)
{
    // 'record' ist hier ein einzelner Datensatz
    var dictRecord = (IDictionary<string, object>)record;
    
    var klasse = dictRecord.ContainsKey("Field5") ? dictRecord["Field5"]?.ToString() ?? "" : "";
    var rawFach = dictRecord.ContainsKey("Field7") ? dictRecord["Field7"]?.ToString() ?? "" : "";
    var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat);

    // Prüfen, ob für dieses Fach in der Klasse mindestens 1 Eintrag OHNE Schülergruppe (Field42/43) existiert
    return alleGpu002
        .Select(r => (IDictionary<string, object>)r)
        .Any(d => 
        {
            var k = d.ContainsKey("Field5") ? d["Field5"]?.ToString() ?? "" : "";
            var f = GetBaseSubjectWithCourseType(d.ContainsKey("Field7") ? d["Field7"]?.ToString() ?? "" : "", faecherDat);
            var sg = ExtrahiereSchuelergruppe(d); // Verwendet die Hilfsmethode von vorhin

            return k == klasse && f == baseFach && string.IsNullOrEmpty(sg);
        });
}

#region Lokale Extraktions-Hilfsmethoden

private string ExtrahiereSchuelergruppe(IDictionary<string, object> r)
{
    string[] candidateFields = { "Field42", "Field43" };
    foreach (var field in candidateFields)
    {
        if (r.ContainsKey(field) && r[field] != null)
        {
            var val = r[field]?.ToString()?.Trim()?.Trim('"');

            if (string.IsNullOrEmpty(val) || val == "0" || val == "n") 
                continue;

            // Reine Zahlen ignorieren
            if (double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _))
                continue;

            return val;
        }
    }
    return string.Empty;
}

#endregion






















































public string GetBaseSubjectWithCourseType(string subject, IEnumerable<dynamic> faecherDat)
{
    if (string.IsNullOrWhiteSpace(subject)) 
        return string.Empty;

    string originalSubject = subject.Trim();

    // Hilfsfunktion: Prüft, ob ein Fach-Kürzel in faecherDat (Spalte1: InternKrz) existiert
    bool IsMatch(string krz) => 
        faecherDat.Any(f => string.Equals(f.InternKrz?.ToString(), krz, StringComparison.OrdinalIgnoreCase));

    // Schritt 1: Exakter Match mit dem ursprünglichen Fach
    if (IsMatch(originalSubject))
    {
        return originalSubject;
    }

    // Schritt 2: Möglicherweise vorhandenen Zähler am Ende abschneiden
    // (Entfernt alle Ziffern am Ende, z. B. "D1" -> "D", "E2" -> "E")
    string withoutDigits = originalSubject.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');
    
    if (withoutDigits != originalSubject && IsMatch(withoutDigits))
    {
        return withoutDigits;
    }

    // Schritt 3: Nur bei exakt 4-stelligen Fächern eine Stelle hinten abschneiden
    if (originalSubject.Length == 4)
    {
        string fourToThree = originalSubject.Substring(0, 3);
        if (IsMatch(fourToThree))
        {
            return fourToThree;
        }
    }

    // Schritt 4: Kein Match gefunden -> AnsiConsole Fehlermeldung
    // Option B: Interpolierten String über MarkupInterpolated
    throw new Exception($"Das Untisfach '{originalSubject}' kann keinem Fach in SchILD zugeordnet werden. Bitte korrigieren!");

    //return originalSubject; // Oder string.Empty / null, je nach gewünschtem Fallback
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