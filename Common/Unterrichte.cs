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

            foreach (var klasse in m.IKlassen)
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

                // Extrahiere bereinigte, eindeutige und alphabetisch sortierte Liste von Fächern.
                List<string> faecherRoh = GetFaecherRoh(klassenZeilen);

                // Wandelt Untis-Fächer in SchILD-Basisfächer aus faecherDat um und vermeidet Duplikate.
                // Diese Fächerliste wird in die Leistungsdaten übernommen.
                // Zuordnungsschritte in GetBaseSubjectWithCourseType:
                // 1. Exakter Treffer werden unverändert zurückgegeben(z. B. "D")
                // 2. Ziffern am Ende entfernen (z. B. "D1" -> "D")
                // 3. Auf 3 Zeichen kürzen, wenn 4-stellig (z. B. "MAT1" -> "MAT")
                List<string> faecherBereinigt = GetFaecherBereinigt(faecherRoh, faecherDat, configuration);

                foreach (var fach in faecherBereinigt)
                {
                    // Alle Zeilen des Unterrichts aus der GPU002. Bsp.: Für das Fach GPF wird auch GPF1 usw. herausgesucht
                    var fachZeilen = GetZeilenZuKlassenFaechern(klasse, fach, klassenZeilen, faecherDat, configuration);

                    // Alle Untis-IDs
                    List<int> idsBeteiligteSortiert = GetBeteiligteIds(fachZeilen);

                    // Alle Lehrer, die dieses Fach unterrichten (alphabetisch sortiert)
                    var lehrerDesFachs = GetLehrerDesFachs(fach, fachZeilen, faecherDat, configuration);

                    // Alle Lehrer, die dieses Fach unterrichten (alphabetisch sortiert) mit Fachbezeichnung
                    var lehrerFachs = GetLehrerFachs(fach, fachZeilen, faecherDat, configuration);

                    // Alle Roh-Fächer, die in diesem Fach vorkommen (alphabetisch sortiert) Bsp.: GPF1, GPF2, GPF3, GPF4, GPF5
                    List<string> faecherRohFach = GetFaecherRoh(fachZeilen);
                    bool fachMitSchülergruppe = istFachMitSchülergruppe(fachZeilen);

                    var susDerKlasse = m.IStudents.Where(x => x.Klasse == klasse).ToList();

                    if (fach == "GPF")
                    {
                        string a = "";
                    }

                    var listeIds = new List<List<int>>();
                    var listeLehrer = new List<List<string>>();                    

                    foreach (var s in susDerKlasse)
                    {
                        // Untis-Schülergruppen zu diesem Fach, in denen der Schüler vorkommt.
                        var schuelergruppenDiesesSchuelersinDiesemKlassenfach = GetSchuelergruppenDiesesSusInDiesemFach(s.Id, studentgroupStudents, fachZeilen);

                        // Gib mir alle Lehrer, die fach bei diesem Schüler unterrichten. (alphabetisch sortiert)
                        var lehrerDiesesSchuelersinDiesemKlassenfach = GetLehrerDesFachsDiesesSchülers(fach, fachZeilen, schuelergruppenDiesesSchuelersinDiesemKlassenfach);

                        // Gib mir Untis-IDs, die fach bei diesem Schüler unterrichten. (alphabetisch sortiert)
                        var untisIdsDiesesSchuelersinDiesemKlassenfach = GetUntisIdsDesFachsDiesesSchülers(fach, fachZeilen, schuelergruppenDiesesSchuelersinDiesemKlassenfach);

                        listeIds.Add(untisIdsDiesesSchuelersinDiesemKlassenfach);
                        listeLehrer.Add(lehrerDiesesSchuelersinDiesemKlassenfach);

                        foreach (var sg in schuelergruppenDiesesSchuelersinDiesemKlassenfach)
                        {
                            // Erstelle den Kurs wie folgt als string: "Kurs|Fach|Leh1,Leh2|Id1,Id2|2,1"
                            // wobei Kurs als Wort stehenbleibt. Danach werden alle beteiligten Lehrer (also weitere Zeilen mit derselben ID und demselben fach) am Fachunterricht bei diesem Schüler aufgeführt.
                            //Console.WriteLine(GetUnterrichteString(fachZeilen, sg));
                            var unterricht = GetUnterricht(fachZeilen, faecherRohFach, faecherDat, configuration, sg);

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

                        if (!fachMitSchülergruppe)
                        {
                            var unterricht = GetUnterricht(fachZeilen, faecherRohFach, faecherDat, configuration);

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

                    // Der erste Lehrer, bei dem alle SuS dabei sind, wird Kursleiter. Alle weiteren Lehrer werden als Lehrkräfte mit Stundenangabe hinzugefügt.
                    var kursleiter = lehrerDesFachs.FirstOrDefault();

                }

                foreach (var u in this)
                {
                    if (u.Klassen.Contains(klasse))
                    {
                        Spectre.Console.TableExtensions.AddRow(
                            table,
                            Markup.Escape(u.Fach),
                            Markup.Escape($"{u.Kursleiter}({u.KursleiterWochenstunden}) " + string.Join(" ", u.Lehrkraefte.Zip(u.LehrkraefteWochenstunden, (lk, std) => $"{lk}({std})"))).Trim(),
                            Markup.Escape(string.Join(',', u.Schülergruppe)),
                            Markup.Escape(string.Join(',', u.UnterrichtsIds)),
                            Markup.Escape(string.Join(',', u.Klassen)),
                            Markup.Escape(string.Join(',', u.FaecherRoh)),
                            Markup.Escape(u.Kursart),
                            Markup.Escape(u.KursBez),
                            Markup.Escape($"{u.Von:dd.MM.yyyy} - {u.Bis:dd.MM.yyyy}"),
                            Markup.Escape(u.Students.Count() + "")
                        );
                    }

                    // Bei PUK wird die Kursbezeichnung entfernt, weil eine existierende die Kursbezeichnung 
                    // zu einem Eintrag in Kurse.dat führen wird.
                    if (u.Kursart == "PUK")
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

    private List<int> GetUntisIdsDesFachsDiesesSchülers(string fach, List<IDictionary<string, object>> fachZeilen, List<string> schuelergruppenDiesesSchuelersinDiesemKlassenfach)
    {
        if (fachZeilen == null || schuelergruppenDiesesSchuelersinDiesemKlassenfach == null)
            return new List<int>();

        List<int> untisIds = new List<int>();

        // Durchlaufe alle Fachzeilen und filtere nach dem gesuchten Fach und den Schülergruppen dieses Schülers.
        foreach (var d in fachZeilen)
        {
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var idString = d.ContainsKey("Field1") ? d["Field1"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var id = int.TryParse(idString, out var parsedId) ? parsedId : -1;
            var schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var klasse = d.ContainsKey("Field5") ? d["Field5"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "";

            // Fall 1: Die Schülergruppe ist in den Schülergruppen dieses Schülers enthalten.
            if (schuelergruppenDiesesSchuelersinDiesemKlassenfach.Contains(schuelergruppe))
                if (!untisIds.Contains(id))
                    untisIds.Add(id);

            // Fall 2: Die Schülergruppe ist leer oder nicht vorhanden.
            if (string.IsNullOrWhiteSpace(schuelergruppe))
                if (!untisIds.Contains(id))
                    untisIds.Add(id);

        }
        return untisIds.OfType<int>().Where(id => id > 0).Distinct().OrderBy(id => id).ToList();
    }

    /// <summary>
    /// Gibt eine Liste von Lehrern zurück, die das angegebene Fach bei diesem Schüler unterrichten.
    /// Dazu gehören alle Lehrer, die das Fach in den Zeilen unterrichten, die mit den Schülergruppen dieses Schülers übereinstimmen.
    /// Und es werden auch Lehrer berücksichtigt, die das Fach unterrichten, aber keine Schülergruppe angegeben haben.
    /// </summary>
    /// <param name="fach"></param>
    /// <param name="fachZeilen"></param>
    /// <param name="schuelergruppenDiesesSchuelersinDiesemKlassenfach"></param>
    /// <returns></returns>
    private List<string> GetLehrerDesFachsDiesesSchülers(string fach, List<IDictionary<string, object>> fachZeilen, List<string> schuelergruppenDiesesSchuelersinDiesemKlassenfach)
    {
        if (fachZeilen == null || schuelergruppenDiesesSchuelersinDiesemKlassenfach == null)
            return new List<string>();

        List<string> lehrerListe = new List<string>();

        // Durchlaufe alle Fachzeilen und filtere nach dem gesuchten Fach und den Schülergruppen dieses Schülers.
        foreach (var d in fachZeilen)
        {
            var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var id = d.ContainsKey("Field1") ? d["Field1"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var schuelergruppe = d.ContainsKey("Field42") ? d["Field42"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var klasse = d.ContainsKey("Field5") ? d["Field5"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
            var lehrer = ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "";

            // Fall 1: Die Schülergruppe ist in den Schülergruppen dieses Schülers enthalten.
            if (schuelergruppenDiesesSchuelersinDiesemKlassenfach.Contains(schuelergruppe))
                if (!string.IsNullOrWhiteSpace(lehrer) && lehrer != "?")
                    if (!lehrerListe.Contains(lehrer))
                        lehrerListe.Add(lehrer);

            // Fall 2: Die Schülergruppe ist leer oder nicht vorhanden.
            if (string.IsNullOrWhiteSpace(schuelergruppe))
                if (!string.IsNullOrWhiteSpace(lehrer) && lehrer != "?")
                    if (!lehrerListe.Contains(lehrer))
                        lehrerListe.Add(lehrer);

        }
        return lehrerListe;
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

    private Unterricht GetUnterricht(IEnumerable<object> fachZeilen, List<string> faecherRohFach, IEnumerable<dynamic> faecherDat, IConfiguration configuration, string schuelergruppe = null)
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
        var fachbereinigt = GetBaseSubjectWithCourseType(erstesFach, faecherDat, configuration);

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

        if (!string.IsNullOrEmpty(schuelergruppeFilter))
        {
            kursart = "Kurs";
        }
        if (weitereLehrer.Count() > 0)
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

    private List<string> GetLehrerFachs(string fach, IEnumerable<object> fachZeilen, IEnumerable<dynamic> faecherDat, IConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(fach) || fachZeilen == null)
            return new List<string>();

        var gesuchtesFach = GetBaseSubjectWithCourseType(fach.Trim().Trim('"'), faecherDat, configuration);

        return fachZeilen
            .OfType<IDictionary<string, object>>()
            .Where(d =>
            {
                var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
                var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat, configuration);

                return baseFach.Equals(gesuchtesFach, StringComparison.OrdinalIgnoreCase);
            })
            .Select(d => ExtrahiereLehrer(d)?.Trim().Trim('"') ?? "")
            .Where(l => !string.IsNullOrWhiteSpace(l) && l != "?")
            .Distinct()
            .OrderBy(l => l)
            .Select(l => $"{gesuchtesFach}-{l}")
            .ToList();
    }

    private List<string> GetLehrerDesFachs(string fach, IEnumerable<object> fachZeilen, IEnumerable<dynamic> faecherDat, IConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(fach) || fachZeilen == null)
            return new List<string>();

        var gesuchtesFach = GetBaseSubjectWithCourseType(fach.Trim().Trim('"'), faecherDat, configuration);

        return fachZeilen
            .OfType<IDictionary<string, object>>()
            .Where(d =>
            {
                var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
                var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat, configuration);

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

    private List<IDictionary<string, object>> GetZeilenZuKlassenFaechern(string klasse, string fach, IEnumerable<object> klassenZeilen, IEnumerable<dynamic> faecherDat, IConfiguration configuration)
    {
        if (string.IsNullOrWhiteSpace(klasse) || string.IsNullOrWhiteSpace(fach) || klassenZeilen == null)
            return new List<IDictionary<string, object>>();

        var gesuchteKlasse = klasse.Trim().Trim('"');
        var gesuchtesFach = GetBaseSubjectWithCourseType(fach.Trim().Trim('"'), faecherDat, configuration);

        return klassenZeilen
            .OfType<IDictionary<string, object>>()
            .Where(d =>
            {
                var k = d.ContainsKey("Field5") ? d["Field5"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
                var rawFach = d.ContainsKey("Field7") ? d["Field7"]?.ToString()?.Trim()?.Trim('"') ?? "" : "";
                var baseFach = GetBaseSubjectWithCourseType(rawFach, faecherDat, configuration);

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

    private List<string> GetFaecherBereinigt(List<string> faecherRoh, IEnumerable<dynamic> faecherDat, IConfiguration configuration)
    {
        List<string> faecherBereinigt = new List<string>();

        foreach (var f in faecherRoh)
        {
            if (f == "GPF1")
            {
                string a = "";
            }


            // Es wird versucht das gegebene Untis-Fach (das oft Zusätze oder Kursnummern enthält) auf ein gültiges Basis-Fach (InternKrz) aus einer SchILD-Datenquelle (faecherDat) abzubilden.
            var ff = GetBaseSubjectWithCourseType(f, faecherDat, configuration);

            if (!faecherBereinigt.Contains(ff))
            {
                faecherBereinigt.Add(ff);
            }
        }
        return faecherBereinigt;
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

    public string GetBaseSubjectWithCourseType(string subject, IEnumerable<dynamic> faecherDat, IConfiguration configuration = null)
    {
        // Optional: Konfiguration für Suffix-Abschneiden
        // 1: Suffix auch dann abschneiden, wenn das Fach mit Suffix in faecherDat gefunden wird.
        // 2: Suffix nur abschneiden, wenn das Fach mit Suffix nicht in faecherDat gefunden wird.
        // 3: Suffix abschneiden, aber nicht in Gym (XX G1 oder YY L2)
        var suffixAbschneiden = configuration["SuffixAbschneiden"].ToString().ToLower();

        if (string.IsNullOrWhiteSpace(subject))
            return string.Empty;

        string originalSubject = subject.Trim();
        bool letztesZeichenIstZiffer = char.IsDigit(originalSubject.Last());
        bool kursDesBeruflichenGymnasiums = originalSubject.Contains(" G") || originalSubject.StartsWith(" L");

        // Hilfsfunktion: Prüft, ob ein Fach-Kürzel in faecherDat (Spalte1: InternKrz) existiert
        bool IsMatch(string krz) => faecherDat.Any(f => string.Equals(f.InternKrz?.ToString(), krz, StringComparison.OrdinalIgnoreCase));

        if (originalSubject == "GPF1")
        {
            string a = "";
        }

        // Schritt 1: Exakter Match mit dem ursprünglichen Fach
        if (IsMatch(originalSubject))
        {
            // Das Fach matcht direkt in faecherDat, aber
            // wenn suffixImmerAbschneiden = true ist, wird trotzdem der Suffix-Abschneideprozess durchgeführt.
            if (
                (suffixAbschneiden == "1" && letztesZeichenIstZiffer) ||
                (suffixAbschneiden == "3" && letztesZeichenIstZiffer && !kursDesBeruflichenGymnasiums)
                )
            {
                // Suffix-Abschneideprozess durchführen
                var originalSubjectOhneZiffer = originalSubject.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');

                if (IsMatch(originalSubjectOhneZiffer))
                    return originalSubjectOhneZiffer;
            }
            else
            {
                return originalSubject;
            }
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
}