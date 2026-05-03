using Common;
using Microsoft.Extensions.Configuration;
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

                if (wochentundenLehrkraft == 0) continue; // Unterricht mit 0 Wochenstunden werden nicht berücksichtigt.

                // Wenn es um Kurse geht und die Zeile ein Kurs ist oder zu einem Kurs gehört
                if (art == Global.Art.KursUnterrichte && zeileIstKursOderGehörtZuKurs(gpu002, dict))
                {
                    // Suche nach einem bestehenden Kurs mit identischer UntisID
                    var kurs = this.FirstOrDefault(k => k.UnterrichtsIds.Contains(Convert.ToInt32(unterrichtsId)));

                    // Wenn kein passender Kurs gefunden wurde, s
                    if (kurs == null)
                    {
                        kurs = this.FirstOrDefault(k =>
                        Bereinigen(k.Fach) == Bereinigen(fach) &&
                        k.Schülergruppe == schuelergruppe &&
                        (k.Klassen.Contains(klasse) ||
                            (k.KursBez.StartsWith(lehrer) && k.KursBez.Contains(unterrichtsId) // Kopplung von zwei Klassen in Untis
                        )));
                    }

                    if (kurs == null)
                        {
                            var kursBez = $"{lehrer}-{unterrichtsId}";
                            Add(new Unterricht(kursBez, zweck, m, unterrichtsId, fach, schuelergruppe, klasse, lehrer, wochentundenLehrkraft, kurseDat, configuration, studentgroupStudents));
                        }
                        else
                        {
                            // Wochenstunden werden bei Bedarf erhöht usw.
                            var kursBez = kurs.Updaten(zweck, m, configuration, fach, lehrer, unterrichtsId, schuelergruppe, klasse, wochentundenLehrkraft, studentgroupStudents);
                        }
                }
                else if (art == Global.Art.NichtKursUnterrichte && !zeileIstKursOderGehörtZuKurs(gpu002, dict))
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
    }

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

    private DateTime jjjjmmddNachDateTime(object v)
    {
        if (v == null || v.ToString().Length != 8)
            return DateTime.MinValue; // Ungültiges Datum

        if (DateTime.TryParseExact(v.ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime dateTime))
        {
            return dateTime;
        }
        return DateTime.MinValue; // Ungültiges Datum
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

    private bool zeileIstKursOderGehörtZuKurs(List<dynamic> gpu002, IDictionary<string, object> dict)
    {
        if (dict["Field1"]?.ToString() == "3000" || dict["Field1"]?.ToString() == "2263")
        {
            string aa = "Test"; // Debugging purpose
        }

        // Ein Kurs ist definiert, wenn Field42 (Schülergruppe) nicht leer ist.
        if (dict.ContainsKey("Field42") && !string.IsNullOrEmpty(dict["Field42"]?.ToString()))
        {   
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