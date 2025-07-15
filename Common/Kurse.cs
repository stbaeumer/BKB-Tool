using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

public class Kurse : List<Kurs>
{    
    public Kurse(IConfiguration configuration, Menüeintrag m, Global.Art art, Klassen alleKlassen)
    {
        // Aus der Kurse.dat wird die Kursart ermittelt. Wenn die Kursart in SchILD einmal gesetzt ist, wird sie nicht mehr geändert.
        List<dynamic> kurseDat = m.Quelldateien.GetMatchingList(configuration, "kurse.", m.IStudents, m.Klassen);        
        if (kurseDat == null)
        {
            AnsiConsole.MarkupLine($"[grey]Keine Kurse.dat gefunden. Bitte exportieren Sie die Datei erneut.[/]");
            return;
        }

        List<dynamic> gpu002 = m.Quelldateien.GetMatchingList(configuration, "gpu002", m.IStudents, m.Klassen);
        if (gpu002 == null || gpu002.Count == 0)
        {
            AnsiConsole.MarkupLine($"[grey]Keine GPU002-Daten gefunden. Bitte exportieren Sie die Datei erneut.[/]");
            return;
        }

        List<dynamic> klassen = m.Quelldateien.GetMatchingList(configuration, "klassen", m.IStudents, m.Klassen);
        if (klassen == null || klassen.Count == 0)
        {
            AnsiConsole.MarkupLine($"[grey]Keine Klassen.dat gefunden. Bitte exportieren Sie die Datei erneut.[/]");
            return;
        }

        

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Kurse aus GPU002.TXT: ...", ctx =>
        {
            // Ordne die GPU002 aufsteigend nach Field6. Dadurch wird erreicht, dass die erste Lehrkrft im Alphabet Kursleiter ist.
            gpu002 = gpu002.OrderBy(record =>
            {
                var dict = (IDictionary<string, object>)record;
                return dict.ContainsKey("Field6") ? dict["Field6"]?.ToString() : string.Empty;
            }).ToList();

            foreach (var record in gpu002)
            {
                var Klassen = new Klassen();
                var dict = (IDictionary<string, object>)record;

                if(nichtInteressierdendeUnterrichtsgruppe(configuration, dict))
                    continue; // Diese Gruppe ist nicht interessant, also überspringen.                

                if(art == Global.Art.Statistik && !dict.ContainsKey("Field12") && außerhalbDesStatistikdatums(configuration, dict))
                    continue; // Statistikdatum liegt außerhalb des Zeitraums.

                //if(art == Global.Art.Zeugnis && !dict.ContainsKey("Field12") && außerhalbDesAktuellenAbschnitts(configuration, dict))
                //    continue; // Statistikdatum liegt außerhalb des Zeitraums.

                var unterrichtsId = dict["Field1"]?.ToString();

                if (unterrichtsId == "985")
                {
                    string test = "Test"; // Debugging purpose
                }

                if (zeileIstKursOderGehörtZuKurs(gpu002, dict))
                {
                    var jahrgang = GetJahrgang(alleKlassen, dict);
                    var fach = Bereinigen(dict["Field7"]?.ToString());
                    var schuelergruppe = dict["Field42"]?.ToString();
                    var klasse = dict["Field5"]?.ToString();
                    var lehrer = dict["Field6"]?.ToString();                    
                    int wochentundenLehrkraft = int.TryParse(dict["Field2"]?.ToString(), out int ws) ? ws : 0;
                    int wochenstundenKurs = 0;
                    var untisId = dict["Field1"]?.ToString();

                    if (wochentundenLehrkraft == 0) continue;

                    var kurs = this.FirstOrDefault(k =>
                        //k.UnterrichtsIds.Contains(int.Parse(unterrichtsId)) &&
                        Bereinigen(k.Fach) == Bereinigen(fach) &&
                        k.Schülergruppe == schuelergruppe &&
                        (k.Klassen.Contains(klasse) ||
                        (k.KursBez.StartsWith(lehrer) && k.KursBez.Contains(untisId) // Kopplung von zwei Klassen in Untis
                        )));

                    // Wenn die Kombination aus vorhandener Schülergruppe, Klasse & Fach noch nicht vorhanden -> neuer Kurs
                    if (kurs == null)
                    {
                        // Die Kursart wird aus der Kurse.dat ermittelt, wenn sie dort einmal gesetzt ist. Sie steckt in "Kursart" des Dictionaries.
                        var dictKurs = kurseDat
                            .Select(k => k as IDictionary<string, object>)
                            .FirstOrDefault(k => k != null && k.ContainsKey("Fach") && k["Fach"]?.ToString() == fach && k.ContainsKey("Kursart") && k["Kursart"] != null);
                        string kursart = dictKurs?["Kursart"]?.ToString() ?? string.Empty;

                        this.Add(new Kurs
                        {
                            // Klassen nicht leer -> Alle bekommen den Kurs zugewiesen
                            //                       KursBez darf leer bleiben
                            // Klassen       leer -> keine automatische Zuweisung, muss manuell erfolgen

                            KursBez = $"{lehrer}-{untisId}",                            
                            Fach = Bereinigen(fach),
                            Klassen = new List<string>() { klasse },
                            Jahrgaenge = jahrgang != null ? new List<string> { jahrgang } : new List<string>(),
                            Kursart = kursart,
                            Wochenstunden = wochentundenLehrkraft,
                            Kursleiter = lehrer,
                            KursleiterWochenstunden = wochentundenLehrkraft,
                            Lehrkraefte = new List<string>(),
                            LehrkraefteWochenstunden = new List<int>(),
                            Schülergruppe = schuelergruppe,
                            UnterrichtsIds = new List<int> { int.Parse(unterrichtsId) }
                        });
                        continue;
                    }
                    else
                    {
                        // Wenn der Kurs bereits existiert und in einer zweiten Zeile eine weitere Klasse hinzugefügt wird
                        if (
                                kurs.KursBez.StartsWith(lehrer) &&
                                kurs.KursBez.Contains(unterrichtsId) &&
                                kurs.Schülergruppe == schuelergruppe &&
                                Bereinigen(kurs.Fach) == Bereinigen(fach)
                            )
                        {
                            if (!kurs.Klassen.Contains(klasse))
                                kurs.Klassen.Add(klasse);

                            if (kurs.Jahrgaenge.Contains(jahrgang) == false && jahrgang != null && !string.IsNullOrEmpty(jahrgang))
                                kurs.Jahrgaenge.Add(jahrgang);


                            // Die Anzahl der Wochenstunden wird nicht erhöht.
                            continue;
                        }




                        // Wenn UnterrichtsId nicht in der Liste der UnterrichtsIds des Kurses enthalten ist, wird sie hinzugefügt
                        if (!kurs.UnterrichtsIds.Contains(int.Parse(unterrichtsId)))
                        {
                            kurs.UnterrichtsIds.Add(int.Parse(unterrichtsId));
                            // Eine weitere UterrichtsId verändert die Kursbezeichnung                            
                            kurs.KursBez = $"{kurs.Kursleiter}-{string.Join('-', kurs.UnterrichtsIds)}";
                            // Die Wochenstunden des Kurses erhöhen sich nur, wenn die UnterrichtsId neu ist
                            kurs.Wochenstunden += wochentundenLehrkraft;

                            if (kurs.Jahrgaenge.Contains(jahrgang) == false && jahrgang != null && !string.IsNullOrEmpty(jahrgang))
                                kurs.Jahrgaenge.Add(jahrgang);
                        }

                        // Wenn lehrer == Kursleiter, dann wird die Wochenstundenzahl des Kursleiters aktualisiert
                        if (lehrer == kurs.Kursleiter)
                        {
                            kurs.KursleiterWochenstunden += wochentundenLehrkraft;
                        }
                        else
                        {
                            // Wenn lehrer != Kursleiter, dann wird der Lehrer als Zusatzlehrer hinzugefügt
                            if (!kurs.Lehrkraefte.Contains(lehrer))
                            {
                                kurs.Lehrkraefte.Add(lehrer);
                                kurs.LehrkraefteWochenstunden.Add(wochentundenLehrkraft);
                            }
                            else
                            {
                                // Wenn der Lehrer bereits als Zusatzlehrer existiert, wird die Wochenstundenzahl aktualisiert
                                int index = kurs.Lehrkraefte.IndexOf(lehrer);
                                if (index >= 0)
                                {
                                    kurs.LehrkraefteWochenstunden[index] += wochentundenLehrkraft;
                                }
                            }
                        }
                    }
                }
            }
            Global.ZeileSchreiben("Kurse aus GPU002.TXT:", this.Count().ToString());
        });
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
            var interessierendeGruppen = configuration["InteressierendeUnterrichtsgruppen"]?.Split(',').Select(s => s.Trim()).ToList();
            if (interessierendeGruppen != null && !interessierendeGruppen.Contains(dict["Field12"].ToString()))
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
        // Ein Kurs ist definiert, wenn Field42 nicht leer ist.
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
        // Ein Zähler am Ende des Fachs wird nicht berücksichtigt. Bsp: M1 wird zu M., M G1 bleibt M G1, weil M das Fach ist und G1 der Grundkurs.

        var count1 = gpu002.Count(record =>
            {
                var recDict = (IDictionary<string, object>)record;
                return Bereinigen(recDict["Field7"].ToString()) == Bereinigen(dict["Field7"].ToString()) &&
                recDict["Field5"].ToString() == dict["Field5"].ToString() &&
                recDict["Field42"].ToString() == dict["Field42"].ToString();
            });

        if (count1 > 1)
        {
            return true; // Es ist ein Kurs, wenn dasselbe Fach und dieselbe Klasse mehrfach vorkommen.
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

public class Kurs
{
    public string Kursleiter { get; set; }
    public string? Schülergruppe { get; set; }
    public List<int> UnterrichtsIds { get; set; }
    public int KursleiterWochenstunden { get; set; }
    public string? Fach { get; set; }
    public List<string> Klassen { get; internal set; }
    public List<string> Lehrkraefte { get; internal set; }
    public List<int> LehrkraefteWochenstunden { get; internal set; }
    public string KursBez { get; internal set; }
    public string Kursart { get; internal set; }
    public int Wochenstunden { get; internal set; }
    public List<string> Jahrgaenge { get; internal set; }
}