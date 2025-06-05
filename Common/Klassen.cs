using System.Text;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

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

public class Klassen : List<Klasse>
{
    public List<Lehrer> Klassenleitungen { get; private set; }

    public Klassen(int periode, Lehrers lehrers, Raums raums, IConfiguration configuration)
    {
        using (SqlConnection odbcConnection = new SqlConnection(configuration["ConnectionStringUntis"]))
        {
            try
            {
                string queryString = @"SELECT DISTINCT 
Class.Class_ID, 
Class.Name,
Class.TeacherIds,
Class.Longname,
Teacher.Name, 
Class.ClassLevel,
Class.PERIODS_TABLE_ID,
Class.ROOM_ID,
Class.Text
FROM Class LEFT JOIN Teacher ON Class.TEACHER_ID = Teacher.TEACHER_ID WHERE (((Class.SCHOOLYEAR_ID)=" +
                                     Global.AktSj[0] + Global.AktSj[1] + ") AND (((Class.TERM_ID)=" + periode +
                                     ")) AND ((Teacher.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] +
                                     ") AND ((Teacher.TERM_ID)=" + periode + ")) OR (((Class.SCHOOLYEAR_ID)=" +
                                     Global.AktSj[0] +
                                     Global.AktSj[1] + ") AND ((Class.TERM_ID)=" + periode +
                                     ") AND ((Class.SCHOOL_ID)=177659) AND ((Teacher.SCHOOLYEAR_ID) Is Null) AND ((Teacher.TERM_ID) Is Null)) ORDER BY Class.Name ASC;";

                var odbcCommand = new SqlCommand(queryString, odbcConnection);
                odbcConnection.Open();
                var sqlDataReader = odbcCommand.ExecuteReader();

                while (sqlDataReader.Read())
                {
                    var klassenleitungen = (Global.SafeGetString(sqlDataReader, 2)).Split(',')
                        .Select(klassenleitungIdUntis => (from l in lehrers
                            where l.IdUntis.ToString() == klassenleitungIdUntis
                            where l.Mail != null
                            where l.Mail != "" // Wer keine Mail hat, kann nicht Klassenleitung sein.
                            select l).FirstOrDefault())
                        .OfType<Lehrer>()
                        .ToList();

                    var istVollzeit = IstVollzeitKlasse(Global.SafeGetString(sqlDataReader, 1));

                    var klasse = new Klasse();

                    klasse.IdUntis = sqlDataReader.GetInt32(0);
                    klasse.Name = Global.SafeGetString(sqlDataReader, 1);
                    klasse.BildungsgangLangname = Global.SafeGetString(sqlDataReader, 3);
                    klasse.Klassenleitungen = klassenleitungen;
                    klasse.IstVollzeit = istVollzeit;
                    klasse.Stufe = Global.SafeGetString(sqlDataReader, 5); // BS-35J-01
                    klasse.WikiLink = Global.SafeGetString(sqlDataReader, 8);
                    klasse.Raum = (from r in raums where r.IdUntis == sqlDataReader.GetInt32(7) select r.Raumnummer)
                        .FirstOrDefault()!;

                    if (klasse.BildungsgangLangname.Contains("("))
                    {
                        klasse.BildungsgangGekürzt = klasse.BildungsgangLangname
                            .Substring(0, klasse.BildungsgangLangname.IndexOf('(')).Trim();
                    }
                    else
                    {
                        klasse.BildungsgangGekürzt = klasse.BildungsgangLangname.Trim();
                    }

                    this.Add(klasse);
                }
                sqlDataReader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
            finally
            {
                odbcConnection.Close();
                //Global.ZeileSchreiben("Klassen", this.Count().ToString(),                    ConsoleColor.Green,ConsoleColor.White);
            }
        }
    }

    public Klassen()
    {
     }

    private static bool IstVollzeitKlasse(string? klassenname)
    {
        var vollzeitBeginn = new List<string>() { "BS", "BW", "BT", "FM", "FS", "G", "HB" };

        foreach (var item in vollzeitBeginn)
        {
            if (klassenname.StartsWith(item))
            {
                return true;
            }
        }

        return false;
    }

    internal void Klassenpflegschaft(IConfiguration configuration, string dateiname, Students schuelers, Lehrers lehrers, Raums raums,
        Common.Anrechnungen untisanrechnungs)
    {
        var vergebeneRäume = new List<string?>();
        var mehrfachVergebeneRäume = new List<string?>();
        Global.OrdnerAnlegen(dateiname);

        var zeilen = new List<string>();

        File.WriteAllText(Path.GetTempPath() + dateiname + ".txt",
            "====== Klassenpflegschaft ======" + Environment.NewLine, Encoding.UTF8);

        File.WriteAllText(Path.GetTempPath() + dateiname + "-datenquelle.csv",
            "\"Klasse\",\"Klassenleitung\",\"Bildungsgang\",\"Raum\",\"Anzahl\"" + Environment.NewLine,
            Encoding.UTF8);

        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "" + Environment.NewLine, Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "" + Environment.NewLine, Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "" + Environment.NewLine, Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt",
            "==== Wir begrüßen alle Eltern sehr herzlich zur Klassenpflegschaft am: ====" + Environment.NewLine,
            Encoding.UTF8);

        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "---- struct global ----" + Environment.NewLine,
            Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "schema: termine_kollegium" + Environment.NewLine,
            Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "cols:Datum" + Environment.NewLine, Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "head:Datum/Uhrzeit" + Environment.NewLine,
            Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt",
            "filterand: Seite~*lassenpflegschaft*" + Environment.NewLine, Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "dynfilters: 0" + Environment.NewLine,
            Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "csv: 0" + Environment.NewLine, Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "----" + Environment.NewLine, Encoding.UTF8);

        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt",
            "**Haben Sie Fragen? Dann melden Sie sich gerne!**" + Environment.NewLine, Encoding.UTF8);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "" + Environment.NewLine, Encoding.UTF8);

        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "==== Raumplan ====" + Environment.NewLine,
            Encoding.UTF8);

        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", "<searchtable>" + Environment.NewLine);
        File.AppendAllText(Path.GetTempPath() + dateiname + ".txt",
            "^  Klasse  ^  Klassenleitung  ^  Bildungsgang  ^  Raum  ^" + Environment.NewLine, Encoding.UTF8);

        List<string?> berufe = new List<string?>();

        foreach (var klasse in (from t in this.OrderBy(x => x.Stufe).ThenBy(x => x.Name)
                     where t.Stufe != null
                     where t.Stufe != ""
                     select t).ToList())
        {
            // Berufsschule wird nach Berufen geclustert

            if (!berufe.Contains(klasse.BildungsgangLangname) && klasse.Stufe.StartsWith("BS"))
            {
                berufe.Add(klasse.BildungsgangLangname);

                var
                    anzahlSuS =
                        10; //todo (from s in schuelers where SchneideVorErsterZahlAb(s.Klasse.NameUntis) == SchneideVorErsterZahlAb(klasse.NameUntis) select s).Count();

                try
                {
                    var lehrer = (from u in untisanrechnungs
                        where u.Text.Contains("ildungsgangleitung")
                        where BgAusschneiden(u.Beschr.ToLower()) ==
                              SchneideVorErsterZahlAb(klasse.Name.ToLower())
                        select u.Lehrer).FirstOrDefault();

                    string leitung = "";

                    if (lehrer != null)
                    {
                        leitung = " [[" + lehrer.Mail + "|" + (lehrer.Geschlecht == "m" ? "Herr" : "Frau") + " " +
                                  (lehrer.Titel == "" ? "" : lehrer.Titel + " ") + lehrer.Nachname + "]]";
                    }

                    if (klasse.Raum != null && klasse.Raum != "")
                    {
                        if (!vergebeneRäume.Contains(klasse.Raum))
                        {
                            vergebeneRäume.Add(klasse.Raum);
                        }
                        else
                        {
                            mehrfachVergebeneRäume.Add(klasse.Raum);
                        }
                    }

                    File.AppendAllText(Path.GetTempPath() + dateiname + "-datenquelle.csv",
                        "\"" + SchneideVorErsterZahlAb(klasse.Name) + "*" + "\",\"" + leitung + "\",\"" +
                        klasse.BildungsgangLangname + "\",\"" + klasse.Raum + "\",\"" + anzahlSuS + "\"" +
                        Environment.NewLine, Encoding.UTF8);

                    zeilen.Add("|" + SchneideVorErsterZahlAb(klasse.Name) + "*" + "  |" + leitung + "  |" +
                               klasse.BildungsgangLangname + "  |" + klasse.Raum + "  |");
                }
                catch (Exception ex)
                {
                    throw;
                }
            }

            if (!klasse.Stufe.StartsWith("BS"))
            {
                try
                {
                    var leitung = "[[" + klasse.Klassenleitungen[0].Mail + "|" +
                                  (klasse.Klassenleitungen[0].Geschlecht == "m" ? "Herr" : "Frau") + " " +
                                  (klasse.Klassenleitungen[0].Titel == ""
                                      ? ""
                                      : klasse.Klassenleitungen[0].Titel + " ") + " " +
                                  klasse.Klassenleitungen[0].Nachname + "]]";

                    var anzahlSuS =
                        0; //todo (from s in schuelers where s.Klasse.NameUntis == klasse.NameUntis select s).Count();

                    if (klasse.Raum != null && klasse.Raum != "")
                    {
                        if (!vergebeneRäume.Contains(klasse.Raum))
                        {
                            vergebeneRäume.Add(klasse.Raum);
                        }
                        else
                        {
                            mehrfachVergebeneRäume.Add(klasse.Raum);
                        }
                    }

                    File.AppendAllText(Path.GetTempPath() + dateiname + "-datenquelle.csv",
                        "\"" + klasse.Name + "\",\"" + leitung + "\",\"" + klasse.BildungsgangLangname +
                        "\",\"" + klasse.Raum + "\",\"" + anzahlSuS + "\"" + Environment.NewLine, Encoding.UTF8);
                    zeilen.Add("|" + klasse.Name + "  |" + leitung + "  |" + klasse.BildungsgangLangname +
                               "  |" + klasse.Raum + "  |");
                }
                catch (Exception ex)
                {
                    throw;
                }
            }
        }

        zeilen.Add("</searchtable>" + Environment.NewLine);
        //todo Global.WriteLine("Klassenpflegschaften", zeilen.Count());
        //todo Global.WriteLine(" Mehrfach vergebene Räume:" + String.Join(",", mehrfachVergebeneRäume), Console.WindowWidth);

        string freieR = "";

        foreach (var raum in raums.OrderBy(x => x.Raumnummer))
        {
            if (!vergebeneRäume.Contains(raum.Raumnummer))
            {
                freieR += raum.Raumnummer + ",";
            }
        }

        freieR = Global.InsertLineBreaks(freieR, 110);

        Console.WriteLine("  Freie Räume: " + freieR.TrimEnd(','));

        foreach (var zeile in zeilen)
        {
            File.AppendAllText(Path.GetTempPath() + dateiname + ".txt", zeile + Environment.NewLine, Encoding.UTF8);
        }

        Global.Dateischreiben(configuration, dateiname + ".txt");
        Global.Dateischreiben(configuration, dateiname + "-datenquelle.csv");
    }

    private string BgAusschneiden(string input)
    {
        // Den Index des letzten und vorletzten Doppelpunkts finden
        int lastColonIndex = input.LastIndexOf(':');
        int secondLastColonIndex = input.LastIndexOf(':', lastColonIndex - 1);

        // Den Teilstring zwischen den letzten beiden Doppelpunkten extrahieren
        if (lastColonIndex != -1 && secondLastColonIndex != -1)
        {
            return input.Substring(secondLastColonIndex + 1, lastColonIndex - secondLastColonIndex - 1);
        }

        // Falls nicht genügend Doppelpunkte vorhanden sind, wird ein leerer String zurückgegeben
        return string.Empty;
    }

    private string? SchneideVorErsterZahlAb(string? input)
    {
        // Schleife über den String, um die Position der ersten Zahl zu finden
        for (int i = 0; i < input.Length; i++)
        {
            if (char.IsDigit(input[i]))
            {
                // Den Teil des Strings bis zur ersten Zahl zurückgeben
                return input.Substring(0, i);
            }
        }

        // Falls keine Zahl gefunden wurde, wird der ursprüngliche String zurückgegeben
        return input;
    }

    

    public void Relationsgruppen(Relationsgruppen relationsgruppen)
    {
        foreach (var klasse in this)
        {
            foreach (var r in relationsgruppen)
            {
                if (r.Gliederungen.Contains(klasse.Gliederung))
                {
                    if ((r.Fachklassenschlüssel.Count == 0 ||
                         r.Fachklassenschlüssel.Contains(klasse.Fachklassenschlüssel)))
                    {
                        if (r.Jahrgänge.Contains(klasse.Jahrgang) && klasse.Name != "Z" &&
                            klasse.Name != "Abgang")
                        {
                            klasse.Relationsgruppe = r.BeschreibungSchulministerium;
                        }
                    }
                }
            }
        }
    }

    public void KlasseOhneRelationsgruppe(Students students)
    {
        foreach (var klasse in (from k in this where k.Relationsgruppe == null select k).ToList())
        {
            Console.Write("Die Klasse " + klasse.Name + " kann keiner Relationsgruppe zugeordnet werden.");
            if ((from s in students where s.Klasse == klasse.Name select s).Any())
            {
                Console.Write("Die Klasse hat aber Schüler. Das muss unbedingt gefixt werden!");
                Console.ReadKey();
            }

            Console.WriteLine();
        }
    }
}