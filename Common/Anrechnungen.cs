using System.Dynamic;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Common;

public partial class Anrechnungen : List<Anrechnung>
{
    [GeneratedRegex(@"\[[^\]]*\]")]
    private static partial Regex MyRegex();

    public Anrechnungen()
    {
    }

    public Anrechnungen(Lehrers lehrers, IConfiguration configuration)
    {
        using var odbcConnection = new SqlConnection(configuration["ConnectionStringUntis"]);
        var beschreibungs = new Beschreibungs();
        var cvreasons = new CvReasons();

        try
        {
            var queryString = $"""
                               SELECT 
                               CV_REASON_ID, 
                               Name, 
                               Longname
                               FROM CV_Reason
                               WHERE (SCHOOLYEAR_ID={Global.AktSj[0]}{Global.AktSj[1]});
                               """;

            var odbcCommand = new SqlCommand(queryString, odbcConnection);
            odbcConnection.Open();
            var sqlDataReader = odbcCommand.ExecuteReader();

            while (sqlDataReader.Read())
            {
                var cvreason = new CvReason()
                {
                    Id = sqlDataReader.GetInt32(0),
                    Name = Global.SafeGetString(sqlDataReader, 1)
                };

                cvreasons.Add(cvreason);
            }

            sqlDataReader.Close();
        }
        finally
        {
            odbcConnection.Close();
        }

        try
        {
            var queryString = $@"SELECT 
DESCRIPTION_ID, 
Name, 
Longname
FROM Description
WHERE (SCHOOLYEAR_ID={Global.AktSj[0]}{Global.AktSj[1]});";

            var odbcCommand = new SqlCommand(queryString, odbcConnection);
            odbcConnection.Open();
            var sqlDataReader = odbcCommand.ExecuteReader();

            while (sqlDataReader.Read())
            {
                var beschreibung = new Beschreibung()
                {
                    BeschreibungId = sqlDataReader.GetInt32(0),
                    Name = Global.SafeGetString(sqlDataReader, 1),
                    Langname = Global.SafeGetString(sqlDataReader, 2)
                };

                beschreibungs.Add(beschreibung);
            }

            sqlDataReader.Close();
        }
        finally
        {
            odbcConnection.Close();
        }

        try
        {
            var queryString = $@"SELECT 
CountValue.TEACHER_ID,  
DESCRIPTION_ID, 
CountValue.Text,
CountValue.Value,
CountValue.DateFrom,
CountValue.DateTo,
CountValue.CV_REASON_ID

FROM CountValue
WHERE (((CountValue.SCHOOLYEAR_ID)={Global.AktSj[0]}{Global.AktSj[1]}) AND ((CountValue.Deleted)='false') AND ((CountValue.Deleted)='false'))
ORDER BY CountValue.TEACHER_ID;
";

            var sqlCommand = new SqlCommand(queryString, odbcConnection);
            odbcConnection.Open();
            var sqlDataReader = sqlCommand.ExecuteReader();

            while (sqlDataReader.Read())
            {
                var anrechnung = new Anrechnung
                {
                    TeacherIdUntis = sqlDataReader.GetInt32(0),
                    Grund = Convert.ToInt32((from c in cvreasons where c.Id == sqlDataReader.GetInt32(6) select c.Name)
                        .FirstOrDefault()),
                    Wert = Convert.ToDouble(sqlDataReader.GetInt32(3)) / 100000,
                    Lehrer = (from l in lehrers where l.IdUntis == sqlDataReader.GetInt32(0) select l).FirstOrDefault(),
                    // Die Beschr muss auf eine Wiki-Seite matchen. Beschr entspricht einem Thema oder einem Gremium
                    Beschr = (from b in beschreibungs where b.BeschreibungId == sqlDataReader.GetInt32(1) select b.Name)
                        .FirstOrDefault() == null
                            ? ""
                            : (from b in beschreibungs
                                where b.BeschreibungId == sqlDataReader.GetInt32(1)
                                select b.Name).FirstOrDefault(),
                    // Amt und Rolle ergeben sich aus dem Text bei Grund 500 und nur dann, wenn ein KuK zugeordnet wurde. Angaben in Klammern werden ignoriert.
                    Text = Global.SafeGetString(sqlDataReader, 2) == null
                        ? ""
                        : Global.SafeGetString(sqlDataReader, 2) // Vorsitz etc.                            
                };

                anrechnung.Amt = anrechnung.Text != null && anrechnung.Text.Contains("A14") ? "A14" :
                    anrechnung.Text != null && anrechnung.Text.Contains("A15") ? "A15" :
                    anrechnung.Text != null && anrechnung.Text.Contains("A16") ? "A16" : "";

                // Regex für alles in runden, eckigen und geschweiften Klammern inklusive der Klammern selbst

                if (anrechnung.Text != null)
                {
                    var allesAusserKlammern = (Regex.Replace(anrechnung.Text, @"[\(\[\{][^)\]\}]*[\)\]\}]", "")).Trim();
                    if (!string.IsNullOrEmpty(allesAusserKlammern))
                        if (!string.IsNullOrEmpty(allesAusserKlammern))
                            anrechnung.Rolle = (allesAusserKlammern.Replace("A14", "").Replace("A15", "").Replace("A16", ""))
                                .Trim(',').Trim();
                }

                if (anrechnung.Text != null)
                {
                    anrechnung.Hinweis = ZwischenEckigenKlammernStehenHinweise(anrechnung.Text);
                    anrechnung.Kategorien = ZwischenGeschweiftenKlammernStehtDieKategorie(anrechnung.Text);
                }

                anrechnung.Von = sqlDataReader.GetInt32(4) > 0
                    ? DateTime.ParseExact((sqlDataReader.GetInt32(4)).ToString(), "yyyyMMdd",
                        CultureInfo.InvariantCulture)
                    : new DateTime();
                anrechnung.Bis = sqlDataReader.GetInt32(5) > 0
                    ? DateTime.ParseExact((sqlDataReader.GetInt32(5)).ToString(), "yyyyMMdd",
                        CultureInfo.InvariantCulture)
                    : new DateTime();


                if (anrechnung.TeacherIdUntis == 0) continue;
                if (anrechnung.Grund != 0 && anrechnung.Grund <= 210 && anrechnung.Grund != 200 &&
                    anrechnung.Beschr != "Interessen") continue; // Schwerbehinderung etc. nicht einlesen
                if (anrechnung.Lehrer == null) continue;
                if (!string.IsNullOrEmpty(anrechnung.Lehrer.Kürzel))
                {
                    this.Add(anrechnung);
                }
            }

            sqlDataReader.Close();
        }
        finally
        {
            odbcConnection.Close();
            //Global.ZeileSchreiben("Anrechnungen", this.Count().ToString(),                ConsoleColor.Green, ConsoleColor.White);
        }
    }
    private List<string> ZwischenGeschweiftenKlammernStehtDieKategorie(string text)
    {
        List<string> list = new List<string>();
        string pattern = @"\{([^}]*)\}";
        MatchCollection matches = Regex.Matches(text, pattern);

        foreach (Match match in matches)
        {
            list.AddRange(match.Value.Trim().Split(',').Select(_ => match.Value.Trim('{', '}')));
        }

        return list;
    }

    private static string ZwischenEckigenKlammernStehenHinweise(string text)
    {
        var matches = MyRegex().Matches(text);

        foreach (Match match in matches)
        {
            // Entferne die eckigen Klammern selbst
            var content = match.Value.Trim('[', ']');
            return content;
        }

        return "";
    }

    public Datei Anlegen(string dateiname, List<int> nurDieseGrunde, List<int> furDieseGrundeKeinenWert, List<string?> furDieseLehrerKeineWerte)
    {
        var zieldatei = new Datei(dateiname);

        try
        {
            foreach (var anrechnung in this.OrderBy(a => a.Lehrer?.Kürzel))
            {
                if (!nurDieseGrunde.Contains(anrechnung.Grund)) continue;
                var wert = (anrechnung.Wert == 0 ? "" : anrechnung.Wert.ToString(CultureInfo.InvariantCulture));

                if (!furDieseGrundeKeinenWert.Contains(anrechnung.Grund))
                {
                    wert = "";
                }

                if (furDieseLehrerKeineWerte.Contains(anrechnung.Lehrer?.Kürzel))
                {
                    wert = "";
                }

                var kategorien = "";
                if (anrechnung.Kategorien != null)
                    kategorien = anrechnung.Kategorien.Aggregate(kategorien, (current, c) => current + (c + ","));

                anrechnung.Name = (anrechnung.Lehrer?.Titel == "" ? "" : anrechnung.Lehrer?.Titel + " ") +
                                  anrechnung.Lehrer?.Vorname + " " + anrechnung.Lehrer?.Nachname;

                dynamic record = new ExpandoObject();
                record.Name = anrechnung.Name;
                record.Kuerzel = anrechnung.Lehrer?.Kürzel;
                record.Mail = anrechnung.Lehrer?.Mail;
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

            return zieldatei;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Console.ReadKey();
        }

        return [];
    }
}