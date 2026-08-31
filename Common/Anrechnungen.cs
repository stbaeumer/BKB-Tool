using System.Dynamic;
using System.Globalization;
using System.Text;
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
        configuration = Global.Konfig("ConnectionStringUntis", Global.Modus.Update, configuration);

        using var odbcConnection = new SqlConnection(configuration["ConnectionStringUntis"]);
        var beschreibungs = new Beschreibungs();
        var cvreasons = new CvReasons();

        try
        {
            var queryString = $"""
                               SELECT
                                   cr.CV_REASON_ID,
                                   cr.Name,
                                   cr.Longname
                               FROM dbo.CV_Reason cr
                               WHERE (cr.SCHOOLYEAR_ID={Global.AktSj[0]}{Global.AktSj[1]});
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

        int id = 0;

        try
        {
            var queryString = """
WITH RankedCountValues AS (
    SELECT 
        t.TEACHER_ID,
        t.Longname AS TeacherLongname,
        t.FirstName AS TeacherFirstName,
        t.Email AS TeacherEmail,
        t.Name AS TeacherName,
        t.Title,
        cv.SCHOOLYEAR_ID AS Schuljahr,
        cv.Text AS CountValueText,
        cv.Value,
        cv.DateFrom,
        cv.DateTo,
        cvr.Longname AS CVReasonLongname,
        cvr.Name AS CVReasonName,
        d.Name AS DescriptionName,
        ROW_NUMBER() OVER (
            PARTITION BY 
                t.TEACHER_ID, 
                cv.Text, 
                d.Name
            ORDER BY 
                cv.DateFrom ASC, 
                cv.COUNT_VALUE_ID ASC
        ) AS rn
    FROM 
        [dbo].[Teacher] AS t
        LEFT JOIN [dbo].[CountValue] AS cv 
            ON cv.TEACHER_ID = t.TEACHER_ID
            AND cv.SCHOOLYEAR_ID = t.SCHOOLYEAR_ID
            AND cv.VERSION_ID = t.VERSION_ID
            AND cv.SCHOOLYEAR_ID = 20252026
            AND cv.VERSION_ID = 1
            AND cv.Deleted = 0
            AND NOT (cv.Text IS NULL AND cv.DESCRIPTION_ID IS NULL)
        LEFT JOIN [dbo].[CV_Reason] AS cvr 
            ON cv.CV_REASON_ID = cvr.CV_REASON_ID
            AND cv.SCHOOLYEAR_ID = cvr.SCHOOLYEAR_ID
            AND cv.VERSION_ID = cvr.VERSION_ID
        LEFT JOIN [dbo].[Description] AS d 
            ON cv.DESCRIPTION_ID = d.DESCRIPTION_ID
            AND cv.SCHOOLYEAR_ID = d.SCHOOLYEAR_ID
            AND cv.VERSION_ID = d.VERSION_ID
    WHERE 
        t.SCHOOLYEAR_ID = 20252026
        AND t.VERSION_ID = 1
        AND t.Email IS NOT NULL
        AND t.Email <> ''
)
SELECT 
    Schuljahr,
    TEACHER_ID,
    CountValueText,
    Value,
    DateFrom,
    DateTo,
    TeacherLongname,
    TeacherFirstName,
    TeacherEmail,
    TeacherName,
    CVReasonLongname,
    CVReasonName,
    DescriptionName,
    Title
FROM RankedCountValues
WHERE rn = 1 OR rn IS NULL
ORDER BY TeacherName;
""";

            var sqlCommand = new SqlCommand(queryString, odbcConnection);
            odbcConnection.Open();
            var sqlDataReader = sqlCommand.ExecuteReader();

            while (sqlDataReader.Read())
            {
                var anrechnung = new Anrechnung();

                anrechnung.TeacherIdUntis = sqlDataReader.GetInt32(1);
                id = anrechnung.TeacherIdUntis;

                var teacherName = Global.SafeGetString(sqlDataReader, 9);
                if (string.IsNullOrEmpty(teacherName))
                {
                    continue;
                }

                anrechnung.LehrerKuerzel = teacherName;
                anrechnung.Vorname = Global.SafeGetString(sqlDataReader, 7);
                anrechnung.Nachname = Global.SafeGetString(sqlDataReader, 6);
                anrechnung.Mail = Global.SafeGetString(sqlDataReader, 8);
                anrechnung.Titel = Global.SafeGetString(sqlDataReader, 13);

                if (anrechnung.LehrerKuerzel == "HR")
                {
                    var debug = 1;
                }

                anrechnung.Lehrer = (from l in lehrers where l.Kürzel == teacherName select l).FirstOrDefault();

                anrechnung.Beschr = Global.SafeGetString(sqlDataReader, 12);

                if(anrechnung.Beschr.ToLower().StartsWith("kollegium:"))
                 anrechnung.Beschr.Replace("kollegium:", "schulgemeinschaft:");
                

                anrechnung.Text = Global.SafeGetString(sqlDataReader, 2) == null
                    ? ""
                    : Global.SafeGetString(sqlDataReader, 2); // Vorsitz etc.

                var grund = Global.SafeGetString(sqlDataReader, 11);

                // Versuche grund in eine Zahl zu konvertieren, wenn nicht möglich -> überspringe diesen Datensatz
                if (!int.TryParse(grund, out int grundWert))
                {
                    anrechnung.Grund = -1;    // -1 bei denen, die keinen Grund haben
                }
                else
                {
                    anrechnung.Grund = grundWert;
                }

                try
                {
                    anrechnung.Wert = Convert.ToDouble(sqlDataReader.GetInt32(3)) / 100000;
                }
                catch
                {
                    anrechnung.Wert = 0;
                }
                
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

                try
                {
                    anrechnung.Hinweis = anrechnung.Hinweis.Replace("[", "").Replace("]", "").Trim();
                }
                catch
                {
                    anrechnung.Hinweis = "";
                }
                try
                {
                    anrechnung.Von = sqlDataReader.GetInt32(4) > 0
                    ? DateTime.ParseExact((sqlDataReader.GetInt32(4)).ToString(), "yyyyMMdd",
                        CultureInfo.InvariantCulture)
                    : new DateTime();                 
                }
                catch
                {
                    anrechnung.Von = new DateTime();                    
                }
                
                try
                {                    
                    anrechnung.Bis = sqlDataReader.GetInt32(5) > 0
                    ? DateTime.ParseExact((sqlDataReader.GetInt32(5)).ToString(), "yyyyMMdd",
                        CultureInfo.InvariantCulture)
                    : new DateTime();
                }
                catch
                {                    
                    anrechnung.Bis =  new DateTime();
                }
                

                if (anrechnung.TeacherIdUntis == 0) continue;
                //if (anrechnung.Grund != 0 && anrechnung.Grund <= 210 && anrechnung.Grund != 200 &&
                //    anrechnung.Beschr != "Interessen") continue; // Schwerbehinderung etc. nicht einlesen

                this.Add(anrechnung);
            }

            sqlDataReader.Close();
        }
        catch (Exception ex)
        {
            throw new Exception("Fehler beim Einlesen der Anrechnungen: " + ex.Message + " (LehrerIdUntis: " + id + ")");
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
}