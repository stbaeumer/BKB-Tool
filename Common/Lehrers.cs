using System.Globalization;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

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


public class Lehrers : List<Lehrer>
{
    //private Dateien dateien;

    public Lehrers()
    {
    }

    public Lehrers(string mitgliederMail)
    {
        foreach (var l in mitgliederMail.Split(','))
        {
            var lehrer = new Lehrer();
            lehrer.Mail = l;
            this.Add(lehrer);
        }
    }
    public Lehrers(IConfiguration configuration, Dateien dateien)
    {
        List<dynamic>? lul = dateien.GetMatchingList(configuration, "lehrkraefte", null, null);

        var liste = new List<dynamic>();

        foreach (var rec in lul)
        {
            var dict = (IDictionary<string, object>)rec;

            var l = new Lehrer();
            l.Kürzel = dict["InternKrz"].ToString();
            l.Nachname = dict["Nachname"].ToString();
            l.Vorname = dict["Vorname"].ToString();
            l.Mail = dict["dienstl. E-Mail"].ToString();
            l.Titel = dict["Titel"].ToString();
            this.Add(l);
        }
    }

    public Lehrers(int periode, Raums raums, IConfiguration configuration)
    {
        var keinGeburtsdatum = new List<string?>();
        var keinGeschlecht = new List<string?>();

        using (var odbcConnection = new SqlConnection(configuration["ConnectionStringUntis"]))
        {
            try
            {
                string queryString = @"SELECT DISTINCT 
Teacher.Teacher_ID, 
Teacher.Name,
Teacher.Longname, 
Teacher.FirstName,
Teacher.Email,
Teacher.ROOM_ID,
Teacher.Title,
Teacher.PlannedWeek,
Teacher.Flags,
Teacher.BirthDate,
Teacher.Text2
 FROM Teacher
WHERE (((SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ") AND  ((TERM_ID)=" + periode +
                                     ") AND ((Teacher.SCHOOL_ID)=177659) AND (((Teacher.Deleted)='false'))) ORDER BY Teacher.Name;";

                SqlCommand odbcCommand = new SqlCommand(queryString, odbcConnection);
                odbcConnection.Open();
                SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                while (sqlDataReader.Read())
                {
                    var lehrer = new Lehrer();

                    lehrer.IdUntis = sqlDataReader.GetInt32(0);
                    lehrer.Kürzel = Global.SafeGetString(sqlDataReader, 1);
                    lehrer.Nachname = Global.SafeGetString(sqlDataReader, 2);

                    if (lehrer.Nachname != "")
                    {
                        try
                        {
                            lehrer.Flags = Global.SafeGetString(sqlDataReader, 10);
                            lehrer.Vorname = Global.SafeGetString(sqlDataReader, 3);
                            lehrer.Mail = Global.SafeGetString(sqlDataReader, 4);
                            lehrer.Raum =
                                (from r in raums where r.IdUntis == sqlDataReader.GetInt32(5) select r.Raumnummer)
                                .FirstOrDefault();
                            lehrer.Titel = Global.SafeGetString(sqlDataReader, 6);
                            lehrer.Text2 = Global.SafeGetString(sqlDataReader, 10);
                            lehrer.Deputat = Convert.ToDouble(sqlDataReader.GetInt32(7)) / 1000;
                            lehrer.Geschlecht = Global.SafeGetString(sqlDataReader, 8).Contains("W") ? "w" : "m";

                            if (lehrer.Geschlecht != "w" && lehrer.Geschlecht != "m")
                            {
                                keinGeschlecht.Add(lehrer.Kürzel);
                            }

                            try
                            {
                                lehrer.Geburtsdatum = DateTime.ParseExact(sqlDataReader.GetInt32(9).ToString(),
                                    "yyyyMMdd", CultureInfo.InvariantCulture);
                            }
                            catch (Exception)
                            {
                                // Bei Nicht-Lehrern ist das Geb.Dat. egal
                                if (lehrer.Deputat > 0)
                                {
                                    if (lehrer.Kürzel != "MOR" && lehrer.Kürzel != "TIS")
                                    {
                                        keinGeburtsdatum.Add(lehrer.Kürzel);
                                    }
                                }
                            }

                            if (lehrer.Geburtsdatum.Year > 1)
                            {
                                //lehrer.AlterAmErstenSchultagDiesesJahres =  lehrer.GetAlterAmErstenSchultagDesSchuljahres();
                                //lehrer.ProzentStelle = lehrer.GetProzentStelle();
                            }

                            this.Add(lehrer);
                        }
                        catch (Exception ex)
                        {
                            Global.ZeileSchreiben(lehrer.Nachname, this.Count().ToString(),
                                ConsoleColor.Red,ConsoleColor.Gray);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw new Exception(ex.ToString());
            }
            finally
            {
                odbcConnection.Close();
                //Global.ZeileSchreiben("Lehrer", this.Count().ToString(),ConsoleColor.Green, ConsoleColor.White);
                foreach (var kuerzel in keinGeburtsdatum)
                {
                    //Global.ZeileSchreiben(1, (kuerzel + " ").PadRight(70, '.'), "kein Geburtsdatum", null);
                }

                foreach (var kuerzel in keinGeschlecht)
                {
                    //Global.ZeileSchreiben(1, (kuerzel + " ").PadRight(70, '.'), "kein Geschlecht", null);
                }
            }
        }
    }

    public void GetTeamsUrl(IEnumerable<dynamic> verschiedeneLehrerKuerzel, string klassen)
    {
        var url = "https://teams.microsoft.com/l/chat/0/0?users=";

        if (Global.User.ToUpper() != "HS")
        {
            url += "Kerstin.hues@berufskolleg-borken.de,";
        }

        if (Global.User.ToUpper() != "LS")
        {
            url += "klaus.lienenklaus@berufskolleg-borken.de,";
        }

        if (Global.User.ToUpper() != "BM")
        {
            url += "stefan.baeumer@berufskolleg-borken.de,";
        }

        if (!Global.User.ToUpper().StartsWith("MOR"))
        {
            url += "ursula.moritz@berufskolleg-borken.de,";
        }

        int anzahlTeamsChat = 0;

        // Der Teams-Chat ist auf Zeichen begrenzt.
        List<string> überzählige = new List<string>();

        foreach (var kuerzel in verschiedeneLehrerKuerzel)
        {
            //var mail = (from l in this where l.Kürzel == kuerzel select l.Mail).FirstOrDefault();

            if (!string.IsNullOrEmpty(kuerzel) && !url.Contains(kuerzel))
            {
                // Es können maximal 7 in den Chat URL übergeben werden.
                anzahlTeamsChat++;

                if (anzahlTeamsChat < 7)
                {
                    url += kuerzel + ",";
                }
                else
                {
                    if (!überzählige.Contains(kuerzel))
                    {
                        überzählige.Add(kuerzel);
                    }
                }
            }
        }

        if (überzählige.Count > 0)
        {
            var table = new Table();
            table.AddColumn("[bold red]A C H T U N G:  [/] \nBrowser-Links dürfen nicht beliebig lang sein. Deswegen werden nicht alle Adressaten an den Chat übergeben. Also bitte folgende LuL dem Teams-Chat zusätzlich manuell hinzufügen:");

            foreach (var item in überzählige)
            {
                table.AddRow(item);
            }
            AnsiConsole.Write(table);
        }
        

            Global.OpenWebseite(url.TrimEnd());

            AnsiConsole.WriteLine("Der Browser sollte jetzt folgenden Link öffnen:\n" +  url.TrimEnd());

        
    }

    public void OffeneKlassenbuchEinträgeMahnen(Dateien dateien, IConfiguration configuration)
    {
        var matchingDatei = dateien.FirstOrDefault(x => x.Name.ToLower().StartsWith("openperiod"));
        var dateiName = matchingDatei != null ? matchingDatei.AbsoluterPfad : throw new InvalidOperationException("No matching file found for 'openperiod'.");

        List<string> lehrer = new List<string>();

        using (var pdfDocument = PdfDocument.Open(dateiName))
        {
            foreach (Page page in pdfDocument.GetPages())
            {
                foreach (var word in page.GetWords())
                {
                    // Prüfe, ob die linke x-Koordinate (BoundingBox.Left) ungefähr 100 ist
                    if (Math.Abs(word.BoundingBox.Left - 100) < 0.1)
                    {
                        if (word.Text != "Lehrkraft"){
                            lehrer.Add(word.Text);
                        }                        
                    }
                }                
            }
        }

        // Gib die 10 häufigsten Nennungen aus der Liste "lehrer" aus
        var topLehrer = lehrer.GroupBy(x => x)
                      .OrderByDescending(g => g.Count())
                      .Take(10)
                      .Select(g => new { Name = g.Key, Count = g.Count() });

        Global.ZeileSchreiben("TOP10 Offene Klassenbuch-Einträge", "Häufigkeit", ConsoleColor.Black, ConsoleColor.Blue);
                
        foreach (var item in topLehrer)
        {
            var nachricht = item.Count.ToString() + " Einträge";

            if(item.Count < 10)
            {
                nachricht += " (keine Nachricht, da <10 Einträge)";
            }
            else if(item.Count > 20)
            {
                nachricht += " (SL in CC, da >20 Einträge)";
            }            

            Global.ZeileSchreiben($"{item.Name}", nachricht, ConsoleColor.Blue, ConsoleColor.Black);        
        }   
        
        Console.WriteLine("  Jetzt die TOP10 per Mail anschreiben? [J/n]");
        var x = Console.ReadKey().Key;
        if (x == ConsoleKey.J || x == ConsoleKey.Enter)
        {
            int i = 1;
            foreach (var item in topLehrer.Where(x => x.Count >= 10))
            {
                var le = (from l in this where l.Kürzel == item.Name select l).FirstOrDefault();
                
                if(le != null)
                {
                    var body = "Guten Morgen " + le.Titel+ le.Vorname + " " + le.Nachname + ",\n\n";
                    body += "es liegen sehr" + (item.Count >= 20 ? ", sehr":"") + " viele offene Klassenbuch-Einträge (" + item.Count + ") vor, die Ihrer Verantwortung zugeordnet sind. \n\n";
                    body += "Bitte kümmern Sie sich zeitnah um die Bearbeitung dieser Einträge.\n\n";
                    body += "Vielen Dank für Ihre Unterstützung.\n\n";
                    body += "Mit freundlichen Grüßen\n\n";
                    body += "Ihr Webuntis-Team";
                
                    var mail = new Mail();                        
                    mail.Senden(  $" Offenen Klassenbuch-Einträge (" + le.Kürzel + ")", configuration, body, null, le.Mail, (item.Count >= 20 ? "stefan.baeumer@berufskolleg-borken.de" : ""), "");
                }
                
                i++;
            }
        }else
        {
            Console.WriteLine("  Sie haben sich gegen den Mailversand entschieden.");
        }
    }
}