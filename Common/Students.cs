using System.Dynamic;
using System.Globalization;
using System.Reflection;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using ICSharpCode.SharpZipLib.Zip;
using Common;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using Org.BouncyCastle.Bcpg.OpenPgp;

#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise null-Objekts.
#pragma warning disable CS8604 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8600 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8618 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0219 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8625 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8601 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0168 // Möglicher Null-Verweis-Argument

public class Students : List<Student>
{
    private Students InteressierendeStudents { get; set; }
    private string DateiPfad { get; set; }
    private DateTime Erstelldatum { get; }
    public string Klasse { get; }
    public string? Schuelergruppe { get; }
    public string QuellFoto { get; private set; }

    public Students()
    {
    }

    public Students(IConfiguration configuration, Dateien quelldateien)
    {
        var schuelerBasisdaten = quelldateien.GetMatchingList(configuration, "schuelerbasisdaten", new Students(), null);
        if (schuelerBasisdaten == null || schuelerBasisdaten.Count == 0) return;
        var schuelerZusatzdaten = quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", new Students(), null);
        if (schuelerZusatzdaten == null || schuelerZusatzdaten.Count == 0) return;

        if (schuelerBasisdaten.Count != schuelerZusatzdaten.Count)
        {
            Console.WriteLine("Die Anzahl der Schüler in den Basis- und Zusatzdaten stimmt nicht überein.");
            return;
        }

        for (int j = 0; j < schuelerBasisdaten.Count; j++)
        {
            var student = new Student();
            try
            {
                var sb = (IDictionary<string, object>)schuelerBasisdaten[j];
                var sz = (IDictionary<string, object>)schuelerZusatzdaten[j];

                student.Vorname = sb["Vorname"].ToString();
                student.Nachname = sb["Nachname"].ToString();
           
                if(student.Nachname.ToLower().Contains("boldt"))
                {
                    string aa = "";
                }
           
                student.Geburtsdatum = sb["Geburtsdatum"].ToString();
                student.Klasse = sb["Klasse"].ToString();
                student.Status = sb["Status"].ToString();
                student.Geschlecht = sb["Geschlecht"].ToString();
                student.Ort = sb["Ort"].ToString();
                student.Postleitzahl = sb["PLZ"].ToString();
                student.Straße = sb["Straße"].ToString();
                student.MailSchulisch = sz["schulische E-Mail"].ToString();
                student.BeginnDesBildungsganges = sz["BeginnBildungsgang"].ToString();
                student.Fachklasse = sb["Fachklasse"].ToString();
                student.Jahrgang = sb["Jahrgang"].ToString();
                student.Schulgliederung = sb["Schulgliederung"].ToString();
                // Wenn eine externe ID eingetragen ist, wird diese verwendet, ansonsten die schulische E-Mail-Adresse ohne Domain.
                student.Id = string.IsNullOrEmpty(sz["Externe ID-Nr"].ToString()) ? sz["schulische E-Mail"].ToString().Split('@')[0] : sz["Externe ID-Nr"].ToString();
                Add(student);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Verarbeiten von {student.Vorname} {student.Nachname}, Zeile {j + 2}: {ex.Message}");
                continue; // Weiter mit dem nächsten Schüler
            }
        }
    }

    public Students(IConfiguration configuration, string dateiName, string dateiendung, string delimiter = "|")
    {
        DateiPfad = Global.CheckFile(configuration, dateiName, dateiendung);

        if (DateiPfad == null)
        {
            return;
        }

        Erstelldatum = File.GetLastWriteTime(DateiPfad);

        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HeaderValidated = null,
            MissingFieldFound = null,
            HasHeaderRecord = true,
            Delimiter = delimiter
        };

        using (var reader = new StreamReader(DateiPfad))
        using (var csv = new CsvReader(reader, config))
        {
            csv.Context.RegisterClassMap<SchuelersMap>();
            var records = csv.GetRecords<Student>();

            this.AddRange(records);
            foreach (var student in this)
            {
                student.MailSchulisch = student.GenerateMailAusGebdat(configuration);

                if (student.AktuellerAbschnitt == "")
                {
                    student.AktuellerAbschnitt = configuration["Abschnitt"];
                }

                if (student.AktuellesHalbjahr == "")
                {
                    student.AktuellesHalbjahr = Global.AktSj[0];
                }
            }
        }

        Global.ZeileSchreiben((DateiPfad + " ").PadRight(90, '.') + " " + Erstelldatum, this.Count().ToString(),
            ConsoleColor.Yellow, ConsoleColor.Gray);
    }

    

    private void ErgänzeFehlendeEigenschaften(Student vS, Student nS)
    {
        foreach (PropertyInfo property in typeof(Student).GetProperties())
        {
            if (property.PropertyType == typeof(string))
            {
                string vSvalue = (string)property.GetValue(vS);

                if (vSvalue.ToLower().Contains("externe"))
                {
                    string aa = "";
                }

                string? nSvalue = (string)property.GetValue(nS)!;

                if (string.IsNullOrEmpty(vSvalue))
                {
                    if (!string.IsNullOrEmpty(nSvalue))
                    {
                        property.SetValue(vS, nSvalue);
                    }
                    else
                    {
                        property.SetValue(vS, "");
                    }
                }
            }
        }
    }

    public void SchulpflichtüberwachungTxt(
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string dateiname,
        int schonfrist,
        int warnungAbAnzahl,
        int verjaehrungUnbescholtene,
        int nachSovielenTagenVerjährenFehlzeitenBeiMaßnahme,
        Klassen klasses,
        Lehrers lehrers,
        Dateien dateien)
    {

        var schuelerZusatzdaten = dateien.GetMatchingList(configuration, "schuelerzusatzdaten", null, null);
        if (schuelerZusatzdaten == null || !schuelerZusatzdaten.Any()) throw new Exception("Keine SchuelerZusatzdaten.dat");

        var gpu003 = dateien.GetMatchingList(configuration, "gpu003", null, null);
        if (gpu003 == null || gpu003.Count == 0) return;

        var zieldatei = new Datei(dateiname, funktionen, configuration);
        zieldatei.Lehrers = new Lehrers();        

        zieldatei.Add("====== Schulpflichtüberwachung ======");

        var zeilen = new List<string>();
        zieldatei.Add("");

        zieldatei.Add(@"**Hallo Klassenleitung,**");
        zieldatei.Add(@"");
        zieldatei.Add(
            @"Du wurdest hierher verlinkt, weil bei der automatisierten, wöchentlichen Durchsicht der Fehlzeiten eine mögliche Schulpflichtverletzung in Deiner Klasse aufgepoppt ist. Können wir Dir Arbeit abnehmen?");
        zieldatei.Add(@"");
        zieldatei.Add(@"**Fragen & Antworten**");
        zieldatei.Add(@"");
        zieldatei.Add(
            @"  * :?: Was ist das Ziel dieser Seite? :!: Kritische Fälle erkennen, Reaktionszeiten verkürzen, Klassenleitungen Arbeit abnehmen, SuS signalisieren, dass wir hinschauen.");

        zieldatei.Add(
            @"  * :?: Wie oft soll ich mahnen? :!: Nach der Mahnung folgt i.d.R. die Teilkonferenz oder das Bußgeldverfahren. Wenn die letzte Mahnung sehr lange her ist, kommt eine weitere Mahnung in Betracht. ");

        zieldatei.Add(
            @"  * :?: Was, wenn die Zahlen nicht stimmen? :!: Dann gerne melden bei [[chat>stefan.baeumer|Stefan Bäumer]].");

        zieldatei.Add(
            @"  * :?: Was kann ich tun, wenn der Fall bereits der BezReg vorliegt? :!: Bei dauerhaft fehlenden SuS müssen die Klassenleitungen keine weiteren Aufgaben übernehmen, wenn der Fall der Bezirksregierung Münster gemeldet wurde.");
        
        zieldatei.Add(
            @"  * :?: Muss ich eine irgendwem eine Rückmeldung zu den Fällen in meiner Klasse geben? :!: Nein. Eine Rückmeldung ist nicht notwendig. Wer Fragen hat, kann sich natürlich immer melden: [[chat>stefan.baeumer|Stefan Bäumer]].");
        zieldatei.Add(@"");
        zieldatei.Add(@"");
        zieldatei.Add(@"");
        zieldatei.Add(@"");
        zieldatei.Add(@"===== Tabelle Schulpflichtüberwachung KW " + ISOWeek.GetWeekOfYear(DateTime.Now) + "=====");
        zieldatei.Add(@"");

        zieldatei.Add("<searchtable>");
        zieldatei.Add("^  Klasse  ^  Klassenleitung  ^  Name  ^  Alter am 1.Schultag im SJ " + Global.AktSj[0] + "/" +
                   Global.AktSj[1] +
                   "  ^  bisherige Maßnahmen  ^  Aussage  ^Womit können wir Arbeit abnehmen?  ^");

        string teamsChatLink =
            "chats>sina.milewski@berufskolleg-borken.de,stefan.gantefort@berufskolleg-borken.de,ursula.moritz@berufskolleg-borken.de,";
        var mailliste =
            "mailto:sina.milewski@berufskolleg-borken.de;stefan.gantefort@berufskolleg-borken.de;ursula.moritz@berufskolleg-borken.de;";

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Datei Schulpflichtüberwachung erstellen ...", ctx =>
        {
            foreach (var kl in (from k in this.OrderBy(x => x.Klasse) select k.Klasse).Distinct().ToList())
            {
                var klassenlehrerKürzel = (from k in klasses where k.Name == kl select k.Klassenlehrer).FirstOrDefault();
                var klassenlehrerKürzels = gpu003.Where(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return kl == dict["Field1"].ToString();
                    }).Select(rec =>
                    {
                        var dict = (IDictionary<string, object>)rec;
                        return dict["Field30"].ToString();
                    }).FirstOrDefault()?.Split(',').Select(k => k.Trim()).ToList() ?? new List<string>();

                var Klassenlehrer = lehrers.FirstOrDefault(l => l.Kürzel == klassenlehrerKürzel);
                var Klassenlehrers = new Lehrers();
                foreach (var klassenlehrerKü in klassenlehrerKürzels)
                {
                    var klassenlehrer = lehrers.FirstOrDefault(l => l.Kürzel == klassenlehrerKü.ToString());
                    if (klassenlehrer != null)
                    {
                        Klassenlehrers.Add(klassenlehrer);
                    }
                }
                

                foreach (var student in this.OrderBy(x => x.Nachname))
                {
                    if (student.Klasse == kl)
                    {
                        var name = student.Vorname!.Substring(0, 2) + "." + student.Nachname!.Substring(0, 2);

                        if (student.Nachname.StartsWith("Li") && student.Vorname.StartsWith("De"))
                        {
                            string aa = "";
                        }

                        // Geburtsdatum der Person
                        DateTime geburtsdatum = DateTime.Parse(student.Geburtsdatum!);

                        // Datum, an dem das Alter berechnet werden soll
                        DateTime ersterSchultag = new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);

                        int alter = ersterSchultag.Year - geburtsdatum.Year;

                        // Prüfen, ob der Geburtstag nach dem 1. August liegt, um das Alter korrekt anzupassen
                        if (geburtsdatum > ersterSchultag.AddYears(-alter))
                        {
                            alter--;
                        }

                        student.GetJüngsteMaßnahmeInDiesemSj();

                        // Für SuS ohne bisherige Maßnahme: Alle Fehlstunden, die noch nicht verjährt sind.


                        /*
                         * |------------|-------------------|-----------------------|-------------------|
                         *            Fehlzeit1           Fehlzeit2              Fehlzeit3            jetzt
                         *            21.8.                24.8.                   27.8.               28.8.
                         *            6 Stunden            5 Stunden               4 Stunden
                         *
                         *                      |<----------------------------------------------------->|
                         *                            Verjährung oder Zeit seit Maßnahme 6 Tage
                         *
                         *                                               |<---------------------------->|
                         *                                                     Schonfrist für KL
                         *                                                     zur Behandlung
                         *                                                     von Fehlzeiten
                         *                                                     3 Tage
                         *
                         *                       |-----------------------|
                         *                        Da die Fehlzeit2 innerhalb der Verjährung
                         *                        aber vor der Schonfrist liegt, werden
                         *                        alle Fehlzeiten (auch die in der Schonfrist) gewarnt.
                         */

                        student.GetF2(verjaehrungUnbescholtene, schonfrist);
                        student.GetF3(verjaehrungUnbescholtene, schonfrist);
                        student.GetF2PlusF3(verjaehrungUnbescholtene, schonfrist);
                        student.GetF2M(verjaehrungUnbescholtene, schonfrist);
                        student.GetF2MplusF3();

                        var aussage = "";
                        var mahnung = "";
                        var mahnungWikiLink = "";
                        var attestpflicht = "";
                        var attestpflichtWikiLink = "";
                        var teilkonferenz = "";
                        var bußgeldverfahren = "";

                        var anzahlMassnahmenInDiesemSj = student.Massnahmen.Where(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return DateTime.Parse(dict["Datum"].ToString()!) >
                                   new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);
                        }).Count();

                        // Wenn es noch keine Maßnahme in diesem SJ gab, ...
                        if (anzahlMassnahmenInDiesemSj == 0)
                        {
                            // ... und wenn es eine F2 gibt ...
                            if (student.F2 > 0)
                            {
                                if (student.F2PlusF3 > warnungAbAnzahl)
                                {
                                    // ... dann werden F2 und F3 angemahnt.

                                    aussage += student.F2PlusF3 + " unent. Fehlst. in den letzten " +
                                               verjaehrungUnbescholtene + " Tagen. ";
                                    mahnung = student.GetUrl("Mahnungen");
                                    mahnungWikiLink = student.GetWikiLink("Mahnung", student.F2PlusF3);
                                    attestpflicht = student.GetUrl("Attestpflicht");
                                    attestpflichtWikiLink = student.GetWikiLink("Attestpflicht", student.F2PlusF3);
                                }
                            }
                        }

                        var schonMassnahmen = student.Massnahmen.Where(rec =>
                        {
                            var dict = (IDictionary<string, object>)rec;
                            return DateTime.Parse(dict["Datum"].ToString()!) >
                                   new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);
                        }).Count();

                        // Wenn es schon Maßnahmen gab, ...
                        if (schonMassnahmen > 0)
                        {
                            // ... und wenn es eine F2M gibt ...
                            if (student.F2M > 0)
                            {
                                // ... dann werden F2M und F3 angemahnt.

                                var dictS = (IDictionary<string, object>)student.JuengsteMassnahmeInDiesemSj;

                                aussage += student.F2MplusF3 + " unent. Fehlstd. seit " +
                                           dictS["Vermerkart"].ToString() + "(" +
                                           dictS["Datum"].ToString() + ").";

                                if (dictS["Vermerkart"].ToString() == "Mahnung")
                                {
                                    if (alter < 18)
                                    {
                                        bußgeldverfahren =
                                            @"\\ [[eskalationsstufen_erzieherische_einwirkung_ordnungsmassnahmen:bussgeldverfahren:start|Bußgeldverfahren]]";
                                    }
                                    else
                                    {
                                        teilkonferenz =
                                            @"\\ [[eskalationsstufen_erzieherische_einwirkung_ordnungsmassnahmen:bussgeldverfahren:start|Teilkonferenz]]";
                                    }
                                }
                            }
                        }

                        student.MaßnahmenAlsWikiLinkAufzählung = student.GetMaßnahmenAlsWikiLinkAufzählung();
                        var klassenleitungString = "";

                        if (aussage.Length > 0)
                        {
                            //klassenleitungString += Klassenlehrer.Kürzel + ",";
                            klassenleitungString += string.Join(",", Klassenlehrers.Select(klassenlehrer => klassenlehrer.Kürzel));

                            if (!mailliste.Contains(Klassenlehrer.Mail))
                            {
                                mailliste += Klassenlehrer.Mail + ";";
                                if(Klassenlehrer != null)
                                    zieldatei.Lehrers.Add(Klassenlehrer);
                            }

                            if (!teamsChatLink.Contains(Klassenlehrer.Mail))
                            {
                                teamsChatLink += string.Join(",", Klassenlehrers.Select(klassenlehrer => klassenlehrer.Mail));
                            }
                        }                        
                    }
                }
            }
        });

        zieldatei.Add("</searchtable>");

        teamsChatLink = teamsChatLink.TrimEnd(',') + "&topicName=Schulpflichtüberwachung KW " + ISOWeek.GetWeekOfYear(DateTime.Now) +
                        "&message=Bitte beachten: https://bkb.wiki/schulpflichtueberwachung";

        foreach (var zeile in zieldatei)
        {
            zeile.Replace("Teams", @"[[" + teamsChatLink + @"|Teams]]");
            //File.AppendAllText(tempdatei, z, Encoding.UTF8);
        }

        Global.ZeileSchreiben($"Datei Schulpflichtüberwachung erstellt. Zeilen:", zieldatei.Count().ToString());

        foreach (var aktion in zieldatei.Funktionen)
            aktion(zieldatei);
    }

    public List<dynamic> Reliabmelder()
    {
        UTF8Encoding utf8NoBom = new UTF8Encoding(false);
        var records = new List<dynamic>();
        return records;
    }


    public List<dynamic>? SchuelerFehlstunden(Dateien dateien, IConfiguration configuration)
    {
        var records = new List<dynamic>();

        foreach (var student in this.OrderBy(datei => datei.Klasse).ThenBy(zeile => zeile.Nachname).ToList())
        {
            List<dynamic> absencePerStudents =
                dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.ToLower().Contains("absenceperstudent"))!
                    .ToList();

            foreach (var recAbs in absencePerStudents)
            {
                var dictAbs = (IDictionary<string, object>)recAbs;

                if (dictAbs["Schüler*innen"].ToString()!.Contains(student.Nachname!) &&
                    dictAbs["Schüler*innen"].ToString()!.Contains(student.Vorname!))
                {
                    if (dictAbs["Datum"] != null)
                    {
                        int fehlstd = string.IsNullOrEmpty(dictAbs["Fehlstd."].ToString())
                            ? 0
                            : Convert.ToInt32(dictAbs["Fehlstd."].ToString());
                        var entschuldigt = "N";
                        if (dictAbs["Status"] != null)
                        {
                            if (dictAbs["Status"].ToString() == "entsch.")
                            {
                                entschuldigt = "J";
                            }
                        }

                        if (fehlstd > 0)
                        {
                            dynamic record = new ExpandoObject();
                            record.Nachname = $"{student.Nachname}#{student.Klasse}";
                            record.Vorname = student.Vorname;
                            record.Geburtsdatum = student.Geburtsdatum;
                            record.Klasse = student.Klasse;
                            record.Jahr = Global.AktSj[0];
                            record.Abschnitt = configuration["Abschnitt"];
                            record.Datum = dictAbs["Datum"].ToString();
                            record.Fach = "";
                            record.von_StdPUNKT = "";
                            record.bis_StdPUNKT = "";
                            record.FehlstdPUNKT = fehlstd;
                            record.Entschuldigt = entschuldigt;
                            record.Lehrkraft = "";
                            records.Add(record);
                        }
                    }
                }
            }
        }

        return records;
    }

    public void Fotos(IConfiguration configuration)
    {
        var pfad = Path.Combine(configuration["PfadDownloads"], "Fotos-Input");

        if (!Directory.Exists(pfad))
        {
            Directory.CreateDirectory(pfad);
            Console.WriteLine($"   Der Ordner '{pfad}' wurde erstellt.");
        }

        Console.WriteLine("   Hier müssen die Fotos liegen: " + pfad);

        do
        {
            Console.WriteLine("");
            Console.Write("   Klasse eingeben: ");
            var klasse = "";
            do
            {
                while (Console.KeyAvailable) Console.ReadKey(true);

                klasse = Console.ReadLine();
            } while (!this.Any(x => x.Klasse == klasse.ToUpper()));

            Console.WriteLine("");
            List<Student> sortierteKlasse = this
                .Where(x => x.Klasse == klasse.ToUpper())
                .OrderBy(x => x.Nachname)
                .ThenBy(x => x.Vorname).ToList();
            Console.WriteLine("   Jetzt die Schüler*innen in dieser Reihenfolge fotografieren:");
            var z = 1;
            foreach (var student in sortierteKlasse)
            {
                Console.WriteLine("    " + z.ToString().PadLeft(2) + ". " + student.Nachname + ", " + student.Vorname);
                z++;
            }

            Console.WriteLine("");
            Console.WriteLine("   Liegen die " + sortierteKlasse.Count() + " Fotos der Klasse " + klasse.ToUpper() +
                              " im Ordner '" + pfad + "'? Dann ENTER");

            while (Console.KeyAvailable) Console.ReadKey(true);

            var x = Console.ReadKey();

            if (x.Key != ConsoleKey.Enter)
            {
                break;
            }

            // Alle jpg-Dateien im Ordner finden
            var jpgDateien = Directory.GetFiles(pfad, "*.*", SearchOption.TopDirectoryOnly)
                .Where(file => file.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
                               file.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            // Anzahl der Dateien ausgeben
            if (sortierteKlasse.Count() == jpgDateien.Length)
            {
                Console.WriteLine($"   Anzahl der JPG-Bilder im Ordner '{pfad}'");
                Console.WriteLine($"   stimmt mit der Anzahl der SuS in der Klasse überein: {jpgDateien.Length}");

                var klassenpfad = Path.Combine(configuration["PfadDownloads"], "Fotos", klasse.ToUpper());
                // Ordner "Klasse" erstellen, falls er nicht existiert
                if (!Directory.Exists(klassenpfad))
                {
                    Directory.CreateDirectory(klassenpfad);
                    Console.WriteLine($"   Der Ordner '{klassenpfad}' wurde erstellt.");
                }

                // Laufende Nummer und Dateiverarbeitung
                for (int i = 0; i < jpgDateien.Length; i++)
                {
                    string bildPfad = jpgDateien[i];
                    FileInfo bildInfo = new FileInfo(bildPfad);

                    // Bildinformationen ausgeben
                    Console.WriteLine($"    {i + 1}. {bildInfo.Name} - Erstellt am: {bildInfo.CreationTime}");

                    // Umbenennung vorbereiten
                    if (i < sortierteKlasse.Count())
                    {
                        string neuerDateiname =
                            $"{sortierteKlasse[i].Nachname}_{sortierteKlasse[i].Vorname}_{sortierteKlasse[i].Id}.jpg";
                        string neuerPfad = Path.Combine(klassenpfad, neuerDateiname);

                        // Datei verschieben und umbenennen
                        File.Move(bildPfad, neuerPfad);

                        Console.WriteLine("     verschoben nach" + neuerPfad);
                    }
                    else
                    {
                        Console.WriteLine(
                            $"    Keine weiteren Studenten in der Liste verfügbar für Datei: {bildInfo.Name}");
                    }

                }
            }
            else
            {
                Console.WriteLine(
                    $"   Anzahl der JPG-Bilder im Ordner '{pfad}' stimmt NICHT der Anzahl der SuS in der Klasse überein.");
            }
        } while (true);
    }

    private sealed class SchuelersMap : ClassMap<Student>
    {

    }

    public List<dynamic> AdressenImportieren(Dateien dateien)
    {
        var records = new List<dynamic>();
        var adressen = dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.Contains("Adressen"))!.ToList();

        foreach (var student in this)
        {
            var elternStattVundM = false;

            var adressenDiesesSchuelers = adressen.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Schüler: Vorname"].ToString() == student.Vorname &&
                       dict["Schüler: Nachname"].ToString() == student.Nachname &&
                       dict["Schüler: Geburtsdatum"].ToString() == student.Geburtsdatum.ToString();
            }).ToList();

            var vater = adressenDiesesSchuelers.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Adresse: Typ Adresse"].ToString() == "V";
            }).FirstOrDefault();

            var mutter = adressenDiesesSchuelers.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Adresse: Typ Adresse"].ToString() == "M";
            }).FirstOrDefault();

            if (vater != null && mutter != null)
            {
                var dictVater = (IDictionary<string, object>)vater;
                var dictMutter = (IDictionary<string, object>)mutter;

                // Wenn Mama und Papa dieselbe Adresse haben, ...
                if (dictVater!["Adresse: Straße"].ToString() == dictMutter!["Adresse: Straße"].ToString())
                {
                    // und der student nicht volljährig ist, ...
                    if (!student.IstVolljaehrig())
                    {
                        // ... wird ein Erzieher namens Eltern erzeugt, der als einziger angeschrieben wird.

                        dynamic record = new ExpandoObject();
                        record.Nachname = $"{student.Nachname}#{student.Klasse}";
                        record.Vorname = student.Vorname;
                        record.Geburtsdatum = student.Geburtsdatum;
                        record.Erzieherart = "Eltern";
                        record.AnredeLEERZEICHEN1PUNKTPerson = "";
                        record.TitelLEERZEICHEN1PUNKTPerson = "";
                        record.NachnameLEERZEICHEN1PUNKTPerson = student.Nachname;
                        record.VornameLEERZEICHEN1PUNKTPerson =
                            dictMutter!["Adresse: Name 2"] + " und " + dictVater!["Adresse: Name 2"];
                        record.AnredeLEERZEICHEN2PUNKTPerson = "";
                        record.TitelLEERZEICHEN2PUNKTPerson = "";
                        record.NachnameLEERZEICHEN2PUNKTPerson = "";
                        record.VornameLEERZEICHEN2PUNKTPerson = "";
                        record.Straße = dictVater!["Adresse: Straße"].ToString();
                        record.PLZ = dictVater!["Adresse: Plz"].ToString();
                        record.Ort = dictVater!["Adresse: Ort"].ToString();
                        record.Ortsteil = "";
                        record.EMINUSMailLEERZEICHEN1PUNKTLEERZEICHENPerson = dictMutter!["Adresse: E-Mail"].ToString();
                        record.Anschreiben = "J";
                        record.EMINUSMailLEERZEICHEN2PUNKTLEERZEICHENPerson =
                            string.IsNullOrEmpty(dictVater!["Adresse: E-Mail"].ToString())
                                ? ""
                                : dictVater!["Adresse: E-Mail"].ToString();
                        records.Add(record);
                        elternStattVundM = true;
                    }
                }
            }

            foreach (var rec in adressenDiesesSchuelers)
            {
                var dict = (IDictionary<string, object>)rec;
                dynamic record = new ExpandoObject();
                record.Nachname = $"{student.Nachname}#{student.Klasse}";
                record.Vorname = student.Vorname;
                record.Geburtsdatum = student.Geburtsdatum;
                record.Erzieherart = student.GetErzieherart(dict["Adresse: Typ Adresse"].ToString(),
                    dict["Adresse: Sorgeberechtigt (J/N)"].ToString(), dict["Schüler: Geschlecht"].ToString());
                record.AnredeLEERZEICHEN1PUNKTPerson = dict["Adresse: Anrede (Auflösung)"].ToString();
                record.TitelLEERZEICHEN1PUNKTPerson = "";
                record.NachnameLEERZEICHEN1PUNKTPerson = dict["Adresse: Name 1"].ToString();
                record.VornameLEERZEICHEN1PUNKTPerson = dict["Adresse: Name 2"].ToString();
                record.AnredeLEERZEICHEN2PUNKTPerson = "";
                record.TitelLEERZEICHEN2PUNKTPerson = "";
                record.NachnameLEERZEICHEN2PUNKTPerson = "";
                record.VornameLEERZEICHEN2PUNKTPerson = "";
                record.Straße = dict!["Adresse: Straße"].ToString();
                record.PLZ = dict!["Adresse: Plz"].ToString();
                record.Ort = dict!["Adresse: Ort"].ToString();
                record.Ortsteil = "";
                record.EMINUSMailLEERZEICHEN1PUNKTLEERZEICHENPerson = dict!["Adresse: E-Mail"].ToString();
                record.Anschreiben = Anschreiben(record.Erzieherart, student.IstVolljaehrig(), elternStattVundM);
                record.EMINUSMailLEERZEICHEN2PUNKTLEERZEICHENPerson = "";
                if (!(!student.IstVolljaehrig() && record.Erzieherart == ""))
                {
                    records.Add(record);
                }
            }
        }

        return records;
    }

    private static dynamic Anschreiben(dynamic erzieherart, bool volljaehrig, bool elternStattVundM)
    {
        // Minderjährige werden nicht informiert
        if (!volljaehrig && erzieherart == "")
        {
            return "N";
        }

        // Volljährige immer
        if (volljaehrig && erzieherart.StartsWith("Schüler"))
        {
            return "J";
        }

        if (erzieherart == "Eltern")
        {
            return "J";
        }

        // Vater oder Mutter Minderjähriger, wenn es Eltern gibt
        if (!volljaehrig && erzieherart != "" && elternStattVundM)
        {
            return "N";
        }

        // Vater oder Mutter Minderjähriger, ohne Eltern
        if (!volljaehrig && erzieherart != "" && !elternStattVundM)
        {
            return "J";
        }

        return "N";
    }

    public List<dynamic> AdressenBetriebeImportieren(Dateien dateien)
    {
        var records = new List<dynamic>();
        List<dynamic> adressen =
            dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.Contains("Adressen"))!.ToList();
        List<dynamic> datums =
            dateien.FirstOrDefault(datei => datei.UnterordnerUndDateiname.Contains("DatumsAusAtlantis"))!.ToList();

        foreach (var student in this)
        {
            var adressenDiesesSchuelers = adressen.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Schüler: Vorname"].ToString() == student.Vorname &&
                       dict["Schüler: Nachname"].ToString() == student.Nachname &&
                       dict["Schüler: Geburtsdatum"].ToString() == student.Geburtsdatum.ToString();
            }).ToList();

            var datumsDiesesSchuelers = datums.Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Vorname"].ToString() == student.Vorname &&
                       dict["Nachname"].ToString() == student.Nachname &&
                       DateTime.Parse(dict["Geburtsdatum"].ToString()).ToString("dd.MM.yyyy") ==
                       student.Geburtsdatum.ToString();
            }).LastOrDefault();


            var dictD = (IDictionary<string, object>)datumsDiesesSchuelers;

            foreach (var rec in adressenDiesesSchuelers)
            {
                var dictA = (IDictionary<string, object>)rec;

                if (!string.IsNullOrEmpty(dictA["Betrieb: Name 1"].ToString()))
                {
                    dynamic record = new ExpandoObject();
                    record.Nachname =
                        student.Nachname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                    record.Vorname =
                        student.Vorname; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                    record.Geburtsdatum =
                        student.Geburtsdatum; // Wenn die ersten 3 Spalten leer sind, dann wird der Betrieb ohne Zuordnung importiert
                    record.Adressart = "Betrieb";
                    record.Name1 = dictA["Betrieb: Name 1"].ToString();
                    record.Name2 = dictA["Betrieb: Name 2"].ToString();
                    record.Straße = dictA["Betrieb: Straße"].ToString();
                    record.PLZ = dictA["Betrieb: Plz"].ToString();
                    record.Ort = dictA["Betrieb: Ort"].ToString();
                    record.EINSPUNKTLEERZEICHENTelPUNKTMINUSNrPUNKT = dictA["Betrieb: Telefon 1"].ToString();
                    record.ZWEIPUNKTLEERZEICHENTelPUNKTMINUSNrPUNKT = dictA["Betrieb: Telefon 2"].ToString();
                    record.EMINUSMail = dictA["Betrieb: E-Mail-Adresse"].ToString();
                    record.BetreuerLEERZEICHENNachname = dictA["Betrieb: Ansprechpartner"].ToString();
                    record.BetreuerLEERZEICHENVorname = "";
                    record.BetreuerLEERZEICHENAnrede = "";
                    record.BetreuerLEERZEICHENTelPUNKTMINUSNrPUNKT = "";
                    record.BetreuerLEERZEICHENEMINUSMail = "";
                    record.BetreuerLEERZEICHENAbteilung = "";
                    record.Vertragsbeginn = dictD != null && !string.IsNullOrEmpty(dictD["Ausbildungsbeginn"]?.ToString())
                        ? DateTime.Parse(dictD["Ausbildungsbeginn"]?.ToString() ?? string.Empty).ToString("dd.MM.yyyy")
                        : "";
                    record.Vertragsende = dictD != null && !string.IsNullOrEmpty(dictD["Ausbildungsende"]?.ToString())
                        ? DateTime.Parse(dictD["Ausbildungsende"]?.ToString() ?? string.Empty).ToString("dd.MM.yyyy")
                        : "";

                    var gibtEsSchon = records.Where(recB =>
                    {
                        var dictB = (IDictionary<string, object>)recB;
                        return dictB["Nachname"].ToString() == dictA["Schüler: Nachname"].ToString() &&
                               dictB["Vorname"].ToString() == dictA["Schüler: Vorname"].ToString() &&
                               dictB["Name1"].ToString() == dictA["Betrieb: Name 1"].ToString();
                    }).Any();

                    if (!gibtEsSchon)
                    {
                        records.Add(record);
                    }
                }
            }
        }

        return records;
    }

    public Klassen GetKlassen()
    {
        Klassen gefilterteklassen = new Klassen();

        foreach (var student in this)
        {
            if (!gefilterteklassen.Any(x => x.Name == student.Klasse))
            {
                gefilterteklassen.Add(new Klasse { Name = student.Klasse });
            }

            if (!gefilterteklassen.Any(x => x.Name == student.Klasse))
            {
                gefilterteklassen.Add(new Klasse { Name = student.Klasse });
            }
        }

        return gefilterteklassen;
    }

    public Students GetInteressierende(string klasse)
    {
        var students = new Students();
        foreach (var stu in this)
        {
            if (klasse.Contains("alle") || klasse == stu.Klasse || klasse == stu.KlasseWebuntis)
            {
                students.Add(stu);
            }
        }

        return students;
    }

    public bool IsAllesOk()
    {
        var stati = this.Select(x => x.Status).Distinct().ToList();

        if (Count > 0 && stati.Count > 1)
        {
            return true;
        }

        return false;
    }

    internal void GetPfadAtlantisFotos(IConfiguration configuration)
    {
        // Ordner "Fotos" unterhalb von Global.PfadExportdateien
        var fotosOrdner = Path.Combine(configuration["PfadDownloads"] ?? "", "Fotos");

        // Prüfen, ob der Ordner existiert
        if (!Directory.Exists(fotosOrdner))
        {
            Console.WriteLine($"Der Ordner [bold red]{fotosOrdner}[/] existiert nicht.");
            return;
        }

        // Alle .jpg-Dateien im Ordner "Fotos" einlesen
        var fotoDateien = Directory.GetFiles(fotosOrdner, "*.jpg").ToList();

        // Durchlaufe alle Students
        foreach (var student in this)
        {
            // Prüfe, ob eine Datei die externe ID-Nummer enthält
            var passendeDatei = fotoDateien.FirstOrDefault(foto => Path.GetFileName(foto).Contains("_" + student.ExterneIdNummer + ".jpg"));

            if (!string.IsNullOrEmpty(passendeDatei))
            {
                // Setze die AtlantisFotoUrl-Eigenschaft
                student.ZielFotoPfad = passendeDatei;
            }
        }
    }

    internal void GetPfadNeueFotos()
    {
        do
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("  Ziehen Sie genau " + this.Count() + " Schülerfotos in folgender Reihenfolge aus dem Explorer auf die App:");

            var nr = 0;
            foreach (var student in this)
            {
                nr++;
                Console.WriteLine(nr.ToString().PadLeft(4) + ". " + student.Nachname + "," + student.Vorname);
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("  Danach ENTER:");
            while (Console.KeyAvailable) Console.ReadKey(true);

            var picsString = Console.ReadLine();
            var pics = new List<string>();

            // Regex für Pfade in einfachen oder doppelten Anführungszeichen
            string pattern = @"['""]([^'""]+)['""]"; // Alles zwischen einfachen oder doppelten Quotes

            // Regex.Match für mehrere Gruppen
            MatchCollection matches = Regex.Matches(picsString, pattern);

            foreach (Match match in matches)
            {
                if (match.Groups.Count > 1) // Überprüfen, ob ein Pfad gefunden wurde
                {
                    pics.Add(match.Groups[1].Value);
                }
            }

            if (pics.Count() == nr)
            {
                break;
            }

            Console.ForegroundColor = ConsoleColor.Red;

            Console.WriteLine("   Es müssen exakt " + this.Count() + " Bilder hochgeladen werden. Es wurden " +
                              pics.Count() + " Bilder hochgeladen.");
            Console.WriteLine(
                "   Die Anzahl stimmt nicht überein. Es dürfen nur jpg oder jpeg-Dateien hochgeladen werden.");
            Console.WriteLine("");
        } while (true);
    }

    internal void Pfad2FotoStream()
    {
        foreach (var student in this)
        {
            if (!string.IsNullOrEmpty(student.ZielFotoPfad))
            {
                student.Pfad2FotoStream();
            }
        }
        Global.ZeileSchreiben("Fotos in Stream umgewandelt", this.Where(x => x.ZielFotoPfad != null).Count().ToString(), ConsoleColor.Green, ConsoleColor.Black);
    }

    internal void GetPfadDokumentenverwaltung(IConfiguration configuration)
    {
        Global.Konfig("PfadDokumentenverwaltung", Global.Modus.Update, configuration);

        foreach (var student in this)
        {
            student.GetPfadDokumentenverwaltung(configuration);
        }
    }

    internal void ErstellenPfadDokumentenverwaltung(IConfiguration configuration)
    {
        foreach (var student in this)
        {
            student.ErstellePfadDokumentenverwaltung(configuration);
        }
    }

    internal void BilderNachPfadDokumentenverwaltungKopieren(IConfiguration configuration)
    {
        foreach (var student in this)
        {
            student.BilderNachPfadDokumentenverwaltungKopieren(configuration);
        }
    }

    internal void KlassenordnerErstellenFotosOrdnerÖffnen(IConfiguration configuration)
    {
        var pfadDownloads = configuration["PfadDownloads"];

        // Ordner "Fotos" unterhalb von Global.PfadExportdateien
        var fotosOrdner = Path.Combine(pfadDownloads, "Fotos");

        var verschiedeneKlassen = this.Select(x => x.Klasse).Distinct().ToList();
        foreach (var klasse in verschiedeneKlassen)
        {
            var pfad = Path.Combine(fotosOrdner, klasse);
            if (!Directory.Exists(pfad))
            {
                Directory.CreateDirectory(pfad);
                Global.ZeileSchreiben("Ordner " + pfad, "erstellt.", ConsoleColor.Green, ConsoleColor.Black);
            }
            else
            {
                Global.ZeileSchreiben("Ordner " + pfad, "existiert bereits.", ConsoleColor.Green, ConsoleColor.Black);
            }
            // Öffne das Verzeichnis im Explorer, sofern der Ordner noch nicht geöffnet ist
            /*if (!Global.IsExplorerOpen(fotosOrdner))
            {
                System.Diagnostics.Process.Start("explorer.exe", fotosOrdner);
            }*/
        }
    }

    public void KlassenListenAnzeigen(IConfiguration configuration)
    {
        var verschiedeneKlassen = configuration["Klassen"].Split(",").ToList();
        var pfadDownloads = configuration["PfadDownloads"];

        foreach (var klasse in verschiedeneKlassen)
        {
            var panel = new Panel($"[{Global.GetColor(Global.ColorHinweise)}]Wichtig:[/] Wenn ein*e Schüler*in fehlt, fotografieren Sie die weiße Wand, damit die Anzahl stimmt." +
                                    $"\nAm Ende müssen Sie die Fotos händisch in den Ordner [{Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(pfadDownloads, "Fotos", klasse) + "[/] verschieben.")
                            .Header($"[{Global.GetColor(Global.ColorInfoBox)}] Jetzt {this.Where(x => x.Klasse == klasse).Count()} Fotos der {klasse} machen! [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Spectre.Console.Color.DodgerBlue1);

            //AnsiConsole.Write(panel);

            int i = 1;
            var table = new Table();
            table.Title = new TableTitle($"Jetzt {this.Where(x => x.Klasse == klasse).Count()} Fotos machen und in [{Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(pfadDownloads, "Fotos", klasse) + "[/] ablegen.");
            table.Expand();
            var orangeStyle = new Style(foreground: Spectre.Console.Color.Orange1);
            table.AddColumn(new TableColumn($"[{Global.GetColor(Global.ColorInfoBox)}]Nr.[/]"));
            table.AddColumn(new TableColumn($"[{Global.GetColor(Global.ColorInfoBox)}]Nachname[/]"));
            table.AddColumn(new TableColumn($"[{Global.GetColor(Global.ColorInfoBox)}]Vorname[/]"));
            table.AddColumn(new TableColumn($"[{Global.GetColor(Global.ColorInfoBox)}]Geburtsdatum[/]"));
            table.AddColumn(new TableColumn($"[{Global.GetColor(Global.ColorInfoBox)}]Klasse[/]"));
            table.BorderColor(Spectre.Console.Color.DodgerBlue1);
            table.SquareBorder();

            foreach (var student in this.Where(x => x.Klasse == klasse))
            {
                table.AddRow(i.ToString().PadLeft(2), student.Nachname, student.Vorname, student.Geburtsdatum, student.Klasse);
                i++;
            }
            AnsiConsole.Write(table);

            /*AnsiConsole.Write(new Panel($"Verschieben Sie die [{Global.GetColor(Global.ColorZahlen)}]{this.Where(x => x.Klasse == klasse).Count()}[/] Fotos nach [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads, "Fotos", klasse)}[/].")
                            .Header($"[{Global.GetColor(Global.ColorInfoBox)}] [{Global.GetColor(Global.ColorZahlen)}]{this.Where(x => x.Klasse == klasse).Count()}[/] Fotos gemacht? [/]")
                            .HeaderAlignment(Justify.Left)
                            .SquareBorder()
                            .Expand()
                            .BorderColor(Spectre.Console.Color.DodgerBlue1));*/
            Thread.Sleep(1000);
        }
    }

    internal void FotosFürUploadNachSchildAuswählen(IConfiguration configuration)
    {
        if (this.Count() == 0)
        {
            throw new ArgumentException("Es wurden keine Klassen für den Fotoimport ausgewählt.");
        }

        configuration = Global.Konfig("PfadFotosImSchILD-Ordner", Global.Modus.Update, configuration);

        // Alle *.jpg-Dateien aus dem Ordner (und allen Unterordnern) PfadFotosImSchILD-Ordner werden eingelesen.
        var alleFotosImSchildOrdner = Directory.GetFiles(configuration["PfadFotosImSchILD-Ordner"], "*.jpg", SearchOption.AllDirectories);

        var gefilterteFotosImSchildOrdner = new List<string>();
        var gefilterteFotosImFotos_Aus_SchildOrdner = new List<string>();

        // Filter alle Fotos, die auf IStudents matchen.
        foreach (var foto in alleFotosImSchildOrdner)
        {
            // Dateiname ohne Endung
            var dateiName = Path.GetFileNameWithoutExtension(foto);
            var student = this.FirstOrDefault(s => !string.IsNullOrEmpty(dateiName) && !string.IsNullOrEmpty(s.MailSchulisch) && s.MailSchulisch.StartsWith(dateiName));

            if (student != null)
            {
                gefilterteFotosImSchildOrdner.Add(foto);
            }
        }

        Global.ZeileSchreiben("Vorhandene Fotos zur obigen Auswahl in " + configuration["PfadFotosImSchILD-Ordner"], gefilterteFotosImSchildOrdner.Count().ToString(), ConsoleColor.Green, ConsoleColor.Black);


        configuration = Global.Konfig("PfadFotosAusSchild", Global.Modus.Update, configuration);

        // Alle *.jpg-Dateien aus dem Ordner (und allen Unterordnern) pfadFotosAusSchild werden eingelesen.
        var fotosAusSchildOrdner = Directory.GetFiles(configuration["PfadFotosAusSchild"], "*.jpg", SearchOption.AllDirectories);

        // Filter alle Fotos, die auf IStudents matchen.
        foreach (var foto in fotosAusSchildOrdner)
        {
            // Dateiname ohne Endung
            var dateiName = Path.GetFileNameWithoutExtension(foto);
            var student = this.FirstOrDefault(s => dateiName.StartsWith(s.Nachname) && dateiName.Contains(s.Vorname) && dateiName.Contains(s.Geburtsdatum));

            if (student != null)
            {
                gefilterteFotosImFotos_Aus_SchildOrdner.Add(foto);
            }
        }

        Global.ZeileSchreiben("Vorhandene Fotos zur obigen Auswahl in " + configuration["PfadFotosAusSchild"], gefilterteFotosImFotos_Aus_SchildOrdner.Count().ToString(), ConsoleColor.Green, ConsoleColor.Black);

        foreach (var foto in gefilterteFotosImSchildOrdner)
        {
            // Dateiname ohne Endung
            var dateiName = Path.GetFileNameWithoutExtension(foto);
            var student = this.FirstOrDefault(s => !string.IsNullOrEmpty(dateiName) && !string.IsNullOrEmpty(s.MailSchulisch) && s.MailSchulisch.StartsWith(dateiName));

            if (student == null) continue;

            // Wenn der student schon ein FotoAusSchild-Foto hat, erkennbar daran, 
            // dass der Dateiname Vorname, Nachname und Geburtsdatum durch Unterstrich getrennt enthält, 
            // dann wird übersprungen.
            var fotoBereitsInSchildVorhanden = gefilterteFotosImFotos_Aus_SchildOrdner.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).Contains(student.Nachname) && Path.GetFileNameWithoutExtension(f).Contains(student.Vorname) && Path.GetFileNameWithoutExtension(f).Contains(student.Geburtsdatum));
            if (!string.IsNullOrEmpty(fotoBereitsInSchildVorhanden)) continue;

            // Wenn der student noch kein Foto hat, wird das aktuelle Foto zugewiesen.
            student.ZielFotoPfad = foto;
        }
    }

    internal Students OhneWebuntisFoto(IConfiguration configuration, string fotosTxt)
    {
        if (!File.Exists(fotosTxt))
        {
            File.Create(fotosTxt).Close();
        }

        // Lies aus der Datei fotos.txt im current folder alle fotos in eine List<string>
        var fotoIst = File.ReadAllLines(fotosTxt).ToList();

        Students students = new Students();

        foreach (var student in this)
        {
            if (string.IsNullOrEmpty(student.ExterneIdNummer))
            {
                if (!fotoIst.Contains(student.Id))
                {
                    students.Add(student);
                }
            }
        }
        Global.ZeileSchreiben("SUS ohne Webuntis-Foto", students.Count().ToString(), ConsoleColor.Green, ConsoleColor.White);
        return students;
    }

    internal void FotosFürWebuntisZippen(IConfiguration configuration, string zipPfad, string fotosTxt, List<string> importhinweise = null)
    {
        try
        {
            using (FileStream zipStream = File.Create(zipPfad))
            using (ZipOutputStream zip = new ZipOutputStream(zipStream))
            {
                zip.SetLevel(0); // Keine Komprimierung

                byte[] buffer = new byte[4096];

                foreach (var student in this)
                {
                    var pfadDokumentenverwaltung = student.GetPfadDokumentenverwaltung(configuration);

                    var absoluterPfadZumBild = Path.Combine(pfadDokumentenverwaltung, student.Id + ".jpg");

                    if (File.Exists(absoluterPfadZumBild))
                    {
                        // Name der Datei im ZIP-Archiv
                        string dateiNameImZip = student.Id + ".jpg";

                        // Zip-Eintrag erstellen
                        ZipEntry entry = new ZipEntry(dateiNameImZip)
                        {
                            DateTime = DateTime.Now,
                            CompressionMethod = CompressionMethod.Stored // Keine Komprimierung
                        };

                        zip.PutNextEntry(entry);

                        // Datei in das ZIP-Archiv schreiben
                        using (FileStream dateiStream = File.OpenRead(absoluterPfadZumBild))
                        {
                            int bytesRead;
                            while ((bytesRead = dateiStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                zip.Write(buffer, 0, bytesRead);
                            }
                        }

                        zip.CloseEntry();

                        // Bild in fotos.txt schreiben
                        File.AppendAllText(fotosTxt, student.Id + Environment.NewLine);
                    }
                }

                zip.IsStreamOwner = true;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fehler beim Zippen: " + ex.Message);
        }
        finally
        {
            var rechteSeite = importhinweise != null && importhinweise.Any() ? string.Join("\n", importhinweise) : "";
            Global.ZeileSchreiben(zipPfad, rechteSeite, ConsoleColor.White, ConsoleColor.Blue);
        }
    }

    internal string GetArtUndZahlen()
    {
        var statusstring = $"[{Global.GetColor(Global.ColorZahlen)}]" + this.Count().ToString() + "[/] Schüler*innen: ";

        var i = 0;
        var zeile = new List<string>();
        zeile.Add(this.Count().ToString());

        if (this == null || this.Count == 0)
        {
            throw new Exception($"Keine Schülerdaten gefunden. Notwendige Dateien heruntergeladen? \nDateien veraltet? Drücken Sie [{Global.GetColor(Global.ColorActionInMenüs)}]e[/], um in den Einstellungen festzulegen, ab wann Dateien als veraltet verworfen werden.");
        }

        foreach (var status in this.Select(x => x.Status).Distinct().OrderBy(x => x).ToList())
        {
            statusstring += $"[{Global.GetColor(Global.ColorZahlen)}]{this.Count(x => x.Status == status)}[/]";

            switch (status)
            {
                case "0":
                    statusstring += " in Neuaufnahme, ";
                    break;
                case "2":
                    statusstring += " aktive, ";
                    break;
                case "6":
                    statusstring += " extern, ";
                    break;
                case "8":
                    statusstring += " mit Abschluss, ";
                    break;
                case "9":
                    statusstring += " abgegangen, ";
                    break;
                default:
                    statusstring += " beurlaubt, ";
                    break;
            }
        }

        if (this.Select(x => x.Status).Distinct().Count() == 1)
        {
            return $"[{Global.GetColor(Global.ColorZahlen)}]{this.Count().ToString()}[/] Schüler*innen:[springGreen2] nur aktive exportiert[/]";
        }

        return statusstring.TrimEnd(' ').TrimEnd(',').TrimEnd(' ').TrimEnd(',');
    }

    /// <summary>
    /// Wenn sich die Belegung durch den Schüler mit dem Zeitraum von-bis überschneidet, wird der Schüler zurückgegeben.
    /// Wenn beide Zeiträume sich nicht überschneiden, wird der Schüler nicht zurückgegeben.
    /// Wenn es um die Statistik geht, fallen von und bis auf einen einzigen Tag.
    /// </summary>
    /// <param name="von"></param>
    /// <param name="bis"></param>
    /// <param name="klasse"></param>
    /// <param name="schuelergruppe"></param>
    /// <param name="studentgroupStudents"></param>
    /// <returns></returns>
    internal Students Filter(IConfiguration configuration, Global.Zweck zweck, string klasse, string? schuelergruppe = null, List<dynamic> studentgroupStudents = null)
    {
        var studentsGefiltert = new Students();

        // Fall 1: Schülergruppe leer
        if (string.IsNullOrEmpty(schuelergruppe) && !string.IsNullOrEmpty(klasse))
        {
            // Alle Schüler der angegebenen Klasse hinzugefügt.
            foreach (var student in this.Where(x => x.Klasse == klasse))
            {
                // Doppelte Einträge vermeiden
                if (!studentsGefiltert.Any(x => x.Id == student.Id))
                {
                    studentsGefiltert.Add(student);
                }
            }
        }

        // Fall 2: Schülergruppe angegeben        
        if (!string.IsNullOrEmpty(schuelergruppe))
        {
            // Wenn eine Schülergruppe angegeben ist, nur die Schüler dieser Gruppe hinzufügen.
            var schuelerDieserGruppe = studentgroupStudents
                .Select(x => x as IDictionary<string, object>)
                .Where(dict =>
                    dict != null &&
                    dict["studentgroup.name"]?.ToString() == schuelergruppe)
                .ToList();

            foreach (var student in schuelerDieserGruppe)
            {
                // Wenn es den Schüler mit der Schülergruppe in den IStudents gibt ...
                if (this.Any(x => x.Id == student["studentId"].ToString()))
                {
                    // ... wird er hinzugefügt. 
                    var stu = this.Where(x => x.Id == student["studentId"].ToString()).FirstOrDefault();
                    if (stu != null)
                    {
                        // Nur Kursbelegungen, die sich mit dem Zeitraum von-bis überschneiden, werden zurückgegeben.
                        DateTime ersterSchultag = new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1); // Beispiel für den ersten Schultag
                        DateTime letzterSchultag = new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31); // Beispiel für den ersten Schultag

                        // Interessierender Zeitraum
                        DateTime von = ersterSchultag;
                        DateTime bis = letzterSchultag;

                        // Wenn es um die Statistik geht, wird der Zeitraum von-bis auf einen Tag gesetzt
                        if (zweck == Global.Zweck.Statistik)
                        {
                            von = DateTime.ParseExact(configuration["StatistikDatum"], "dd.MM.yyyy", CultureInfo.InvariantCulture);
                            bis = von;
                        }

                        // Wenn startDate leer ist, nehme ersten Schultag
                        DateTime belegungVon = ersterSchultag;
                        DateTime belegungBis = letzterSchultag;

                        if (!string.IsNullOrEmpty(student["startDate"].ToString()))
                            belegungVon = DateTime.ParseExact(student["startDate"].ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture);

                        // Wenn endDate leer ist, nehme letzten Schultag
                        if (!string.IsNullOrEmpty(student["endDate"].ToString()))
                            belegungBis = DateTime.ParseExact(student["endDate"].ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture);


                        if ((belegungVon <= bis && belegungBis >= von) || (belegungVon == von && belegungBis == bis))
                        {
                            // Wenn der Schüler bereits in der Liste ist, wird er nicht erneut hinzugefügt.
                            if (!studentsGefiltert.Any(x => x.Id == stu.Id))
                            {
                                studentsGefiltert.Add(stu);
                            }
                        }
                    }
                }
            }
        }
        return studentsGefiltert;
    }

    internal void AnzahlPrüfen(IConfiguration configuration)
    {
        var pfadDownloads = configuration["PfadDownloads"];

        // Ordner "Fotos" unterhalb von Global.PfadExportdateien
        var fotosOrdner = Path.Combine(pfadDownloads, "Fotos", this.First().Klasse);

        // Bette das in einer Schleife ein, sodass der Anwender die Fotos anpassen kann und dann erneut die Schleife durchläuft
        while (true)
        {
            // Prüfe in einer Schleife, ob in dem Ordner exakt so viele Fotos wie Schüler vorhanden sind
            var fotos = Directory.GetFiles(fotosOrdner, "*.jpg");
            if (fotos.Length != this.Count)
            {
                // Mache daraus ein rotes Panel
                AnsiConsole.MarkupLine($"[red]Es sind {fotos.Length} statt der erwarteten {this.Count} Fotos im Ordner [aqua]{fotosOrdner}[/] vorhanden.[/]");
                // Der Anwender soll in einem Panel aufgerufen werden die exakte Anzahl in den angezeigten Ordner zu legen. Anschießend soll er ENTER drücken.
                // Abbruch mit x

                AnsiConsole.MarkupLine($"Bitte legen Sie {this.Count} Fotos in den Ordner. Dann [bold green]ENTER[/] drücken. Abbruch mit [bold red]x[/].");

                var key = Console.ReadKey(true);
                if (key.Key == ConsoleKey.X)
                {
                    throw new OperationCanceledException("Sie haben abgebrochen.");
                }
            }

            // Wenn die Anzahl der Fotos jetzt stimmt, breche die Schleife ab
            if (fotos.Length == this.Count)
            {
                AnsiConsole.MarkupLine($"[green]Es liegen jetzt {fotos.Length} Fotos im Ordner [aqua]{fotosOrdner}[/]. Die Verarbeitung wird fortgesetzt ...[/]");
                break;
            }
        }
    }

    internal void KlassenordnerInZielPfadErstellen(IConfiguration configuration)
    {
        configuration = Global.Konfig("PfadFotosImSchILD-Ordner", Global.Modus.ReadSilent, configuration);
        
        var zielpfadZuFotos = configuration["PfadFotosImSchILD-Ordner"];
        
        // Falls der Zielordner nicht existiert, erstelle ihn
        if (!Directory.Exists(zielpfadZuFotos))
        {
            Directory.CreateDirectory(zielpfadZuFotos);
            Global.ZeileSchreiben("Ordner erstellt:", zielpfadZuFotos, ConsoleColor.Green, ConsoleColor.Black);
        }
    }

    internal void FotosNachSchild2Hochladen(IConfiguration configuration)
    {
        // Durchlaufe alle this, deren PfadFoto nicht nulloremptyIst
        var studentsOhneFotoInSchild = this.Where(x => !string.IsNullOrEmpty(x.ZielFotoPfad)).ToList();

        if (studentsOhneFotoInSchild.Count() == 0)
        {
            throw new InvalidOperationException("Keine neuen Fotos für SchILD2");
        }

        try
        {
            configuration = Global.Konfig("ConnectionStringSchild", Global.Modus.Update, configuration);
            DataAccess dataAccess = Global.DataAccessHerstellen(configuration);

            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Fotos hochladen ...", ctx =>
            {
                foreach (var student in studentsOhneFotoInSchild)
                {
                    student.IdSchildInt = dataAccess.GetIdSchildInt(student);
                    var fotostreamErstellt = student.Pfad2FotoStream();
                    var erfolgDelete = dataAccess.DeleteImage(student);
                    var erfolgInsert = dataAccess.InsertImage(student);
                }
            });
        }
        catch (Exception ex)
        {
            throw ex;
        }

        // Erfolgsmeldung im Panel inkl. Anzahl der Fotos
        var panel = new Panel(new Markup($"Fotos erfolgreich hochgeladen: {studentsOhneFotoInSchild.Count}"))
                    //.BorderStyle(new Style(Color.Red))
                    .Expand();
        AnsiConsole.Write(panel);
    }

    internal void FotosFürUploadNachSchild2AuflistenUndBestätigen(IConfiguration configuration)
    {
        var tableFoto = new Spectre.Console.Table();
        tableFoto.Expand();
        tableFoto.Title("Fotos, die bereit für Upload nach SchILD2 sind:");
        tableFoto.Border(TableBorder.Rounded);
        tableFoto.BorderColor(Spectre.Console.Color.SpringGreen2);
        tableFoto.Centered();
        tableFoto.AddColumn("Nr.");
        tableFoto.AddColumn("Klasse");
        tableFoto.AddColumn("Name");
        tableFoto.AddColumn("Foto");

        var studentsSortiert = this.Where(x => !string.IsNullOrEmpty(x.ZielFotoPfad)).OrderBy(x => x.Klasse).ThenBy(x => x.Nachname).ThenBy(x => x.Vorname).ToList();

        for (int i = 0; i < studentsSortiert.Count(); i++)
        {
            var student = studentsSortiert[i];

            tableFoto.AddRow((i + 1).ToString(), student.Klasse, student.Nachname + ", " + student.Vorname, Path.GetFileName(student.ZielFotoPfad));
        }
        AnsiConsole.Write(tableFoto);

        // Mit Enter soll der Anwender bestätigen, dass die Schüler*innen für den Upload ausgewählt wurden
        AnsiConsole.MarkupLine($"[]   Drücken Sie [bold green]Enter[/], um den Upload von Fotos für obige Schüler*innen zu bestätigen.[/]");
        AnsiConsole.MarkupLine($"[]   Drücken Sie [bold red]X[/], um den Upload abzubrechen.[/]");

        var eingabe = Console.ReadKey(true); // true: Zeichen wird nicht angezeigt
        if (eingabe.Key != ConsoleKey.Enter)
        {
            throw new OperationCanceledException("Sie haben den Upload abgebrochen.");
        }
    }

    internal Table GetRelationsgruppe(Relationsgruppen relationsgruppen, Table table)
    {
        foreach (var student in this.Where(x => x.Status == "2" || x.Status == "6"))
        {
            table = student.GetRelationsgruppe(relationsgruppen, table);
        }
        return table;
    }

    internal void FotosZuStudentsZuweisen(IConfiguration configuration)
    {
        configuration = Global.Konfig("PfadFotosImSchILD-Ordner", Global.Modus.ReadSilent, configuration);

        var quellpfad = configuration["PfadDownloads"];
        var zielpfad = configuration["PfadFotosImSchILD-Ordner"];

        var quellpfadZuFotos = Path.Combine(quellpfad, "Fotos", this.First().Klasse);
        var zielpfadZuFotos = zielpfad;

        var webuntisfotostrings = new List<string>();
        var geevoofotostrings = new List<string>();

        var quellFotos = Directory.GetFiles(quellpfadZuFotos, "*.jpg").OrderBy(f => f).ToArray();

        for (int i = 0; i < this.Count; i++)
        {
            this[i].QuellFotoPfad = quellFotos[i];
            this[i].ZielFotoPfad = Path.Combine(zielpfadZuFotos, this[i].MailSchulisch.Replace(configuration["MailDomain"], "").Replace("@", "") + ".jpg");
        }
    }

    internal void FotosNachZielordnerKopieren(IConfiguration configuration)
    {
        for (int i = 0; i < this.Count; i++)
        {
            if (File.Exists(this[i].QuellFotoPfad))
            {
                // Kopiere die Datei 
                File.Copy(this[i].QuellFotoPfad, this[i].ZielFotoPfad, true);
            }
        }
        if (this.Count > 0)
        {
            Global.ZeileSchreiben($"Fotos verkleinert nach {Path.GetDirectoryName(this[0].ZielFotoPfad)} kopiert", this.Count().ToString(), ConsoleColor.Green, ConsoleColor.Black);
        }
    }
    internal Students AlleOderNeueFotopfadeAnStudentsZuweisen(IConfiguration configuration)
    {
        configuration = Global.Konfig("MailDomain", Global.Modus.Read, configuration);
        configuration = Global.Konfig("PfadFotosAusSchild", Global.Modus.Read, configuration);
        var pfadFotosAusSchild = configuration["PfadFotosAusSchild"];

        // Prüfe, ob Bilder im Quellverzeichnis vorhanden sind und ob jpg enthalten sind.
        if (!Directory.Exists(pfadFotosAusSchild) || Directory.GetFiles(pfadFotosAusSchild, "*.jpg").Length == 0)
        {
            throw new Exception(
                $"Keine Fotos für den Export nach Webuntis gefunden. " +
                $"Bitte exportieren Sie zuerst die Fotos aus SchILD ([{Global.GetColor(Global.ColorPfadInDateien)}]Datenaustausch > Fotos > Fotos exportieren[/]) in den Ordner [{Global.GetColor(Global.ColorPfadInDateien)}]{pfadFotosAusSchild}[/]. " +
                $"Die Dateinamen müssen Nachname, Vorname und Geburtsdatum enthalten. " +
                $"Anschließend starten Sie diese Funktion erneut.");
        }

        // Falls Bilder vorhanden sind, wird als erstes ein Vergleich zwischen pfadFotosAusSchild und seinem neuesten
        // Klonordner versucht. Der Klon liegt im selben Ordner wie pfadFotosAusSchild. Der Name des Ordners heißt identisch,
        // hat aber das Erstelldatum als Suffix im Namen.
        var ordnername = new DirectoryInfo(pfadFotosAusSchild).Name;
        var übergeordneterOrdner = Directory.GetParent(pfadFotosAusSchild)?.FullName;
        var unterordner = Directory.GetDirectories(übergeordneterOrdner ?? "", ordnername + "*")
            .Select(d => new DirectoryInfo(d))
            .Where(di => di.FullName != pfadFotosAusSchild)
            .OrderByDescending(di => di.CreationTime)
            .FirstOrDefault()?.FullName;

        if (unterordner != null)
        {
            var table = new Table();
            table.Expand();
            table.Border(TableBorder.Rounded);
            //table.Title = new TableTitle($"Fotos");
            table.Expand();
            table.AddColumn("Schüler*in");
            table.AddColumn("alt");
            table.AddColumn("neu oder verändert & bereit für Webuntis-Zip");

            var maxAnzahlZeilen = 10;
            var alleMöglichenFotos = 0;
            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Neue und veränderte Fotos an SuS zuweisen ...", ctx =>
            {
                foreach (var student in this)
                {
                    var schuelerfotoNeu = Directory.GetFiles(pfadFotosAusSchild, $"{student.Nachname}_{student.Vorname}_{student.Geburtsdatum}*.jpg").FirstOrDefault();
                    if (schuelerfotoNeu == null) continue;
                    alleMöglichenFotos++;

                    var schuelerfotoAlt = unterordner != null ? Directory.GetFiles(unterordner, $"{student.Nachname}_{student.Vorname}_{student.Geburtsdatum}*.jpg").FirstOrDefault() : null;

                    // Wenn das Foto im Unterordner nicht vorhanden ist, dann wird das Foto für den Export nach Webuntis vorgesehen
                    if (schuelerfotoAlt == null)
                    {
                        student.ZielFotoPfad = schuelerfotoNeu;
                        if (maxAnzahlZeilen > 0)
                            table.AddRow(student.Nachname + "," + student.Vorname + ", " + student.Klasse + "(" + student.Geburtsdatum.ToString() + ")", string.IsNullOrEmpty(schuelerfotoAlt) ? "-" : Path.GetFileName(schuelerfotoAlt), Path.GetFileName(schuelerfotoNeu));
                        maxAnzahlZeilen--;
                    }
                    else
                    {
                        // Wenn das Foto im Unterordner zwar vorhanden ist, aber
                        // Eigenschaften abweichen, dann wird das Foto für den Export nach Webuntis vorgesehen
                        var infoNeu = new FileInfo(schuelerfotoNeu);
                        var infoAlt = new FileInfo(schuelerfotoAlt);
                        if (infoNeu.Length != infoAlt.Length)
                        {
                            student.ZielFotoPfad = schuelerfotoNeu;
                            if (maxAnzahlZeilen > 0)
                                table.AddRow(student.Nachname + "," + student.Vorname + ", " + student.Klasse + "(" + student.Geburtsdatum.ToString() + ")", string.IsNullOrEmpty(schuelerfotoAlt) ? "-" : Path.GetFileName(schuelerfotoAlt), Path.GetFileName(schuelerfotoNeu));
                            maxAnzahlZeilen--;
                        }
                    }
                    if (maxAnzahlZeilen == 0)
                    {
                        table.AddRow("...", "...", "...");
                        maxAnzahlZeilen--;
                    }
                }
            });

            var anzahl = this.Count(x => !string.IsNullOrEmpty(x.ZielFotoPfad));
            //Global.ZeileSchreiben("Neue Fotos und veränderte Fotos an SuS zugewiesen:", anzahl.ToString(), ConsoleColor.Green, ConsoleColor.Black);

            table.AddRow($"Summe: [{Global.GetColor(Global.ColorZahlen)}]{anzahl}[/]", "");
            AnsiConsole.Write(table);

            Global.ZeileSchreiben("Alle Fotos im Ordner " + pfadFotosAusSchild, alleMöglichenFotos.ToString(), ConsoleColor.Green, ConsoleColor.Black);
            Global.ZeileSchreiben("Neue oder veränderte Fotos im Ordner", anzahl.ToString(), ConsoleColor.Green, ConsoleColor.Black);
        }

        // Frage den Anwender, ob er die gefundenen Differenzen zwischen den Ordnern nach Webuntis exportieren möchte,
        // oder ob er alle Fotos exportieren möchte.   

        configuration = Global.Konfig("NurNeueFotosExportieren", Global.Modus.Update, configuration);

        if (configuration["NurNeueFotosExportieren"].ToString().ToLower() == "ja")
        {           
            var studentsMitNeuenFotos = new Students();
            studentsMitNeuenFotos.AddRange(this.Where(x => !string.IsNullOrEmpty(x.ZielFotoPfad))); 
            return studentsMitNeuenFotos;
        }   

        // Wenn der Anwender alle Fotos exportieren möchte, dann wird jedem Schüler das Foto aus dem Ordner pfadFotosAusSchild zugewiesen.
        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Fotos an SuS zuweisen ...", ctx =>
        {
            foreach (var student in this)
            {
                var fotoDiesesSchuelersInOrdner = Directory.GetFiles(pfadFotosAusSchild, $"{student.Nachname}_{student.Vorname}_{student.Geburtsdatum}*.jpg").FirstOrDefault();
                if (fotoDiesesSchuelersInOrdner == null)
                    continue;

                student.ZielFotoPfad = fotoDiesesSchuelersInOrdner;
            }
        });

        var anzahlAlle = this.Count(x => !string.IsNullOrEmpty(x.ZielFotoPfad));
        Global.ZeileSchreiben($"Fotos an SuS zugewiesen", anzahlAlle.ToString(), ConsoleColor.Green, ConsoleColor.Black);

        return this;
    }

    internal void FotosRotieren(IConfiguration configuration)
    {
        configuration = Global.Konfig("RotateFotos", Global.Modus.Update, configuration);

        if (configuration["RotateFotos"] == "0")
        {
            return;
        }
        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Fotos drehen ...", ctx =>
        {
            foreach (var student in this)
            {
                if (File.Exists(student.ZielFotoPfad))
                {
                    string inputPath = student.ZielFotoPfad;
                    string tempPath = Path.Combine(Path.GetTempPath(), Path.GetFileName(student.ZielFotoPfad));

                    using (var img = Image.Load(inputPath))
                    {
                        // Entfernt alle Metadaten (inkl. EXIF)
                        //img.Metadata.ExifProfile = null;

                        // Optional: Bildbearbeitung
                        img.Mutate(x => x.Rotate(Convert.ToInt32(configuration["RotateFotos"])));
                        img.Save(student.ZielFotoPfad);
                    }
                }
            }
        });
    }

    public void GetMassnahmen(IConfiguration configuration, List<string> maßnahmenString, Dateien Quelldateien)
    {
        var sMitMassnahmen = new Students();

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("SuS mit Maßnahmen ermitteln ...", ctx =>
        {
            var schuelerZusatzdaten = Quelldateien.GetMatchingList(configuration, "schuelerzusatzdaten", this, null);

            foreach (Student student in this)
            {
                if (student.Nachname.StartsWith("Kov") && student.Vorname.StartsWith("Lau"))
                {
                    string aa = "";
                }
                student.GetMaßnahmen(configuration, maßnahmenString);

                if (student.Massnahmen.Count != 0)
                    sMitMassnahmen.Add(student);
            }
        });

        Global.ZeileSchreiben($"SuS mit Maßnahmen:", $"{sMitMassnahmen.Count}");
    }
}