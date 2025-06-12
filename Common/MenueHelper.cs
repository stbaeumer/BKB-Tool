using System.Text;
using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
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
public static class MenueHelper
{
    public static Menue Einlesen(Dateien quelldateien, IConfiguration configuration)
    {
        var raums = new Raums();
        var lehrers = new Lehrers();
        var klassen = new Klassen();
        var anrechnungen = new Anrechnungen();

        var pfadDownloads = configuration["PfadDownloads"];
        var pfadSchilddatenaustausch = configuration["PfadSchilddatenaustausch"];
        var netmanMailReceiver = configuration["NetmanMailReceiver"];
        var pdfKennwort = configuration["PdfKennwort"];
        var betreff = configuration["Betreff"];
        var body = configuration["Body"];
        var betreffMassenmail = configuration["BetreffMassenmail"];
        var inputFolder = configuration["InputFolder"];
        var outputFolder = configuration["OutputFolder"];        

        try
        {
            var students = new Students(configuration, quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "schildschuelerexport,txt,optional"], true));
            quelldateien.Meldung.Add(students.GetArtUndZahlen());

            lehrers = new Lehrers(configuration, quelldateien.Notwendige(configuration, ["lehrkraefte,dat"], true));
            
            if (students.Count == 0 || lehrers.Count == 0)
            {
                return new Menue(quelldateien, klassen, lehrers, students, []);
            }

            #pragma warning disable CS8601 // Mögliche Nullverweiszuweisung
            //Console.WriteLine("");
            //AnsiConsole.Write(new Rule("").RuleStyle("springgreen2").Centered());

            return new Menue(
                quelldateien,
                klassen,
                lehrers,
                students,
                [
                    new Menüeintrag(
                        "Webuntis: Schüler*innen-Importdatei für Webuntis erstellen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [bold aqua] {Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd") + "-ImportNachWebuntis.csv")}[/] erstellt.",                            
                            "[pink3]Hinweis:[/] Das Zeugnisdatum des letzten Zeugnisses in einer Klasse wird zum Webuntis-Austrittsdatum bei Schüler*innen, deren Status weder aktiv noch extern ist."
                        ],
                        m =>
                        {
                            var zeitstempel = DateTime.Now.ToString("yyyyMMdd-HHmm");
                            m.Zieldatei = m.WebuntisOderNetmanOderLitteraCsv(configuration, Path.Combine(pfadDownloads ?? "", zeitstempel +  @"-ImportNachWebuntis.csv"));
                            m.Zieldatei?.Erstellen(";", '\'', new UTF8Encoding(false), false,
                            [
                                "1. In Webuntis als Webuntis-Admin:  [bold yellow]Stammdaten > Schüler*innen > Import[/]",
                                "2. Datei auswählen, UTF8",
                                "3. Profil: Schuelerimport, dann Vorschau",
                                "Mehr zum Profil Schuelerimport: [lightskyblue3_1][link=https://github.com/stbaeumer/BKB-Tool/wiki]https://github.com/stbaeumer/BKB-Tool/wiki[/][/]"
                            ]
                            );
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Webuntis-Fotos : Zipdatei mit Fotos für Webuntis erstellen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [aqua] {Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd") + "-ImportNachWebuntis.zip")}[/] erstellt.",
                            "[pink3]Hinweis:[/] Schüler*innen, deren Foto hochgeladen wurden, werden in der Datei [aqua]fotos.txt[/] gespeichert, um ein erneutes Hochladen zu vermeiden."
                        ],
                        m =>
                        {
                            var zeitstempel = DateTime.Now.ToString("yyyyMMdd-HHmm");                            
                            m.IStudents = m.Students.OhneWebuntisFoto(configuration, Path.Combine(Directory.GetCurrentDirectory(), "fotos.txt"));
                            m.IStudents.FotosFürWebuntisZippen(configuration, Path.Combine(pfadDownloads ?? "", zeitstempel +  @"-ImportNachWebuntis.zip"), Path.Combine(Directory.GetCurrentDirectory(), "fotos.txt"),
                            [
                                "1. [bold yellow]Stammdaten > Schüler*innen > Bildimport[/]",
                                "2. Identifizierung Fremdschlüssel",
                                "3. [springGreen2]Datei auswählen[/]"
                            ]);
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Littera: Schüler*innen-Importdatei für Littera erstellen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            "Es werden jetzt die Dateien [bold aqua]" + Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-") + "****" +  @"-ImportNachWebuntis.csv") + "[/] und [bold aqua]***-ImportNachWebuntis.zip[/] erstellt. Der Webuntis-Admin muss im Anschluss die Dateien wie folgt importieren:",
                            "1. [bold yellow]Stammdaten > Schüler*innen > Import[/]",
                            "2. Datei auswählen, UTF8",
                            "3. Profil: Schuelerimport, dann Vorschau",
                            "[pink3]Hinweis:[/] Das Zeugnisdatum des letzten Zeugnisses in einer Klasse wird zum Webuntis-Austrittsdatum bei Schüler*innen, deren Status weder aktiv noch extern ist.",
                            "Es wird jetzt die Datei [bold aqua]" + Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-") + "****" +  @"-ImportNachNetman.csv") + "[/] erstellt:",
                            "1. Die Datei Student_***.csv muss als admin aus Webuntis nach [bold aqua]" + pfadDownloads + "[/] heruntergeladen werden: [bold yellow]Stammdaten > Schüler > Export als csv[/].",
                            "2. Nach dem Einlesen werden alle Schülerinnen und Schüler angezeigt, die in Untis zu löschen sind. Es wird bei diesen Schülerinnen und Schülern ein Austrittsdatum gesetzt.",
                            "3. Die Importdatei für Netman und Littera werden geschrieben.",
                            "4. Die Netman-Datei wird gezippt verschickt.",
                            "Es wird jetzt die Datei [bold aqua]" + Path.Combine(@"\\fs01\Littera\Atlantis Import Daten" ?? "", DateTime.Now.ToString("yyyyMMdd-") + "****" +  @"-ImportNachLittera.csv") + "[/] erstellt.",
                        ],
                        m =>
                        {
                            var zeitstempel = DateTime.Now.ToString("yyyyMMdd-HHmm");
                            m.Zieldatei = m.WebuntisOderNetmanOderLitteraCsv(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.AddHours(1).ToString("yyyyMMdd-HHmm") + @"-ImportNachLittera.xml"));
                            m.Zieldatei?.Erstellen(";", '\"', new UTF8Encoding(false), false);
                            m.Zieldatei?.Verschieben(@"\\fs01\Littera\Atlantis Import Daten");
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),                    
                    new Menüeintrag(
                        "Netman: Schüler*innen-Importdatei für Netman erstellen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [aqua] {Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd") + "-ImportNachNetman.csv")}[/] erstellt.",
                            "[pink3]Hinweis:[/] Schüler*innen, die bereits abgegangen sind oder einen Abschluss erworben haben, werden erst sechs Wochen später ausgebucht, um den Zugriff auf Teams nicht direkt zu verlieren."
                        ],
                        m =>
                        {
                            var zeitstempel = DateTime.Now.ToString("yyyyMMdd-HHmm");
                            m.Zieldatei = m.WebuntisOderNetmanOderLitteraCsv(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.AddHours(1).ToString("yyyyMMdd-HHmm") + @"-ImportNachNetman.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', new UTF8Encoding(false), false);
                            m.Zieldatei?.Zippen(m.Zieldatei?.GetAbsoluterPfad(), configuration);
                            m.Zieldatei?.Mailen(Path.GetFileName(m.Zieldatei.AbsoluterPfad) ?? "", "Verwaltung", Path.GetFileName(m.Zieldatei.AbsoluterPfad) ?? "", configuration);
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),                   
                    new Menüeintrag(
                        "Statistik: Unterrichtsverteilung für UVD und Anrechnungen nach SchILD importieren",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["exportlessons,csv", "studentgroupstudents,csv", "schuelerlernabschnitt,dat", "schuelerleistungsdaten,dat", "schuelerbasis,dat", "lehrkraefte,dat", "lehrkraeftesonderzeiten,dat", "schuelerbasisdaten,dat", "GPU020,txt"]),
                        students,
                        klassen,
                        [
                            "Es werden jetzt die Dateien [bold springGreen2]" + Path.Combine(pfadSchilddatenaustausch ?? "", "Lernabschnitte.dat") + "[/], [bold springGreen2]" + Path.Combine(pfadSchilddatenaustausch ?? "", "Leistungsdaten.dat") + "[/] und [bold springGreen2]" + Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat") + "[/] erstellt.",
                            "Unterrichte in Blockklassen werden mit dem Faktor 3 multipliziert.",
                            "Da SchILD bei Unterrichten nur Ganzzahlen entgegennehmen kann werden die Werte aus Webuntis gerundet.",
                            "Die Datei LehrkraefteSonderzeiten.dat muss zuerst aus SchILD exportiert werden. Die exportierte Datei wird dann für den ReImport aufbereitet.",
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            //m.IStudents = m.Students;
                            //m.IKlassen = m.Klassen;
                            m.Zieldatei = m.Lernabschnittsdaten(configuration, Global.Art.Statistik, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLernabschnittsdaten.dat"));
                            //m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            m.Zieldatei = m.Leistungsdaten(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"), Global.Art.Statistik);
                            //m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach"], ["Jahrgang"]);
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            m.Zieldatei = m.LehrkraefteSonderzeiten(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat"));
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            m.Zieldatei = m.Lehrkraefte(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Lehrkraefte.dat"));
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            //m.Zieldatei = m.Faecher(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat"));
                            //m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            //m.Zieldatei = m.Kurse(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat"));
                            //m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["KursBez"], ["Klasse", "Schulnr", "WochenstdPUNKTLEERZEICHENKL"]);
                            //m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);  

                            m.Zieldatei.Ordner = pfadSchilddatenaustausch ?? "";
                        },
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Zeugnisse #1: Unterrichte, Noten & Fehlzeiten nach SchILD importieren",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "absenceperstudent,csv", "schuelerlernabschnitt,dat", "schuelerleistungsdaten,dat", "schuelerbasis,dat", "exportlessons,csv", "studentgroupstudents,csv", "marksperlesson,csv"]),
                        students,
                        klassen,
                        [
                            $"Die Unterrichte (mit Noten) werden in der [aqua]{Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat")}[/] vorbereitet. Hinzu kommen [aqua]{Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat")}[/] und [aqua]{Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat")}[/] und [aqua]{Path.Combine(pfadSchilddatenaustausch ?? "", "Lernabschnittsdaten.dat")}[/].",
                            $"Es empfiehlt sich die Lernabschnitte zuerst in SchILD anzulegen und zu exportieren. [bold springGreen2]BKB-Tool[/] ergänzt dann die Fehlzeiten passend.",
                            "Falls mehrere Kollegen dasselbe Fach zeitgleich unterrichten, dann muss ein Zähler an das Fach angehangen werden. Bsp: Zwei LuL unterrichten Mathe. Dann M und M1.",
                            "Damit M1 in den Leistungsdaten erscheint, aber nicht auf dem Zeugnis gedruckt wird, muss die Eigenschaft 'Nicht auf Zeugnis drucken' in SchILD gesetzt werden."
                            ,
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);

                            m.Zieldatei = m.Lernabschnittsdaten(configuration, Global.Art.Zeugnis, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLernabschnittsdaten.dat"));
                            //m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            m.Zieldatei = m.Leistungsdaten(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"), Global.Art.Zeugnis);
                            //m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach"], ["Jahrgang"]);
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            /*
                            m.Zieldatei = m.Teilleistungen(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerTeilleistungen.dat"));
                            m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                            m.Zieldatei.Erstellen("|", '\0', new UTF8Encoding(true), false);    

                            m.Zieldatei = m.Faecher(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat"));
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            m.Zieldatei = m.Kurse(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat"));
                            m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["KursBez"], ["Klasse", "Schulnr", "WochenstdPUNKTLEERZEICHENKL"]);
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("  Schritt #1: Import nach SchILD durchführen.");
                            Console.WriteLine("  Schritt #2: Notenkontrollliste nach Wiki hochladen.");
                            Console.WriteLine("  Schritt #3: Lehrkräfte benachrichtigen.");                            */
                        },
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Zeugnisse #2: Lehrkräfte an die Eintragung erinnern",
                        anrechnungen,
                        quelldateien,
                        students,
                        klassen,
                        [
                            "Zuerst müssen die Leistungsdaten nach SchILD importiert werden.",
                            "Aus dem anschließenden Datenexport aus SchILD werden hier alle fehlenden Zeugnisnoten aufgelistet.",
                            "Alle betreffenden Mail-Adressen der Lehrkräfte können mit Copy & Paste nach Teams kopiert werden."
                        ],
                        m => m.LuLAnEintragungDerZeugnisnotenErinnern(configuration, lehrers),
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Klassen: Neu anlegen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["klassen,dat", "GPU003,txt"]),
                        students,
                        klassen,
                        [
                            "Es wird angenommen, dass die Klassen des kommenden Schuljahres zuerst in Untis angelegt werden.",
                            "Alle neuen Klassen in Untis können vor dem Schuljahreswechsel nach SchILD übergeben werden.",
                            "Dazu die GPU003.TXT aus Untis exportieren, um neue Klassenbezeichnungen zu identifizieren.",
                            "Eigenschaften der neuen Klassen werden aus SchILD-Vorjahresklassen entnommen.",
                            "Bei vorhandenen Klassen werden abweichende Eigenschaften (z.B. Klassenleitung) angepasst.",
                        ],
                        m =>
                        {
                            m.Zieldatei = m.KlassenErstellen(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", @"Klassen.dat"));
                            if (!m.Zieldatei.Any()) return;
                            m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["InternBez"], ["SonstigeBez", "Folgeklasse"]);
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Mahnungen: Gem. §50(4) SchulG erstellen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["marksperlesson,csv", "schuelerleistungsdaten,dat", "exportlessons,csv", "studentgroupstudents,csv", "schuelerleistungsdaten,dat", "schuelerbasisdaten,dat"]),
                        students,
                        klassen,
                        [
                            "Die Datei SchuelerLeistungsdaten wird erstellt.",
                            "Die Datei SchuelerLeistungsdaten kann dann nach SchILD importiert werden.",
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            m.Zieldatei = m.Leistungsdaten(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"), Global.Art.Mahnung);
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Teams-Chat: Teams-Chat mit gewünschter Gruppe von Lehrkräften beginnen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["exportlessons,csv"]),
                        students,
                        klassen,
                        [
                            "Lehrkräfte können über Teams angeschrieben werden."
                        ],
                        m =>
                        {
                            var datei = m.GetGruppen(configuration, "", anrechnungen, lehrers ?? []);

                            var table = new Table();
                            table.AddColumn("Nr.");
                            table.AddColumn("Gruppe");

                            var zulässigeAuswahlOptionen = "";

                            for (int i = 1; i < datei.Count; i++)
                            {
                                string page = datei[i].Page;
                                table.AddRow(i.ToString(), page);
                                zulässigeAuswahlOptionen += i + ", ";
                            }

                            table.AddRow((datei.Count).ToString(), "(Eine einzelne) Klasse(n) wählen");
                            zulässigeAuswahlOptionen += datei.Count;

                            AnsiConsole.Write(table);

                            configuration = Global.Konfig("TeamsChatAuswahl", Global.Modus.Update, configuration, "Bitte eine Zahl auswählen:", "Bitte eine Zahl auswählen:", Global.Datentyp.Int, (datei.Count).ToString(), students, zulässigeAuswahlOptionen);

                            var nummer = int.Parse(configuration["TeamsChatAuswahl"]);

                            if (nummer > 0 && nummer < datei.Count)
                            {
                                m.ChatErzeugen(configuration, datei[nummer - 1].MitgliederMail);
                            }

                            if (nummer == datei.Count)
                            {
                                m.FilterInteressierendeStudentsUndKlassen(configuration);
                                datei = m.GetLehrerDerKlassen(configuration, lehrers ?? []);
                                m.ChatErzeugen(configuration, datei[0].MitgliederMail);
                            }
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Kursbelegung: Vorbereiten",
                        anrechnungen,
                        quelldateien,
                        students,
                        klassen,
                        [
                            " 1. Alle Gymklassen der Jahrgangsstufen 12 und 13 aus SchILD exportieren.",
                            " 2. Alle Dateien in die Exceldatei namens Kursbelegung.xlsx importieren. Siehe LiesMich.",
                            " 3. Nachdem die Datei Kursbelegung.xlsx gefüllt wurde, die Datei nach SchILD importieren."
                        ],
                        m =>
                        {
                            //dateien.Kursbelegung(dateien.Benötigte([]));
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Fotos #1: Anweisungen zum Fotografieren der Schüler*innen zeigen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat"]),
                        students,
                        klassen,
                        [
                            "Erstellen Sie jetzt Fotos aller Schüler (z.B. mit dem Handy). Dabei ist die [bold red]Reihenfolge[/] und die [bold red]Anzahl[/] der Schüler laut SchILD exakt einzuhalten. ",
                            "   [bold aqua]Hinweis #1:[/] Wenn ein Schüler fehlt, dann die weiße Wand fotografieren, damit Reihenfolge und Anzahl stimmen.",
                            "   [bold aqua]Hinweis #2:[/] Wenn ein Foto nicht gelungen ist, dann löschen und neu erstellen.",
                            "   [bold aqua]Hinweis #3:[/] Wenn mehr als eine Klasse ausgewählt wird, wird nur die erste Klasse berücksichtigt",
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            m.IStudents.KlassenordnerErstellen(configuration);
                            m.IStudents.KlassenListenAnzeigen(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Fotos #2: Erstellte Schüler*innenfotos hochladen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat"]),
                        students,
                        klassen,
                        [
                            "Vorarbeiten: Fotos aller Schüler wurden bereits erstellt und liegen nun in der richtigen [bold aqua]Reihenfolge[/] und [bold aqua]Anzahl[/] im Ordner " + Path.Combine(configuration["PfadDownloads"], "Fotos") +  ".",
                            "Beachten Sie die Hinweise.",
                            "Starten Sie am Ende den Importprozess."
                        ],
                        m =>
                        {
                            m.IKlassen = m.Students.KlassenAuswählen(configuration);
                            m.Students.FotosVerarbeiten(configuration, m.IKlassen);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Atlantis-Fotos: Fotos der Schüler*innen aus Atlantis in die SchILD2-Datenbank (und in die Schild-Dokumentenverwaltung) hochladen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration,["schuelerbasisdaten,dat"]),
                        students,
                        klassen,
                        [
                            "Ablauf in 5 Schritten:",
                            "#1 Vorbereitung: Atlantis-Fotos nach " + Path.Combine(pfadDownloads, "Fotos") + " kopieren.",
                            "#2 Klasse auswählen.",
                            "#3 Die Atlantis-Schülerfotos werden aus " + Path.Combine(pfadDownloads, "Fotos") + " herausgesucht.",
                            "#4 Die Atlantis-Schülerfotos werden in die Dokumentenverwaltung kopiert.",
                            "#5 Geben Sie die Access-Zugangsdaten ein und bestätigen Sie Import."
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            m.IStudents.GetPfadAtlantisFotos(configuration);
                            m.IStudents.GetPfadDokumentenverwaltung(configuration);
                            m.IStudents.ErstellenPfadDokumentenverwaltung(configuration);
                            m.IStudents.BilderNachPfadDokumentenverwaltungKopieren(configuration);
                            m.IStudents.Pfad2FotoStream();
                            //m.IStudents.FotosNachSchild2Schreiben(m.Klassen, configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Teilleistungen: SchuelerTeilleistungen nach Schild importieren",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "exportlessons,csv", "marksperlesson,csv"]),
                        students,
                        klassen,
                        [
                            "Es wird jetzt die Datei [bold springGreen2]" + Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-") + @"SchuelerTeilleistungen.dat") + "[/] erstellt.",
                            "Damit der Import nach SchILD rebuungslos funktioniert muss Folgendes gewährleistet sein:",
                            "Die Teilleistungsarten in SchILD unter Schulverwaltung > Teilleistungsarten müssen gleichlautend mit dem Langnamen in Webuntis (Stammdaten > Prüfungsarten) angelegt sein.",
                            "Es empfiehlt sich, dass die Lernabschnitssdaten und Leistungsdaten zuerst in SchILD importiert bzw. angelegt werden."
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            m.Zieldatei = m.Teilleistungen(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerTeilleistungen.dat"));
                            //m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                            m.Zieldatei.Erstellen("|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Schnellmeldung: Relationsgruppen im September aufbereiten",
                        anrechnungen,
                        quelldateien,
                        students,
                        klassen,
                        [
                            "Dokumentation siehe Schips.webuntis2schildGui.nrw.de",
                            "Realtionen gemäß §93 SchulG"
                        ],
                        _ => { new Relationsgruppen(klassen, students); },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000                        
                    ),
                    new Menüeintrag(
                        $"Altersermäßigung: berechnen für {int.Parse(Global.AktSj[0])}/{int.Parse(Global.AktSj[0]) + 1} und {int.Parse(Global.AktSj[0]) + 1}/{int.Parse(Global.AktSj[0]) + 2}",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["lehrkraefte,dat", "lehrkraeftesonderzeiten,dat,optional", "GPU020,txt,optional", "GPU004,txt,optional"]),
                        students,
                        klassen,
                        [
                            $"Die Altersermäßigung wird aus der Datei [aqua]{Path.Combine(configuration["PfadDownloads"] ?? "", "Lehrkraefte.dat")}[/] berechnet und mit der [aqua]{Path.Combine(configuration["PfadDownloads"] ?? "", "LehrkraefteSonderzeiten.dat")}[/] und optional [aqua]{Path.Combine(configuration["PfadDownloads"] ?? "", "GPU020.TXT")}[/] abgeglichen.",
                            "Alle Lehrkräfte (angestellt, verbeamtet und auch Werkstattlehrer) erhalten die erste Altersermäßigung ab dem Schuljahr, das auf die Vollendung des 55. Lebensjahres folgt. Pech hat also, wer z.B. am 1. August 55 Jahre alt wird. Dann gibt es die Altersermäßigung erst ab dem kommenden Schuljahr.",
                            "Ab dem 55. Lebensjahr erhalten Vollzeitbeschäftigte 1 Stunde Altersermäßigung, Teilzeitbeschäftigte (mindestens 50%) erhalten 0,5 Stunden.",
                            "Lehrkräfte, die ihre Stundenzahl nur um 1 Stunde verringert haben, erhalten ebenfalls die komplette Altersermäßigung.",
                            "Ab dem 60. Lebensjahr beträgt die Altersermäßigung 3 Stunden für Vollzeitbeschäftigte, 2 Stunden für Teilzeitbeschäftigte mit mindestens 75% und 1,5 Stunden für Teilzeitbeschäftigte mit mindestens 50%."
                        ],
                        m =>
                        {
                            m.Zieldatei = m.LehrkraefteSonderzeiten(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat"), "200");
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            m.Zieldatei = m.Lehrkraefte(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Lehrkraefte.dat"));
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Lernabschnittsdaten: Lernabschnitts- & Leistungsdaten alter Abschnitte",
                        anrechnungen,
                        quelldateien,
                        students,
                        klassen,
                        [
                            "Die Lernabschnittsdaten (ohne Fehlzeiten und ohne Zeugnisdatum) und die Leistungsdaten alter Abschnitte werden für SchILD bereitgestellt."
                        ],
                        m =>
                        {
                            m.Zieldatei = m.LernabschnittsdatenAlt(@"DatenaustauschSchild/SchuelerLernabschnittsdaten.dat", configuration);
                            m.Zieldatei.Erstellen("|", '\0', new UTF8Encoding(true), false);

                            m.Zieldatei = m.LeistungsdatenAlt(configuration, @"DatenaustauschSchild/SchuelerLeistungsdaten.dat");
                            m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Klassenbucheinträge: Säumige Lehrer*innen erinnern",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["lehrkraefte,dat", "openperiod,pdf"]),
                        students,
                        klassen,
                        [
                            "Die 10% der KuK mit den meisten offenen Klassenbucheinträgen werden angemahnt.",
                            "Unter 10 fehlenden wird nicht gemahnt. Ab 20 oder mehr Stunden wird die Schulleitung in CC informiert.",
                            "Die Anzahl der offenen Klassenbucheinträge wird aus der Datei 'OpenPeriods' ausgelesen.",
                            "Die KuK werden zuerst angezeigt. Vor dem Mailversand wird nochmal explizit gefragt."
                        ],
                        m =>
                        {
                            lehrers = new Lehrers(configuration, m.Quelldateien);
                            lehrers.OffeneKlassenbuchEinträgeMahnen(m.Quelldateien, configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        $"PDF-Dateien #1: Von {configuration["PfadDownloads"]}\\*.pdf verschlüsselte Kopien erstellen",
                        anrechnungen,
                        new Dateien(),
                        students,
                        klassen,
                        ["Von PDF-Dateien in " + configuration["PfadDownloads"] + " wird eine verschlüsselte Kopie erstellt.",
                        "Kopien bekommen die Dateiendung '-kennwort.pdf'"],
                        _ =>
                        {
                            var pdfDateien = new PdfDateien();
                            pdfDateien.KennwortSetzen(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "PDF-Dateien #2: PDF-Seiten an darauf enthaltene E-Mail-Adressen mailen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["lehrkraefte,dat"]),
                        students,
                        klassen,
                        [
                            "Die zuletzt bearbeitete PDF-Datei wird eingelesen.",
                            "Jede Seite der Datei wird nach E-Mail-Adressen durchsucht.",
                            "Die betreffenden Seiten werden an die E-Mail-Adressen gemailt.",
                            "Optional wird verschlüsselt."
                        ],
                        m =>
                        {
                            var pdfDatei = Directory.GetFiles(pfadDownloads, "*.pdf").OrderByDescending(File.GetLastWriteTime).FirstOrDefault();
                            Global.ZeileSchreiben("Die neueste PDF-Datei wird versendet:", pdfDatei, ConsoleColor.White, ConsoleColor.Black);
                            Global.Konfig("PdfKennwort", Global.Modus.Update, configuration, "Kennwort für das verschlüsseln von PDFs angeben");
                            Global.Konfig("Betreff", Global.Modus.Update, configuration, "Betreff angeben. Der Betreff wird um das Lehrerkürzel ergänzt)");
                            Global.Konfig("Body", Global.Modus.Update, configuration, @"Body angeben (\n wird durch Zeilenumbruch ersetzt; [Lehrer] wird durch Lehrername ersetzt)");
                            foreach (PdfSeite seite in (new PdfDatei(configuration, pdfDatei, new Lehrers(configuration, m.Quelldateien))).Seiten)
                            {
                                seite?.GetMailReceiver(lehrers);
                                seite?.PdfDocumentCreate(pdfDatei);
                                seite?.PdfDocumentEncrypt(pdfKennwort);
                                seite?.Mailen(betreff, body, configuration);
                            }
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "PDF-Zeugnisse: Von Atlantis in die SchILD-Dokumentenverwaltung kopieren",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, []),
                        students,
                        klassen,
                        [
                            $"PDF-Zeugnisse und andere PDF-Dateien werden in die Schüler*innen-Ordner der SchILD-Dokumentenverwaltung einsortiert.",
                            $"[yellow]Vorbereitung #1[/]: Zu kopierende PDF-Dateien nach {Path.Combine(configuration["PfadDownloads"], "PDF-Input")} kopieren.",
                            $"[yellow]Vorbereitung #2[/]: Eine UTF8-CSV-Datei mit Spalten: Nachname, Vorname, Geburtsdatum und Klasse aus Atlantis exportieren und in {Path.Combine(pfadDownloads, "PDF-Input")} ablegen.",
                            $"[aqua]Durchführung #1[/]: (Einzelne) Klasse(n) oder 'alle' auswählen.",
                            $"[aqua]Durchführung #2[/]: Geben Sie die Schlüsselwörter an, um die interessierenden PDF-Dateien einzugrenzen.",
                        ],
                        m =>
                        {

                            configuration = Global.Konfig("PfadDownloads", Global.Modus.Update, configuration, "Pfad zum eigenen Download-Ordner angeben","",Global.Datentyp.Pfad);
                            configuration = Global.Konfig("PfadDokumentenverwaltung", Global.Modus.Update, configuration, "Pfad zur SchILD-Dokumentenverwaltung angeben (z.B. \\fs01\\Schild\\Dokumentenverwaltung)","",Global.Datentyp.Pfad);
                            configuration = Global.Konfig("Schlüsselwörter", Global.Modus.Update, configuration, "Geben Sie kommagetrennt interessierende Schlüsselwörter an (z.B. Abgangszeugnis, Abschlusszeugnis, Jahreszeugnis)","",Global.Datentyp.String);

                            m.IStudents.GetStudentsVonAtlantisCsv(configuration);
                            m.IStudents.PdfDateienVerarbeiten(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Outlook: CSV-Terminexporte für Wiki aufbereiten",
                        anrechnungen,
                        quelldateien.Notwendige(configuration,["termine_fhr,csv", "termine_verwaltung,csv", "termine_berufliches_gymnasium,csv", "termine_kollegium,csv"]),
                        students,
                        klassen,
                        [
                            "Die Kalender müssen mit Copy&Paste aus Outlook in die CSV-Dateien im Download-Ordner kopiert werden.",
                            "Falls der Inhalt im Body (Spalte Nachricht) mehrzeilig ist, wird nur die erste Zeile berücksichtigt.",
                            "Es werden nur Termine berücksichtigt, die mindestens eine Kategorie haben. Kategorien werden zu Links in Wiki.",
                            "Termine aus vergangenen Schuljahren werden nicht mit übertragen.",
                            "Die Kalender im Wiki zuerst leeren. Anschließend die neuen CSV als Global importieren."
                        ],
                        m =>
                        {
                            foreach (var kalender in new List<string>(){"termine_berufliches_gymnasium", "termine_kollegium", "termine_verwaltung", "termine_fhr" })
                            {
                                m.Zieldatei = m.Kalender2Wiki(configuration, kalender, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWiki-" + kalender));
                                m.Zieldatei.Erstellen(",", '\"', new UTF8Encoding(false), true);
                            }
                        },
                        Global.Rubrik.Wiki,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Sprechtag: Lehrerübersichtsseite im Wiki veröffentlichen",
                        anrechnungen,
                        quelldateien,
                        students,
                        klassen,
                        [
                            "Die Wiki-Datei sprechtag.txt wird angepasst. Die Wunschräume werden in den Untis-Stammdaten beim Lehrer eingetragen. Dazu die Fenstergruppe Sprechtag in Untis öffnen. Bei Abwesenheiten die Räume für kommendes Jahr stehen lassen, wenn im Betreff 'außer Haus' steht, dann wird der Raum nicht angezeigt. Fußnoten werden als Text2 in den Untis-Stammdaten eingetragen. Beispiel für eine Fußnote: 'außer Haus; bitte Termin vereinbaren;'",
                            "Lehrkräfte ohne Raum werden in der Liste ignoriert. Lehrkräfte ohne eigenen Unterricht bleiben unberücksichtigt"
                        ],
                        m =>
                        {
                            m.Sprechtag(lehrers, raums, configuration,
                                "Zum jährlichen Sprechtag laden wir sehr herzlich am Mittwoch nach der Zeugnisausgabe in der Zeit von 13:30 bis 17:30 Uhr ein. Der Unterricht endet nach der 5. Stunde um 12:00 Uhr.");
                        },
                        Global.Rubrik.Wiki,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Wiki: Diverse SQLite-Dateien (Organigramm, Praktikum etc.) erstellen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, [
                            "schuelerzusatzdaten,dat", "absenceperstudent,csv", "exportlesson,csv"
                        ]),
                        students,
                        klassen,
                        [
                            "Das Organigramm wird aus Untisanrechnungen gebildet. Beispiele: {...} > KATEGORIE; [...] > HINWEIS, Text ohne Klammern wird zur ROLLE; A14, A15, A16 ohne Klammern > AMT; Untis-Beschreibung > AUFGABE. Im Organigramm wird nach Kategorie, Aufgabe oder Beschreibung gruppiert.",
                            "Untisanrechnungen: 1.Struct Schema Editor > Untisanrechnungen > Löschen/Leeren > 'untisanrechnungen' eingeben, dann Leeren",
                            "Untisanrechnungen: 2.Struct Schema Editor > Untisanrechnungen > Importieren/Exportieren > Importieren von Rohdaten > Global > Durchsuchen"
                        ],
                        m =>
                        {
                            m.Zieldatei = m.GetGruppen(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-gruppen.csv"), anrechnungen, lehrers);
                            m.Zieldatei.Erstellen(",", '\"', new UTF8Encoding(false), true);
                            m.Zieldatei = anrechnungen.Anlegen(Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-untisanrechnungen.csv") ,[500, 510, 530, 590, 900], [500, 510, 530, 590], ["PLA", "BM"]);
                            m.Zieldatei.Erstellen(",", '\"', new UTF8Encoding(false), true);

                            m.Zieldatei = m.GetLehrer(configuration, Path.Combine(Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-lul-utf8OhneBom-einmalig-vor-SJ-Beginn.csv")));
                            m.Zieldatei.Erstellen(",", '\'', new UTF8Encoding(false), false);

                            m.Zieldatei = m.Praktikanten(
                                [
                                    "BW,1", "BT,1", "BS,1", "BS,2", "HBG,1", "HBT,1", "HBW,1", "GG,1", "GT,1", "GW,1", "IFK,1"
                                ],
                                Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-praktikanten-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', new UTF8Encoding(false), false);

                            m.Zieldatei = m.KlassenAnlegen(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-klassen-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', new UTF8Encoding(false), false);

                            m.Schulpflichtüberwachung(configuration);

                            m.Zieldatei = m.GetFaecher(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-faecher.csv"));
                            m.Zieldatei?.Erstellen(",", '\'', new UTF8Encoding(false), false);
                        },
                        Global.Rubrik.Wiki,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    ),
                    new Menüeintrag(
                        "Massen-Mail: Senden",
                        anrechnungen,
                        quelldateien.Notwendige(configuration,["lehrkraefte,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird die Datei [aqua]{Path.Combine(pfadDownloads, "mailadressen.txt")}[/] eingelesen.",
                            "Es werden immer genau 49 Empfänger in BCC angeschrieben.",
                            "Als Inhalt wird das Bild Campusfest_Berufskolleg_Borken.jpg angehängt.",
                            "Die bereits angeschriebenen Empfänger werden in der Datei mailadressen.txt gelöscht und nicht wieder ausgewählt."
                        ],
                        m =>
                        {
                            configuration = Global.Konfig("BetreffMassenmail", Global.Modus.Update, configuration, "Betreff angeben");
                            configuration = Global.Konfig("SmtpServer", Global.Modus.Update, configuration, "Server angeben");
                            configuration = Global.Konfig("SmtpUserMassenmail", Global.Modus.Update, configuration, "Benutzer angeben");
                            configuration = Global.Konfig("SmtpPasswordMassenmail", Global.Modus.Update, configuration, "Passwort angeben");
                            configuration = Global.Konfig("SmtpPort", Global.Modus.Update, configuration, "Port angeben");

                            for(int i = 0; i < 50; i++)
                            {
                                var Mail = new Mail(
                                Path.Combine(Directory.GetCurrentDirectory() ?? "", "mailadressen.txt"),
                                Path.Combine(Directory.GetCurrentDirectory() ?? "", "Campusfest_Berufskolleg_Borken.jpg"),
                                betreffMassenmail,
                                configuration,
                                29 // Anzahl der Mailempfänger in BCC
                            );
                                Console.WriteLine("Warte 60 Sekunden...");
                                Thread.Sleep(60000);
                            }
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur000000
                    )
                ]
            );
        }
        catch (Exception ex)
        {
            var panel3 = new Panel(ex.Message)
            .Header($" [bold red]  Es ist zu einem Fehler gekommen  [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Color.Red);
        
            AnsiConsole.Write(panel3);
            Global.WeiterMitAnykey(configuration);
            return null;            
        }   
    }
}
    