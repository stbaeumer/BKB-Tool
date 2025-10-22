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
        var pfadDownloads = configuration["PfadDownloads"];
        var pfadSchilddatenaustausch = configuration["PfadSchilddatenaustausch"];
        var netmanMailReceiver = configuration["NetmanMailReceiver"];
        var pdfKennwort = configuration["PdfKennwort"];
        var betreffMassenmail = configuration["BetreffMassenmail"];
        var inputFolder = configuration["InputFolder"];
        var outputFolder = configuration["OutputFolder"];

        try
        {
            var students = new Students(configuration, quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "schuelerzusatzdaten,dat"], true));
            klassen = new Klassen(configuration, quelldateien.Notwendige(configuration, ["klassen,dat"], true), students);
            quelldateien.Meldung.Add(students.GetArtUndZahlen());

            var panel = new Panel(string.Join(' ', quelldateien.Meldung))
                .HeaderAlignment(Justify.Center)
                .RoundedBorder()//.SquareBorder()
                .Expand()
                .BorderColor(Global.ColorÜberschrift);

            //AnsiConsole.Write(panel);

            Global.DisplayHeader(configuration, quelldateien.Meldung);

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
                        "Mailadressen: Fehlende Schulinterne Mailadressen in den Individualdaten I ergänzen",
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat","schuelerzusatzdaten,dat", "schueleradressen,dat", "adressen,dat", "schuelertelefonnummern,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", "SchuelerZusatzdaten.dat")}[/] um schulinterne Mailadressen ergänzt und in [{Global.GetColor(Global.ColorPfadInDateien)}]{pfadSchilddatenaustausch}[/] für den Re-Import nach SchILD bereitgestellt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise:[/] ",
                            $"[{Global.GetColor(Global.ColorHinweise)}]1:[/] BKB-Tool bildet die schulinterne Mailadressen wie folgt: [{Global.GetColor(Global.ColorTextHervorheben)}]nv061231@meine-schule.de[/], wobei gilt:",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]    n[/]      : Erster Buchstabe des Nachnamens. Umlaute werden aufgelöst. Bsp.: [{Global.GetColor(Global.ColorTextHervorheben)}]Ü[/] wird zu [{Global.GetColor(Global.ColorTextHervorheben)}]u[/] usw.",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]    v[/]      : Erster Buchstabe des Vornamens. Umlaute werden aufgelöst.",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]    061231[/] : Geburtsdatum in der Notation: JJMMTT.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]2:[/] Vorhandene schulinterne SchILD-Mailadressen in [{Global.GetColor(Global.ColorPfadInProgrammen)}]Individualdaten I[/] bleiben unangetastet.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]3:[/] Doppelungen werden angezeigt und müssen nach Vorgabe behandelt werden."
                        ],
                        m =>
                        {
                            configuration = Global.Konfig("MailDomain", Global.Modus.Update, configuration);
                            m.SchuelerZusatzdatenUmMailAdresseErgaenzen(
                                configuration,
                                Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerZusatzdaten.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.OrdnerOeffnen(),
                                    datei => datei.Erstellen()
                                ],
                                ["Nachname", "Vorname", "Geburtsdatum"],
                                ["BeginnBildungsgang", "Telefon-Nr.","Fax/Mobilnr","bisherige ID"],
                                "|", '\'', new UTF8Encoding(false), false);
                                /*m.SchueleradresseTelefonFormatieren(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerAdressen.dat"),
                                    ["Nachname", "Vorname", "Geburtsdatum"],
                                    [],
                                    "|", '\'', new UTF8Encoding(false), false),
                                m.AdresseTelefonFormatieren(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Adressen.dat"),
                                    ["Nachname", "Vorname", "Geburtsdatum"],
                                    [],
                                    "|", '\'', new UTF8Encoding(false), false),
                                m.SchuelerTelefonnummernFormatieren(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerTelefonnummern.dat"),
                                    ["Nachname", "Vorname", "Geburtsdatum","Art"],
                                    [],
                                    "|", '\'', new UTF8Encoding(false), false)*/
                            
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Webuntis & Co.: Importdateien für Webuntis, Littera, Netman erstellen",
                        quelldateien.Notwendige(configuration, ["legalguardian_,csv,optional","apprenticerepresentative_,csv,optional","student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat", "schuelerTelefonnummern,dat"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt verschiedene Dateien in [bold {Global.GetColor(Global.ColorPfadInDateien)}]{pfadDownloads}[/] erstellt: " +
                            $"[bold {Global.GetColor(Global.ColorPfadInDateien)}]Webuntis-Stammdaten-Schueler.csv[/], [bold {Global.GetColor(Global.ColorPfadInDateien)}]Webuntis-Stammdaten-Betriebe.csv[/], [bold {Global.GetColor(Global.ColorPfadInDateien)}]Webuntis-Stammdaten-Erzieher.csv[/], [bold {Global.GetColor(Global.ColorPfadInDateien)}]-ImportNachLittera.xml[/], [bold {Global.GetColor(Global.ColorPfadInDateien)}]-ImportNachNetman.csv[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]1:[/] Das Zeugnisdatum des letzten Zeugnisses in einer Klasse wird zum Webuntis-Austrittsdatum bei Schüler*innen, deren Status weder aktiv noch extern ist.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]2:[/] Damit korrekt ausgeschult wird, muss auch Abgang und Abschluss beim SchILD-Export angehakt werden.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]3:[/] Für den Betriebeimport sollte im Webuntis-Importprofil die SchildAdressId auf Schlüssel (externe) matchen."
                        ],
                        m =>
                        {
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration, m.Students)) return;
                            m.WebuntisOderNetmanOderLitteraCsv(configuration,
                                [
                                    new Datei(
                                        "ImportNachWebuntis-Stammdaten-Schueler.csv", new string[] { "EMINUSMail" }, new string[] { }, ";", '\'', new UTF8Encoding(false), false,
                                        [
                                            datei => datei.OrdnerOeffnen(),
                                            datei => datei.Erstellen(),
                                            datei => datei.OeffneWebseite("https://nessa.webuntis.com/students"),
                                            datei => datei.OeffneWebseite("https://management.geevoo.de/import/"),
                                            datei => datei.OeffneWebseite("https://nessa.webuntis.com/users")
                                        ],
                                        [
                                            $"1. In Webuntis als Webuntis-Admin:  [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Stammdaten > Schüler*innen > Import[/]",
                                            $"2. Datei auswählen, UTF8",
                                            $"3. Profil: Schuelerimport, dann Vorschau",
                                            $"Mehr zum Profil Schuelerimport: [{Global.GetColor(Global.ColorHyperlink)}][link=https://github.com/stbaeumer/BKB-Tool/wiki]https://github.com/stbaeumer/BKB-Tool/wiki[/][/]"
                                        ]
                                    ),
                                    new Datei(
                                        "ImportNachWebuntis-Stammdaten-Ausbildungsbeauftragte.csv", new string[] { "EMINUSMail" }, new string[] { }, ";", '\'', new UTF8Encoding(false), false,
                                        [
                                            datei => datei.Erstellen(),
                                            datei => datei.OeffneWebseite("https://nessa.webuntis.com/apprenticerepresentatives")
                                        ],
                                        [
                                            $"1. In Webuntis als Webuntis-Admin:  [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Stammdaten > Ausbildungsbeauftragte > Import[/]",
                                            $"2. Datei auswählen, UTF8",
                                            $"3. Profil: Schuelerimport, dann Vorschau",
                                            $"Mehr zum Profil Schuelerimport: [{Global.GetColor(Global.ColorHyperlink)}][link=https://github.com/stbaeumer/BKB-Tool/wiki]https://github.com/stbaeumer/BKB-Tool/wiki[/][/]"
                                        ]
                                    ),
                                    new Datei(
                                        "ImportNachWebuntis-Stammdaten-Erziehungsberechtigte.csv", new string[] { "EMINUSMail" }, new string[] { }, ";", '\'', new UTF8Encoding(false), false,
                                        [
                                            datei => datei.Erstellen(),
                                            datei => datei.OeffneWebseite("https://nessa.webuntis.com/legalguardians")                                            
                                        ],   
                                        [
                                            $"1. In Webuntis als Webuntis-Admin:  [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Stammdaten > Erziehungsberechtigte > Import[/]",
                                            $"2. Datei auswählen, UTF8",
                                            $"3. Profil: Schuelerimport, dann Vorschau",
                                            $"Mehr zum Profil Schuelerimport: [{Global.GetColor(Global.ColorHyperlink)}][link=https://github.com/stbaeumer/BKB-Tool/wiki]https://github.com/stbaeumer/BKB-Tool/wiki[/][/]"
                                        ]
                                    ),
                                    new Datei(
                                        "ImportNachNetman.csv", new string[] { }, new string[] { }, ",", '\'', new UTF8Encoding(false), false,
                                        [
                                            datei => datei.Erstellen(),
                                            datei => datei.ZippenMitKennwort(configuration),
                                            datei => datei.Mailen(datei.ZipPfad, datei.ZipPfad, configuration, datei.ZipPfad),
                                        ],
                                        [
                                            $"Es wird jetzt die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd") + "-ImportNachNetman.csv")}[/] erstellt.",
                                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1:[/] Schüler*innen, die bereits abgegangen sind oder einen Abschluss erworben haben, werden erst sechs Wochen später ausgebucht, um den Zugriff auf Teams nicht direkt zu verlieren.",
                                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #2:[/] Auch Abschluss und Abgang muss beim SchILD-Export angehakt werden.",
                                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #3:[/] Schüler*innen werden 42 Tage nach dem Abgangszeugnis/Abschlusszeugnis ausgeschult. Wenn kein letztes Zeugnisdatum bei Abgängern/Abgeschlossenen ermittelt werden kann, dann wird sofort ausgeschult."
                                        ]
                                    ),
                                    new Datei(
                                        "ImportNachLittera.xml", new string[] { }, new string[] { }, ",", '\'', new UTF8Encoding(false), false,
                                        [
                                            datei => datei.Erstellen(),
                                            datei => datei.Verschieben(configuration["PfadLitteraImport"])
                                        ],
                                        [
                                            $"Es wird jetzt die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(configuration["PfadDownloads"] ?? "", DateTime.Now.ToString("yyyyMMdd-") + @"-ImportNachLittera.csv") + "[/] erstellt.",
                                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1:[/] Auch Abschluss und Abgang muss beim SchILD-Export angehakt werden.",
                                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #2:[/] Schüler*innen werden 42 Tage nach dem Abgangszeugnis/Abschlusszeugnis ausgeschult. Wenn kein letztes Zeugnisdatum bei Abgängern/Abgeschlossenen ermittelt werden kann, dann wird sofort ausgeschult."
                                        ]
                                    )
                                ]
                            );
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),                    
                    new Menüeintrag(
                        "Fotos (c): Schüler*innenfotos aus SchILD für Webuntis und Geevoo bereitstellen",
                        quelldateien.Notwendige(configuration, ["schuelerZusatzdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(configuration["PfadDownloads"] ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-Import-Fotos.zip")}[/] erstellt.",
                            $"[{Global.GetColor(Global.ColorUnterschrift)}]Voraussetzung: [/]Die Fotos müssen zuvor aus SchILD exportiert werden: ([{Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadFotosAusSchild"]}[/].",
                            $"Dazu in SchILD den Weg gehen: [{Global.GetColor(Global.ColorActionInMenüs)}]Datenaustausch > Fotos > Fotos exportieren[/]",
                            $"Wahlweise können alle Fotos bereitgestellt werden oder nur diejenigen, die in SchILD seit dem letzten Fotoexport hinzugefügt wurden.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1: [/]Nach dem Start dieser Funktion werden alle aus SchILD exportierten Fotos nach PfadFotosAusSchild-{DateTime.Now.ToString("yyyyMMdd-HHmm")} verschoben. Somit können die verschiedenen Fotoexporte verglichen werden, um die Differenz für Webuntis zu ermitteln."
                        ],
                        m =>
                        {
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration, m.Students)) return;
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            m.IStudents = m.IStudents.AlleOderNeueFotopfadeAnStudentsZuweisen(configuration);
                            m.Zieldatei = new Datei(Path.Combine(configuration["PfadDownloads"] ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-" + (string.Join("-", m.IKlassen).Substring(0, Math.Min(25, string.Join("-", m.IKlassen).Length)) + "-Fotos.zip")));
                            m.Zieldatei?.FotosZippen(configuration, "", 0, m.IStudents);
                            m.Zieldatei.OrdnerOeffnen();
                            m.OeffneWebseite("https://nessa.webuntis.com/students");
                            m.OeffneWebseite("https://management.geevoo.de/import/");
                            m.NeueFotosAusSchildOrdnerErstellenUndAlteFotosVerschieben(configuration);
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                    "Klassen: Neue Klassen von Untis nach SchILD übergeben und Eigenschaften anpassen",
                    quelldateien.Notwendige(configuration, ["klassen,dat", "GPU003,txt", "GPU002,txt"]),
                    students,
                    klassen,
                    [
                        $"Nachdem die Schule in das neue Schuljahr versetzt worden ist, kann diese Funktion Folgendes:",
                        $"#1 Fehlende Klassen aus Untis werden in den Schuelerzusatzdaten angelegt. Klassenleitung und Jahrgang werden aus Untis übernommen. Andere Eigenschaften werden aus Klassen des selben Jahrgangs der bisherigen Schuelerzusatzdaten übernommen.",
                        $"#2 In bestehenden Klassen werden Klassenleitung und Jahrgang aus Untis in die Schuelerzusatzdaten übernommen.",
                        $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1:[/] Die stellvertretenden Klassenleitungen und die Prüfungsordnung müssen manuell angepasst werden.",
                        $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #2:[/] Klassen ohne Unterrichte werden ignoriert.",
                    ],
                    m =>
                    {
                        m.KlassenErstellen(
                            configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Klassen.dat"),
                            [
                                datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                datei => datei.OrdnerOeffnen(),
                                datei => datei.Erstellen()
                            ],
                            ["InternBez"],
                            ["SonstigeBez", "Folgeklasse"],
                            "|", '\0', new UTF8Encoding(true), false);
                    },
                    Global.Rubrik.Allgemein,
                    Global.NurBeiDiesenSchulnummern.Nur177659
                ),
                 new Menüeintrag(
                                "Klassenbucheinträge: Säumige Lehrer*innen erinnern",
                                quelldateien.Notwendige(configuration, ["lehrkraefte,dat", "openperiod,pdf"]),
                                students,
                                klassen,
                                [
                                    "Die 10% der KuK mit den meisten offenen Klassenbucheinträgen werden (mit folgender Einschränkung) angemahnt: Mit weniger als 10 offenen Eintragungen wird nicht gemahnt. ",
                                    "Ab 20 oder mehr Stunden wird die Schulleitung in CC informiert.",
                                    $"Die Anzahl der offenen Klassenbucheinträge wird aus der Datei [{Global.GetColor(Global.ColorPfadInDateien)}]OpenPeriods[/] ausgelesen.",
                                    "Die KuK werden zuerst angezeigt. Vor dem Mailversand wird nochmal explizit nach Bestätigung gefragt."
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
                    "Wiki: Diverse SQLite-Dateien (Organigramm, Praktikum etc.) erstellen",
                    quelldateien.Notwendige(configuration, ["schuelerzusatzdaten,dat", "absenceperstudent,csv", "GPU006,txt", "GPU002,txt", "GPU003,txt", "klassen,dat"]),
                    students,
                    klassen,
                    [
                        $"Das Organigramm wird aus Untisanrechnungen gebildet. Beispiele: {{...}} = KATEGORIE; [[...]] = HINWEIS, Text ohne Klammern wird zur ROLLE; A14, A15, A16 ohne Klammern > AMT; Untis-Beschreibung > AUFGABE. Im Organigramm wird nach Kategorie, Aufgabe oder Beschreibung gruppiert.",
                        $"Untisanrechnungen: 1.Struct Schema Editor > Untisanrechnungen > Löschen/Leeren > 'untisanrechnungen' eingeben, dann Leeren",
                        $"Untisanrechnungen: 2.Struct Schema Editor > Untisanrechnungen > Importieren/Exportieren > Importieren von Rohdaten > Global > Durchsuchen",
                        $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise zum Text in Anrechnungen:[/]",
                        $"[{Global.GetColor(Global.ColorHinweise)}]#1[/] Das Beförderungsamt wird ausgelesen. Bsp.: A14",
                        $"[{Global.GetColor(Global.ColorHinweise)}]#2[/] Hinweise werden aus eckigen Klammern ausgelesen. Bsp.: Fortbildung 2024",
                        $"[{Global.GetColor(Global.ColorHinweise)}]#3[/] Kategorien werden aus geschweiften Klammern ausgelesen. Bsp.: Technik, Beratung",
                        $"[{Global.GetColor(Global.ColorHinweise)}]#4[/] Bildungsgänge werden daran identifiziert, dass im Text [aqua]Bildungsgangleitung[/] steht und die Beschreibung mit [aqua]bildunggaenge:[/] beginnt, "
                    ],
                    m =>
                    {
                        var anrechnungen = new Anrechnungen(lehrers, configuration);

                        m.GetGruppen(
                            configuration,
                            [
                                zieldatei => zieldatei.Erstellen(),
                                datei => datei.OeffneWebseite("https://bkb.wiki/oeffentlich:organigramm?do=admin&page=struct_schemas&table=gruppen"),
                            ],
                            anrechnungen,
                            Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-gruppen.csv"),
                            lehrers,
                            ",", '\"', new UTF8Encoding(false), true);
                        m.GetUntisAnrechnungen(
                            anrechnungen,
                            Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-untisanrechnungen.csv"),
                            [
                                zieldatei => zieldatei.Erstellen(),
                                datei => datei.OeffneWebseite("https://bkb.wiki/oeffentlich:organigramm?do=admin&page=struct_schemas&table=untisanrechnungen"),
                            ],
                            [500, 510, 530, 590, 900],
                            [500, 510, 530, 590],
                            ["PLA", "BM"],
                            ",", '\"', new UTF8Encoding(false), true);
                        m.GetLehrer(                            
                            [                                
                                zieldatei => zieldatei.Erstellen()
                            ],
                            lehrers,
                            Path.Combine(Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-lul-utf8OhneBom-einmalig-vor-SJ-Beginn.csv")),
                            ",", '\'', new UTF8Encoding(false), false);
                        m.Praktikanten(
                            [
                                zieldatei => zieldatei.Erstellen(),
                                datei => datei.OeffneWebseite("https://bkb.wiki/oeffentlich:organigramm?do=admin&page=struct_schemas&table=praktikanten"),
                            ],
                            [
                                "BW,1,2", "BT,1,2", "BS,1,2", "BS,2,2", "HBG,1,2", "HBT,1,2", "HBT,2,2", "HBW,1,1", "GG,1,1", "GT,1,1", "GW,1,1", "IFK,1,2" // Klasse, Jg, Anzahl Praktika
                            ],
                            Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-praktikanten-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"),
                            ",", '\"', new UTF8Encoding(false), true);
                        m.KlassenAnlegen(
                            configuration,
                            [
                                zieldatei => zieldatei.Erstellen(),
                                datei => datei.OeffneWebseite("https://bkb.wiki/oeffentlich:organigramm?do=admin&page=struct_schemas&table=klassen"),
                            ],
                            Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-klassen-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"), ",", '\"', new UTF8Encoding(false), true);
                        m.Schulpflichtüberwachung(
                            configuration,
                            [
                                zieldatei => zieldatei.Erstellen(),
                                datei => datei.PutPage()
                            ]);
                        m.GetFaecher(
                            configuration,
                            [
                                zieldatei => zieldatei.Erstellen(),
                                datei => datei.OeffneWebseite("https://bkb.wiki/oeffentlich:organigramm?do=admin&page=struct_schemas&table=gruppen"),
                            ],
                            Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-faecher.csv"),
                            ",", '\'', new UTF8Encoding(false), false);
                    },
                    Global.Rubrik.Wiki,
                    Global.NurBeiDiesenSchulnummern.Nur177659
                ),
                new Menüeintrag(
                        "Outlook: CSV-Terminexporte für Wiki aufbereiten",
                        quelldateien.Notwendige(configuration,["termine_fhr,csv", "termine_verwaltung,csv", "termine_berufliches_gymnasium,csv", "termine_kollegium,csv"]),
                        students,
                        klassen,
                        [
                            $"Termine aus Outlook (Kollegium, FHR, Berufliches Gymnasium, Verwaltung) werden nach {Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWiki-Kalendername.csv")} exportiert.",
                            $"[{Global.GetColor(Global.ColorActionInMenüs)}]Vorgehen:[/]",
                            $"[{Global.GetColor(Global.ColorActionInMenüs)}]#1[/] Die Kalender in Listenansicht anzeigen. Notwendige Spalten: [{Global.GetColor(Global.ColorActionInMenüs)}]Beginn, Ende, Betreff, Kategorien, Ressourcen, Ort, Nachricht[/]",
                            $"[{Global.GetColor(Global.ColorActionInMenüs)}]#2[/] Kalender aufsteigend nach Beginn sortieren.",
                            $"[{Global.GetColor(Global.ColorActionInMenüs)}]#3[/] Mit Copy&Paste (Strg+A, Strg+C) die Termine aus Outlook in die CSV-Dateien im Download-Ordner kopieren. Codierung UTF8",
                            $"[{Global.GetColor(Global.ColorActionInMenüs)}]#4[/] Die Kalender im Wiki zuerst leeren ([{Global.GetColor(Global.ColorPfadInProgrammen)}]Admin > Struct Schema Editor > Leeren[/]). Anschließend die neuen CSV als Global importieren.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise:[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#1[/] Falls der Inhalt im Body (Spalte Nachricht) mehrzeilig ist, wird nur die erste Zeile berücksichtigt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#2[/] Es werden nur Termine berücksichtigt, die mindestens eine Kategorie haben. Kategorien werden zu Links in Wiki.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#3[/] Termine aus vergangenen Schuljahren werden nicht mit übertragen.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#4[/] Falls in der Nachricht ein Link zu bkb.wiki enthalten ist, dann wird der Link zum Seitenlink. Ansonsten wird die erste Kategorie zum Seitenlink.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#5[/] Falls in der Nachricht ein Link zu bkb.wiki enthalten ist, wird der Link zur ersten Kategorie.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#6[/] Die Anzahl der Kategorien ist in Outlook begrenzt. Mehr als 6 Kategorien sind evtl. problematisch.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#7[/] Mehrtägige Termine: Nur bei ganztägigen Terminen wird der erste und letzte Tag richtig angezeigt. Wenn Uhrzeiten angegeben werden, wird nur der erste Tag angezeigt."
                        ],
                        m =>
                        {
                            m.OeffneDateienInDownloadsInNotepadPlusPlus(configuration, ["termine_fhr.csv", "termine_verwaltung.csv", "termine_berufliches_gymnasium.csv", "termine_kollegium.csv"]);                            
                            foreach (var kalender in new List<string>(){"termine_berufliches_gymnasium", "termine_kollegium", "termine_verwaltung", "termine_fhr" })
                            {
                                m.Kalender2Wiki(configuration,
                                [
                                    datei => datei.OrdnerOeffnen(),
                                    datei => datei.Erstellen(),
                                    datei => datei.OeffneWebseite($"https://bkb.wiki/start?do=admin&page=struct_schemas&table=" + kalender),
                                    datei => datei.OeffneWebseite($"https://bkb.wiki/start?do=admin&page=struct_schemas&table=" + kalender)
                                ],
                                kalender, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWiki-" + kalender), ",", '\"', new UTF8Encoding(false), true);
                            }
                        },
                        Global.Rubrik.Wiki,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Fotos (a): Schüler*innen klassenweise fotografieren, kopieren, umbenennen, ablegen",
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "schuelerzusatzdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Am Einschulungstag unterstützt diese Funktion beim klassenweisen Fotografieren der Schüler*innen. Konkret beim Kopieren, Verkleinern (160*160), Umbenennen und Zippen der Fotos.",
                            $"Erstellen Sie jetzt quadratische Fotos der vor Ihnen stehenden Klasse (z.B. mit dem Handy). Dabei ist die [{Global.GetColor(Global.ColorInfoBox)}]Reihenfolge & Anzahl[/] laut folgender Tabelle exakt einzuhalten. Die Fotos werden in den Schild-Fotoordner kopiert und dabei mit der schulischen E-Mail der SuS umbenannt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1:[/] Wenn jemand fehlt, dann die weiße Wand fotografieren, damit [{Global.GetColor(Global.ColorInfoBox)}]Reihenfolge & Anzahl[/] stimmen.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #2:[/] Wenn ein Foto nicht gelungen ist, dann löschen und neu erstellen.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #3:[/] Wenn mehr als eine Klasse ausgewählt wird, wird nur die erste Klasse berücksichtigt",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #4:[/] Das Kriterium für die Reihenfolge ist der Dateiname.",
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration, "Klasse", "Geben Sie den Namen der Klasse an, die jetzt vor Ihnen steht.");
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration, m.IStudents)) return;
                            m.IStudents.KlassenordnerErstellenFotosOrdnerÖffnen(configuration);
                            m.IStudents.KlassenListenAnzeigen(configuration);
                            m.IStudents.AnzahlPrüfen(configuration);
                            m.IStudents.KlassenordnerInZielPfadErstellen(configuration);
                            m.IStudents.FotosZuStudentsZuweisen(configuration);
                            m.IStudents.FotosNachZielordnerKopieren(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Fotos (b): Schüler*innenfotos nach SchILD2 hochladen",
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "schuelerzusatzdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt die Fotos nach SchILD2 hochgeladen. Wenden Sie diese Funktion an, wenn Sie zuvor klassenweise Fotos mit [bold {Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] erstellt haben.",
                            $"[{Global.GetColor(Global.ColorUnterschrift)}]Voraussetzung #1: [/]Die Fotos müssen aus SchILD exportiert werden: [{Global.GetColor(Global.ColorActionInMenüs)}]Datenaustausch > Fotos > Fotos exportieren[/]. Beachte, dass der Fotodateiname den Nachnamen, Vornamen und das Geburtsdatum enthält.",
                            $"[{Global.GetColor(Global.ColorUnterschrift)}]Voraussetzung #2: [/]Alle noch nicht zu SchILD hochgeladenen Fotos liegen in einem vorbereiteten Ordner. Beachte: Die Fotos müssen den ersten Teil der E-Mail-Adresse (also alles vor dem @) als Dateinamen haben.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1: [/]Vorhandene Bilder werden nicht überschrieben, sofern der Export aus SchILD zuvor getätigt wurde."
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            m.IStudents.FotosFürUploadNachSchildAuswählen(configuration);
                            m.IStudents.FotosRotieren(configuration);
                            m.IStudents.FotosFürUploadNachSchild2AuflistenUndBestätigen(configuration);
                            m.IStudents.FotosNachSchild2Hochladen(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Schnellmeldung: Relationsgruppen im September aufbereiten",
                        quelldateien.Notwendige(configuration,[]),
                        students,
                        klassen,
                        [
                            "Alle aktiven Schüler*innen und Gastschüler*innen werden angezeigt.",
                            "Weiterleitung der Zahlen die Bereichsleitungen zur Prüfung. Bereichsleitungen werden sodann aufgefordert Änderungen zum Stichtag mitzuteilen.",
                            "Relationen gemäß §93 SchulG",
                            "Dokumentation siehe schips.nrw.de"
                        ],
                        m =>
                        {
                            m.Relationsgruppen = new Relationsgruppen(klassen, students, configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Statistik: Unterrichtsverteilung und Anrechnungen nach SchILD importieren",
                        quelldateien.Notwendige(configuration, ["studentgroupstudents,csv", "klassen,dat", "schuelerlernabschnitt,dat,optional", "schuelerleistungsdaten,dat", "schuelerbasis,dat", "lehrkraefte,dat", "kurse,dat", "lehrkraeftesonderzeiten,dat", "schuelerbasisdaten,dat", "GPU002,txt", "GPU020,txt", "GPU004,txt", "faecher,dat"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt folgende Dateien für den Import nach SchILD erstellt: \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Lernabschnitte.dat")}[/] \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Leistungsdaten.dat")}[/] \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat")}[/]\n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat")}[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise:[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#1[/] Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]LehrkraefteSonderzeiten.dat[/] wird nicht komplett neu erstellt. Die exportierte Datei wird lediglich für den Re-Import aufbereitet.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#2[/] Die Kursbezeichnungen setzen sich zusammen aus dem Kursleiterkürzel plus alle beteiligten Untis-Unterrichtsnummern (bis maximal 20 Zeichen).",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#3[/] Kurse werden gebildet aus: Kopplungen in Untis, Schülergruppen in Untis, identischen Fächern mit unterschiedliche LuL",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#4[/] Zähler im Anschluss an Fächer (M1, M2, ...) werden abgeschnitten (also zu M).",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#5[/] Bei mehreren beteiligten Lehrkräften wird das alphabetisch erste Lehrkraftkürzel zum Kursleiter.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#6[/] Team-Teaching ist daran erkennbar, dass die Summe der Kurs-Wochenstunden kleiner ist als die Summe der Lehrkräfte-Wochenstunden.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#7[/] Wenn zwei ansonsten identische Unterrichte einmal mit und einmal ohne Schülergruppe vorliegen, werden zwei unterschiedliche Einträge erstellt."
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            configuration = Global.Konfig("Abschnitt", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("Schulnummer", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("StatistikDatum", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("Kursarten", Global.Modus.Read, configuration);

                            m.Unterrichte = new Unterrichte(configuration, m, Global.Zweck.Statistik, Global.Art.KursUnterrichte);
                            m.Unterrichte.AddRange(new Unterrichte(configuration, m, Global.Zweck.Statistik, Global.Art.NichtKursUnterrichte));

                            m.Lernabschnittsdaten(
                                configuration, Global.Zweck.Statistik, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLernabschnittsdaten.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.OrdnerOeffnen(),
                                    datei => datei.Erstellen()
                                ],
                                ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);
                            m.Kurse(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.Erstellen()
                                ],
                                ["KursBez", "Jahr", "Abschnitt"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);
                            m.LeistungsdatenStatistik(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.Erstellen()
                                ],
                                ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach", "Kurs"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false, null,
                                Global.Zweck.Statistik);
                            m.Faecher(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.Erstellen()
                                ],
                                ["InternKrz"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);                            
                        },
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        $"Altersermäßigung: berechnen für {int.Parse(Global.AktSj[0])}/{int.Parse(Global.AktSj[0]) + 1} und {int.Parse(Global.AktSj[0]) + 1}/{int.Parse(Global.AktSj[0]) + 2}",
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
                                m.LehrkraefteSonderzeiten(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat"),
                                    [
                                        datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                        datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                        datei => datei.OrdnerOeffnen(),
                                        datei => datei.Erstellen()
                                    ],
                                    ["Lehrkraft", "Zeitart", "Grund"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false, null, "200", Global.Modus.ReadSilent);
                                m.Lehrkraefte(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Lehrkraefte.dat"),
                                    [
                                        datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                        datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                        datei => datei.Erstellen()
                                    ],
                                    ["InternKrz"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false);
                                m.ExportAusSchildVerschieben(configuration);
                            },
                            Global.Rubrik.Allgemein,
                            Global.NurBeiDiesenSchulnummern.Nur177659
                        ),
                    new Menüeintrag(
                        "Sonderzeiten: Lehrkräftesonderzeiten (Anrechnungen) nach SchILD importieren",
                        quelldateien.Notwendige(configuration, ["studentgroupstudents,csv", "klassen,dat", "schuelerlernabschnitt,dat,optional", "schuelerleistungsdaten,dat", "schuelerbasis,dat", "lehrkraefte,dat", "kurse,dat", "lehrkraeftesonderzeiten,dat", "schuelerbasisdaten,dat", "GPU002,txt", "GPU020,txt", "GPU004,txt", "faecher,dat"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt folgende Dateien für den Import nach SchILD erstellt: [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat")}[/], [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Lehrkraefte.dat")}[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise:[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#1[/] Da die Sonderzeiten nur in Kombination mit den Lehrkräften importiert werden, wird eine leere Lehrkraefte.dat erzeugt. Der Import der leeren Lehrkraefte.dat ist unschädlich.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#2[/] Die Gründe können gleichzeitig oder nach und nach sukzessive angegeben werden.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#3[/] Wenn ein Grund bei einer Lehrkraft wegfällt (also Wert = 0), dann wird das im Folgenden angezeigt. Die Zeile muss dann händisch in SchILD gelöscht werden. Neue und veränderte Werte werden in der LehrkräfteSonderzeiten.dat hinzugefügt.",
                        ],
                        m =>
                        {
                            m.LehrkraefteSonderzeiten(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.OrdnerOeffnen(),
                                    datei => datei.Erstellen()
                                ],
                                ["Lehrkraft", "Zeitart", "Grund"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);
                            m.Lehrkraefte(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Lehrkraefte.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.Erstellen()
                                ],
                                ["InternKrz"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        $"PDF verschlüsseln: Von PDF-Dateien in {configuration["PfadDownloads"]} verschlüsselte Kopien erstellen",
                        new Dateien(),
                        students,
                        klassen,
                        [
                            $"Von allen PDF-Dateien in [{Global.GetColor(Global.ColorPfadInDateien)}]" + configuration["PfadDownloads"] + "[/] werden verschlüsselte Kopien erstellt.",
                            $"Es werden nur Dateien berücksichtigt, die nicht bereits die Endung [{Global.GetColor(Global.ColorPfadInDateien)}]-kennwort.pdf[/] haben.",
                            $"Kopien bekommen die Dateiendung [{Global.GetColor(Global.ColorPfadInDateien)}]-kennwort.pdf[/]."
                        ],
                        _ =>
                        {
                            var pdfDateien = new PdfDateien();
                            pdfDateien.KennwortSetzen(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                            $"PDF-Seiten mailen: PDF-Seiten an darauf enthaltene E-Mail-Adressen mailen",
                            quelldateien.Notwendige(configuration, ["lehrkraefte,dat"]),
                            students,
                            klassen,
                            [
                                $"Die zuletzt bearbeitete PDF-Datei in [{Global.GetColor(Global.ColorPfadInDateien)}]{configuration["PfadDownloads"]}[/] wird eingelesen und jede Seite der Datei wird nach E-Mail-Adressen durchsucht. Wenn auf einer Seite eine oder mehrere E-Mail-Adressen gefunden werden, dann wirde die betreffenden Seiten an die enthaltene(n) E-Mail-Adresse(n) gemailt. Das ursprüngliche PDF-Dokument wird also bei Bedarf in mehrere PDF-Dokumente aufgeteilt.",
                                $"Nutzen Sie diese Funktion, um beispielsweise, um eine Datei mit allen Studenplänen an Lehrkräfte zu mailen.",
                                $"Hinweise:",
                                $"1. Die zuletzt bearbeitete PDF-Datei wird eingelesen.",
                                $"2. Jede Seite der Datei wird nach E-Mail-Adressen durchsucht.",
                                $"3. Die betreffenden Seiten werden an die E-Mail-Adressen gemailt.",
                                $"4. Optional wird verschlüsselt."
                            ],
                            m =>
                            {
                                var pdfDatei = Directory.GetFiles(pfadDownloads, "*.pdf").OrderByDescending(File.GetLastWriteTime).FirstOrDefault();
                                Global.ZeileSchreiben("Die neueste PDF-Datei wird versendet:", pdfDatei, ConsoleColor.White, ConsoleColor.Black);
                                Global.Konfig("SmtpServer", Global.Modus.Update, configuration);
                                Global.Konfig("SmtpPort", Global.Modus.Update, configuration);
                                Global.Konfig("SmtpUser", Global.Modus.Update, configuration);
                                Global.Konfig("SmtpKennwort", Global.Modus.Update, configuration);
                                Global.Konfig("PdfKennwort", Global.Modus.Update, configuration);
                                Global.Konfig("Betreff", Global.Modus.Update, configuration);
                                Global.Konfig("Body", Global.Modus.Update, configuration);
                                Global.Konfig("BccAdresse", Global.Modus.Update, configuration);

                                AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("E-Mails verarbeiten ...", ctx =>
                                {
                                    foreach (PdfSeite seite in (new PdfDatei(configuration, pdfDatei, new Lehrers(configuration, m.Quelldateien))).Seiten)
                                    {
                                        seite?.GetMailReceiver(lehrers);
                                        seite?.PdfDocumentCreate(pdfDatei);
                                        if (!string.IsNullOrEmpty(pdfKennwort) && pdfKennwort != "kein" && pdfKennwort != "-")
                                            seite?.PdfDocumentEncrypt(pdfKennwort);
                                        seite?.Mailen(configuration);
                                    }
                                });
                            },
                            Global.Rubrik.Allgemein,
                            Global.NurBeiDiesenSchulnummern.Alle
                        ),
                        new Menüeintrag(
                            "PDF-Zeugnisse: Auslesen und in die SchILD-Dokumentenverwaltung einsortieren",
                            quelldateien.Notwendige(configuration, []),
                            students,
                            klassen,
                            [
                                $"PDF-Zeugnisse und andere PDF-Dateien in [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(configuration["PfadDownloads"], "PDF-Input")}[/] werden in die Schüler*innen-Ordner der SchILD-Dokumentenverwaltung einsortiert.",
                                $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung #1[/]: Zu kopierende PDF-Dateien nach [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(configuration["PfadDownloads"], "PDF-Input")}[/] kopieren.",
                                $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung #2[/]: Eine UTF8-CSV-Datei mit Spalten: Nachname, Vorname, Geburtsdatum und Klasse aus Atlantis exportieren und in [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(configuration["PfadDownloads"], "PDF-Input")}[/] ablegen.",
                                $"[{Global.GetColor(Global.ColorTextHervorheben)}]Durchführung #1[/]: (Einzelne) Klasse(n) oder 'alle' auswählen.",
                                $"[{Global.GetColor(Global.ColorTextHervorheben)}]Durchführung #2[/]: Geben Sie die Schlüsselwörter an, um die interessierenden PDF-Dateien einzugrenzen.",
                            ],
                            m =>
                            {
                                m.IStudents.GetStudentsVonAtlantisCsv(configuration);
                                m.IStudents.PdfDateienVerarbeiten(configuration);
                            },
                            Global.Rubrik.Allgemein,
                            Global.NurBeiDiesenSchulnummern.Alle
                        ),
                        new Menüeintrag(
                            "Zeugnis-Chat: Säumige Lehrer*innen im Teams-Chat an die Noten-Eintragung erinnern",
                            quelldateien.Notwendige(configuration, ["schuelerleistungsdaten,dat"]),
                            students,
                            klassen,
                            [
                                $"Wenn die Frist zur Eintragung der Zeugnisnoten abgelaufen ist, können hier gezielt diejenigen Lehrkräfte erinnert werden, deren Noten noch fehlen. ",
                                $"[{Global.GetColor(Global.ColorHinweise)}]Schritt 1: [/]Alle Leistungsdaten (soweit vorhanden) nach SchILD importieren.",
                                $"[{Global.GetColor(Global.ColorHinweise)}]Schritt 2: [/][{Global.GetColor(Global.ColorPfadInDateien)}]SchuelerLeistungsdaten.dat[/] aus SchILD exportieren.",
                                $"[{Global.GetColor(Global.ColorHinweise)}]Schritt 3: [/][{Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] liest die [{Global.GetColor(Global.ColorPfadInDateien)}]SchuelerLeistungsdaten.dat[/] aus SchILD und öffnet Teams-Chat."
                            ],
                            m =>
                            {
                                m.FilterInteressierendeStudentsUndKlassen(configuration);
                                var lul = m.LuLAnEintragungDerZeugnisnotenErinnern(configuration, lehrers);
                                m.ChatErzeugen(configuration, lul);
                            },
                            Global.Rubrik.Leistungsdaten,
                            Global.NurBeiDiesenSchulnummern.Nur177659
                        ),
                        new Menüeintrag(
                        "Zeugnisse: Leistungsdaten (Unterrichte, Zeugnisnoten, Fehlzeiten, ...) nach SchILD importieren",
                        quelldateien.Notwendige(configuration, ["absenceperstudent,csv", "studentgroupstudents,csv", "klassen,dat", "schuelerlernabschnitt,dat,optional", "schuelerleistungsdaten,dat", "lehrkraefte,dat", "kurse,dat", "schuelerbasisdaten,dat", "GPU002,txt", "GPU004,txt", "faecher,dat"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt folgende Dateien für den Import nach SchILD erstellt: \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Lernabschnitte.dat")}[/] \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Leistungsdaten.dat")}[/] \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat")}[/]\n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat")}[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise:[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#1[/] Die Kursbezeichnungen setzen sich zusammen aus dem Kursleiterkürzel plus alle beteiligten Untis-Unterrichtsnummern (bis maximal 20 Zeichen).",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#2[/] Kurse werden gebildet aus: Kopplungen in Untis, Schülergruppen in Untis, identischen Fächern mit unterschiedliche LuL",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#3[/] Zähler im Anschluss an Fächer (M1, M2, ...) werden abgeschnitten (also zu M).",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#4[/] Bei mehreren beteiligten Lehrkräften wird das alphabetisch erste Lehrkraftkürzel zum Kursleiter.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#5[/] Team-Teaching ist daran erkennbar, dass die Summe der Kurs-Wochenstunden kleiner ist als die Summe der Lehrkräfte-Wochenstunden.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#6[/] Wenn zwei ansonsten identische Unterrichte einmal mit und einmal ohne Schülergruppe vorliegen, werden zwei unterschiedliche Einträge erstellt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#7[/] Sobald einem Fach bei einem Schüler in SchILD eine Kursart (AB3,AB4,GKS,GKM) zugewiesen wurde, wird das in die zu erstellenden Leistungsdaten übernommen."
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            configuration = Global.Konfig("Abschnitt", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("Schulnummer", Global.Modus.Read, configuration);                            
                            configuration = Global.Konfig("ZeugnisDatum", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("Kursarten", Global.Modus.Read, configuration);

                            m.Unterrichte = new Unterrichte(configuration, m, Global.Zweck.Zeugnis, Global.Art.KursUnterrichte);
                            m.Unterrichte.AddRange(new Unterrichte(configuration, m, Global.Zweck.Zeugnis, Global.Art.NichtKursUnterrichte));

                            m.Lernabschnittsdaten(
                                configuration, Global.Zweck.Zeugnis, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLernabschnittsdaten.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.OrdnerOeffnen(),
                                    datei => datei.Erstellen()
                                ],
                                ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);
                            m.Kurse(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.Erstellen()
                                ],
                                ["KursBez", "Jahr", "Abschnitt"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);
                            m.LeistungsdatenStatistik(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.Erstellen()
                                ],
                                ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach", "Kurs"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false, null,
                                Global.Zweck.Zeugnis);
                            m.Faecher(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat"),
                                [
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                    datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                    datei => datei.Erstellen()
                                ],
                                ["InternKrz"],
                                [],
                                "|", '\0', new UTF8Encoding(true), false);                            
                        },
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Klausurbelegung: Wiki-Seite (mit Zuordnung der SuS ZU allen Fächern) erstellen",
                        quelldateien.Notwendige(configuration, ["faecher,dat","gpu002,txt","kurse.,dat","studentgroupstudents,csv", "schuelerleistungsdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird für jede Klasse eine Wiki-Seite erstellt. Die Seite enthält eine Tabelle mit allen Schüler*innen der Klasse und allen Fächern, die die Schüler*innen belegen. Klassenleitungen können dort einfach alle Belegungen eintragen. Am besten wird die Tabelle von Klassenleitungen vor den Sommerferien erstellt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise:[/]",
                            $"[{Global.GetColor(Global.ColorHinweise)}]#1[/] Für das Erstellen wird die Schuelerleistungsdaten.dat ausgelesen."
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);                            
                            m.Unterrichte = new Unterrichte(configuration, m, Global.Zweck.Zeugnis, Global.Art.KursUnterrichte);
                            m.Unterrichte.AddRange(new Unterrichte(configuration, m, Global.Zweck.Zeugnis, Global.Art.NichtKursUnterrichte));
                            Global.Konfig("Klausurbelegung", Global.Modus.Update, configuration, "", -1, -1, "", "1", null,
        "1,2,3");
                            Global.Konfig("InteressierendesSchuljahr", Global.Modus.Update, configuration, "", -1, -1, "", "1", null,
        "25-26,26-27,27-28");
                            if(configuration["Klausurbelegung"] == "1" || configuration["Klausurbelegung"] == "2")
                                m.KlausurbelegungWikiSeiteErstellen(
                                    configuration,
                                    $"oeffentlich:klausurbelegung:{configuration["InteressierendesSchuljahr"]}:",
                                    [
                                        datei => datei.PutPage(),
                                        datei => datei.OeffneWebseite($"https://bkb.wiki/{datei.Name}"),
                                    ]);
                            if(configuration["Klausurbelegung"] == "3")
                                m.KlausurbelegungAusWikiNachSchildEinlesen(
                                    configuration,
                                    Path.Combine(pfadSchilddatenaustausch ?? "", "schuelerLeistungsdaten.dat"),
                                    [
                                        datei => datei.OrdnerOeffnen(),
                                        datei => datei.Erstellen()
                                    ],
                                    ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Leistungsdaten,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                        new Menüeintrag(
                            "Teams-Chat: Teams-Chat mit Gruppe von Lehrkräften beginnen",
                            quelldateien.Notwendige(configuration, ["gpu002,txt","gpu003,txt"]),
                            students,
                            klassen,
                            [
                                $"Mit dieser Funktion wird ein Teams-Chat-Link im Browser geöffnet. Der Chat enthält alle Lehrkräfte der im Folgenden ausgewähten Gruppen.",
                                $"Die Gruppen werden aus den Anrechnungen sowie der Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", "GPU002.TXT")}[/] gebildet."
                            ],
                            m =>
                            {
                                var anrechnungen = new Anrechnungen(lehrers, configuration);
                                m.GetGruppen(
                                    configuration,
                                    [
                                        datei => datei.Auswählen(m, lehrers),
                                        datei => datei.OeffneWebseite("https://teams.microsoft.com/l/chat/0/0?users=", datei.UrlMitte, datei.UrlRechts),
                                    ],
                                    anrechnungen,
                                    Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-gruppen.csv"),
                                    lehrers,
                                    ",", '\"', new UTF8Encoding(false), true);
                            },
                            Global.Rubrik.Allgemein,
                            Global.NurBeiDiesenSchulnummern.Nur177659
                        ),
                        new Menüeintrag(
                                "Teilleistungen: SchuelerTeilleistungen.dat erstellen",
                                quelldateien.Notwendige(configuration, ["schuelerteilleistungen,dat", "schuelerleistungsdaten,dat", "schuelerbasisdaten,dat", "marksperlesson,csv"]),
                                students,
                                klassen,
                                [
                                    $"Es wird jetzt die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerTeilleistungen.dat")}[/] erstellt.",
                                    $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Damit der Import nach SchILD reibungslos funktioniert, müssen zuvor die Teilleistungsarten in SchILD ([{Global.GetColor(Global.ColorPfadInProgrammen)}]Schulverwaltung > Teilleistungsarten[/]) gleichlautend mit dem Langnamen in Webuntis ([{Global.GetColor(Global.ColorPfadInProgrammen)}]Stammdaten > Prüfungsarten[/]) angelegt werden.",
                                    $"[{Global.GetColor(Global.ColorHinweise)}]Wichtig:[/] Um feststellen zu können wo Teilleistungen fehlen, müssen die Leistungsdaten bereits in SchILD importiert bzw. angelegt worden sein."
                                ],
                                m =>
                                {
                                    m.FilterInteressierendeStudentsUndKlassen(configuration);
                                    m.Teilleistungen(
                                        configuration,
                                        Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerTeilleistungen.dat"),
                                        lehrers,
                                        [
                                            datei => datei.Verarbeiten(quelldateien, Global.Modus.Vergleichen),
                                            datei => datei.Verarbeiten(quelldateien, Global.Modus.Filtern),
                                            datei => datei.OeffneWebseite("https://teams.microsoft.com/l/chat/0/0?users=", datei.UrlMitte, datei.UrlRechts),
                                            datei => datei.Erstellen()
                                        ],
                                        ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach", "Datum"],
                                        [],
                                        "|", '\0', new UTF8Encoding(true), false);                                    
                                },
                                Global.Rubrik.Allgemein,
                                Global.NurBeiDiesenSchulnummern.Alle
                            ),
                            new Menüeintrag(
                                "Sprechtag: Lehrerübersichtsseite im Wiki veröffentlichen",
                                quelldateien.Notwendige(configuration, ["exportlessons,csv", "gpu004,txt", "gpu005,txt"]),
                                students,
                                klassen,
                                [
                                    $"Die Wiki-Datei [{Global.GetColor(Global.ColorPfadInDateien)}]sprechtag.txt[/] wird über XML-RPC aktualisiert.",
                                    $"[{Global.GetColor(Global.ColorHinweise)}]Hinweise:[/]",
                                    $"[{Global.GetColor(Global.ColorHinweise)}]1:[/]Die Wunschräume müssen in den Untis-Stammdaten beim Lehrer eingetragen werden. Dazu am besten die Fenstergruppe Sprechtag in Untis öffnen.",
                                    $"[{Global.GetColor(Global.ColorHinweise)}]2:[/]Fußnoten werden als Text2 in den Untis-Stammdaten eingetragen. Beispiel für eine Fußnote: [{Global.GetColor(Global.ColorHinweise)}]'außer Haus; bitte Termin vereinbaren;'[/]",
                                    $"[{Global.GetColor(Global.ColorHinweise)}]3:[/]Lehrkräfte ohne Raum werden in der Liste ignoriert. Bei Abwesenheiten die Räume für kommendes Jahr stehen lassen, wenn im Betreff 'außer Haus' steht, dann wird der Raum nicht angezeigt.",
                                    $"[{Global.GetColor(Global.ColorHinweise)}]4:[/]Lehrkräfte ohne eigenen Unterricht bleiben unberücksichtigt."
                                ],
                                m =>
                                {
                                    m.Sprechtag(
                                        configuration,
                                        "oeffentlich:sprechtag",
                                        [
                                            datei => datei.PutPage(),
                                            datei => datei.OeffneWebseite("https://bkb.wiki/oeffentlich:sprechtag"),
                                        ],                                        
                                        "Zum jährlichen Sprechtag laden wir sehr herzlich am Mittwoch nach der allgemeinen Zeugnisausgabe in der Zeit von 13:30 bis 17:30 Uhr ein. Der Unterricht endet nach der 5. Stunde um 12:00 Uhr.");
                                },
                                Global.Rubrik.Wiki,
                                Global.NurBeiDiesenSchulnummern.Nur177659
                            )
                           
                /*,
                                new Menüeintrag(
                                    "Zeugnisse #1: Lernabschnittsdaten: Fehlzeiten von Webuntis nach SchILD importieren",
                                    anrechnungen,
                                    quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "absenceperstudent,csv", "schuelerlernabschnitt,dat"]),
                                    students,
                                    klassen,
                                    [
                                        $"Die Fehlzeiten aktiver Schüler*innen aus Webuntis werden in der [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLernabschnittsdaten.dat")}[/] für den Import nach SchILD vorbereitet.",
                                        $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung 1:[/] Lernabschnitte in SchILD anlegen.",
                                        $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung 2:[/] Klassenleitungen müssen offene Fehlstunden auf [{Global.GetColor(Global.ColorHinweise)}](nicht) entschuldigt[/] setzen. Anderenfalls bleiben die Fehlzeiten auf dem Zeugnis unberücksichtigt. ",
                                        $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung 3:[/] Alle *.dat-Dateien aus SchILD exportieren. "
                                    ],
                                    m =>
                                    {
                                        m.FilterInteressierendeStudentsUndKlassen(configuration);
                                        m.Zieldatei = m.Lernabschnittsdaten(configuration, Global.Zweck.Zeugnis, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLernabschnittsdaten.dat"));
                                        m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                                        m.Zieldateien.ExportAusSchildVerschieben(configuration);
                                        m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);
                                    },
                                    Global.Rubrik.Leistungsdaten,
                                    Global.NurBeiDiesenSchulnummern.Alle
                                ),                    
                                new Menüeintrag(
                                    "Zeugnisse #2: Kurse, Unterrichte und Gesamtnoten von Webuntis nach SchILD importieren",
                                    anrechnungen,
                                    quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "absenceperstudent,csv", "schuelerlernabschnitt,dat", "schuelerleistungsdaten,dat", "schuelerbasis,dat", "exportlessons,csv", "studentgroupstudents,csv", "marksperlesson,csv", "klassen,dat"]),
                                    students,
                                    klassen,
                                    [
                                        $"Die Kurse und Unterrichte (mit Noten) werden in [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat")}[/] und [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat")}[/] vorbereitet.",
                                        $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung:[/] Lernabschnitte in SchILD anlegen und dann alle *.dat-Dateien frisch exportieren. ",                            
                                        $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Falls mehrere Kollegen dasselbe Fach zeitgleich unterrichten, dann muss ein Zähler an das Fach angehangen werden. Bsp.: Zwei LuL unterrichten Mathe: Dann M und M1. Beide Fächer müssen in SchILD existieren. Damit M1 in den Leistungsdaten erscheint, aber nicht auf dem Zeugnis gedruckt wird, muss die Eigenschaft 'Nicht auf Zeugnis drucken' in SchILD gesetzt werden.",                            
                                    ],
                                    m =>
                                    {
                                        m.FilterInteressierendeStudentsUndKlassen(configuration);

                                        //m.Zieldatei = m.Kurse(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat"));
                                        m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["KursBez"], ["Klasse", "Schulnr", "WochenstdPUNKTLEERZEICHENKL"]);
                                        m.Zieldateien.ExportAusSchildVerschieben(configuration);
                                        m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);

                                        //m.Zieldatei = m.Leistungsdaten(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"), Global.Zweck.Zeugnis);
                                        m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt", "Fach"], ["Jahrgang"]);
                                        m.Zieldateien.ExportAusSchildVerschieben(configuration);
                                        m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);
                                    },
                                    Global.Rubrik.Leistungsdaten,
                                    Global.NurBeiDiesenSchulnummern.Alle
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
                                        //m.Zieldatei = m.Leistungsdaten(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"), Global.Zweck.Mahnung);
                                        m.Zieldatei?.Erstellen("|", '\0', new UTF8Encoding(true), false);
                                    },
                                    Global.Rubrik.Leistungsdaten,
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
                                )
                                                                       */
                ]
            );
        }
        catch (Exception ex)
        {
            throw ex;
        }   
    }
}