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
                        "Webuntis: Schüler*innen-Importdatei für Webuntis erstellen",
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd") + "-ImportNachWebuntis.csv")}[/] erstellt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Das Zeugnisdatum des letzten Zeugnisses in einer Klasse wird zum Webuntis-Austrittsdatum bei Schüler*innen, deren Status weder aktiv noch extern ist."
                        ],
                        m =>
                        {
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration)) return;
                            m.Zieldateien =
                            [
                                m.WebuntisOderNetmanOderLitteraCsv(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") +  @"-ImportNachWebuntis.csv"), ";", '\'', new UTF8Encoding(false), false,
                                [
                                    $"1. In Webuntis als Webuntis-Admin:  [bold {Global.GetColor(Global.ColorPfadInProgrammen)}]Stammdaten > Schüler*innen > Import[/]",
                                    $"2. Datei auswählen, UTF8",
                                    $"3. Profil: Schuelerimport, dann Vorschau",
                                    $"Mehr zum Profil Schuelerimport: [{Global.GetColor(Global.ColorHyperlink)}][link=https://github.com/stbaeumer/BKB-Tool/wiki]https://github.com/stbaeumer/BKB-Tool/wiki[/][/]"
                                ]),
                            ];
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Webuntis-Fotos: Zipdatei mit Fotos für Webuntis erstellen",
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [aqua]{Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd") + "-ImportNachWebuntis.zip")}[/] erstellt.",
                            "[pink3]Hinweis:[/] Schüler*innen, deren Foto hochgeladen wurden, werden in der Datei [aqua]fotos.txt[/] gespeichert, um ein erneutes Hochladen zu vermeiden."
                        ],
                        m =>
                        {
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration)) return;
                            m.IStudents = m.Students.OhneWebuntisFoto(configuration, Path.Combine(Directory.GetCurrentDirectory(), "fotos.txt"));
                            m.IStudents.FotosFürWebuntisZippen(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") +  @"-ImportNachWebuntis.zip"), Path.Combine(Directory.GetCurrentDirectory(), "fotos.txt"),
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
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(configuration["PfadDownloads"] ?? "", DateTime.Now.ToString("yyyyMMdd-") + "****" +  @"-ImportNachLittera.csv") + "[/] erstellt.",
                        ],
                        m =>
                        {
                            var zeitstempel = DateTime.Now.ToString("yyyyMMdd-HHmm");
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration)) return;
                            m.Zieldateien =
                            [
                                m.WebuntisOderNetmanOderLitteraCsv(configuration, Path.Combine(configuration["PfadDownloads"] ?? "", DateTime.Now.AddHours(1).ToString("yyyyMMdd-HHmm") + @"-ImportNachLittera.xml"), ",", '\'', new UTF8Encoding(false), false)
                            ];

                            if(configuration["Schulnummer"] != null && configuration["Schulnummer"] == "177659")
                            {
                                configuration = Global.Konfig("PfadLitteraImport", Global.Modus.Update, configuration, "Littera-Import-Pfad");
                                if(m.Zieldateien[0] == null) return;
                                    m.Zieldatei?.Verschieben(configuration["PfadLitteraImport"]);
                            }
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Netman: Schüler*innen-Importdatei für Netman erstellen",
                        quelldateien.Notwendige(configuration, ["student_,csv","schuelerlernabschnittsdaten,dat", "schuelerzusatzdaten,dat", "schuelererzieher,dat", "schuelerAdressen,dat", "lehrkraefte,dat", "klassen,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd") + "-ImportNachNetman.csv")}[/] erstellt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Schüler*innen, die bereits abgegangen sind oder einen Abschluss erworben haben, werden erst sechs Wochen später ausgebucht, um den Zugriff auf Teams nicht direkt zu verlieren."
                        ],
                        m =>
                        {
                            m.Zieldateien = new Dateien();
                            var zeitstempel = DateTime.Now.ToString("yyyyMMdd-HHmm");
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration)) return;
                            m.Zieldatei = m.WebuntisOderNetmanOderLitteraCsv(configuration, Path.Combine(pfadDownloads ?? "", zeitstempel + @"-ImportNachNetman.csv"), ",", '\'', new UTF8Encoding(false), false);

                            if(configuration["Schulnummer"] != null && configuration["Schulnummer"] == "177659")
                            {
                                configuration = Global.Konfig("ZipKennwort", Global.Modus.Update, configuration, "Zip-Kennwort");
                                m.Zieldatei?.Zippen(Path.Combine(pfadDownloads ?? "", zeitstempel + @"-ImportNachNetman.zip"), configuration, configuration["ZipKennwort"].ToString(), 0, new List<string>(){ Path.Combine(pfadDownloads ?? "", zeitstempel + @"-ImportNachNetman.csv") });
                                m.Zieldatei?.Mailen(Path.Combine(pfadDownloads ?? "", zeitstempel + @"-ImportNachNetman.zip") ?? "", "Verwaltung", Path.GetFileName(m.Zieldatei.AbsoluterPfad) ?? "", configuration);
                            }
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Fotos aus SchILD: Schüler*innen-Fotos aus SchILD für Webuntis und Geevoo bereitstellen",
                        quelldateien.Notwendige(configuration, ["schuelerZusatzdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt die Dateien [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWebuntisFotos.zip")}[/] und [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachGeevooFotos.zip")}[/] erstellt."
                        ],
                        m =>
                        {
                            if(m.NichtAlleSusHabenEineEindeutigeMailAdresse(configuration)) return;
                            m.Zieldatei = new Datei();
                            configuration = Global.Konfig("PfadFotosAusSchILD", Global.Modus.Update, configuration);
                            configuration = Global.Konfig("MailDomain", Global.Modus.Update, configuration);
                            var absoluteFotoPfade = m.GetFotosAusSchildPfade(configuration, m.Students, Global.ZipModus.Webuntis);
                            m.Zieldatei?.Zippen(Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWebuntisFotos.zip"), configuration, "", 0, absoluteFotoPfade);
                            absoluteFotoPfade = m.GetFotosAusSchildPfade(configuration, m.Students, Global.ZipModus.Geevoo);
                            m.Zieldatei?.Zippen(Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachGeevooFotos.zip"), configuration, "", 0, absoluteFotoPfade);
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Mailadressen: fehlende Schulinterne Mailadressen in den Individualdaten I ergänzen",
                        quelldateien.Notwendige(configuration, ["schuelerzusatzdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", "SchuelerZusatzdaten.dat")}[/] um schulinterne Mailadressen ergänzt und in [{Global.GetColor(Global.ColorPfadInDateien)}]{pfadSchilddatenaustausch}[/] für den Re-Import nach SchILD bereitgestellt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1:[/] BKB-Tool bildet die schulinterne Mailadressen wie folgt: [{Global.GetColor(Global.ColorTextHervorheben)}]nv061231@meine-schule.de[/], wobei gilt:",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]            n[/]      : Erster Buchstabe des Nachnamens. Umlaute werden aufgelöst. Bsp.: [{Global.GetColor(Global.ColorTextHervorheben)}]Ü[/] wird zu [{Global.GetColor(Global.ColorTextHervorheben)}]u[/] usw.",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]            v[/]      : Erster Buchstabe des Vornamens. Umlaute werden aufgelöst.",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]            061231[/] : Geburtsdatum in der Notation: JJMMTT.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #2:[/] Vorhandene schulinterne SchILD-Mailadressen in [{Global.GetColor(Global.ColorPfadInProgrammen)}]Individualdaten I[/] bleiben unangetastet.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #3:[/] Doppelungen werden angezeigt und müssen nach Vorgabe behandelt werden."
                        ],
                        m =>
                        {
                            configuration = Global.Konfig("MailDomain", Global.Modus.Read, configuration);

                            m.Zieldateien =
                            [
                                m.SchuelerZusatzdatenUmMailAdresseErgaenzen(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerZusatzdaten.dat"),
                                    ["Nachname", "Vorname", "Geburtsdatum"],
                                    [],
                                    "|", '\'', new UTF8Encoding(false), false)
                            ];                            
                            m.Zieldateien.ExportAusSchildVerschieben(configuration);
                            m.Zieldateien.VergleichenFilternErstellen(quelldateien);
                        },
                        Global.Rubrik.WöchtentlicheArbeiten,
                        Global.NurBeiDiesenSchulnummern.Alle
                    ),
                    new Menüeintrag(
                        "Statistik: Unterrichtsverteilung und Anrechnungen nach SchILD importieren",
                        quelldateien.Notwendige(configuration, ["studentgroupstudents,csv", "klassen,dat", "schuelerlernabschnitt,dat", "schuelerleistungsdaten,dat", "schuelerbasis,dat", "lehrkraefte,dat", "kurse,dat", "lehrkraeftesonderzeiten,dat", "schuelerbasisdaten,dat", "GPU002,txt"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt folgende Dateien für den Import nach SchILD erstellt: \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Lernabschnitte.dat")}[/] \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "Leistungsdaten.dat")}[/] \n[{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat")}[/]",
                            $"Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]LehrkraefteSonderzeiten.dat[/] wird nicht komplett neu erstellt. Die exportierte Datei wird lediglich für den Re-Import aufbereitet.",
                            $"Die Kursbezeichnungen setzen sich zusammen aus dem Kursleiterkürzel plus alle beteiligten Untis-Unterrichtsnummern.",
                            $"Zähler im Anschluss an Fächer (M1, M2, ...) werden abgeschnitten (also zu M).",
                            $"Bei mehreren beteiligten Lehrkräften wird das alphabetisch erste Lehrkraftkürzel zum Kursleiter.",
                            $"Team-Teaching ist daran erkennbar, dass die Summe der Kurs-Wochenstunden kleiner ist als die Summe der Lehrkräfte-Wochenstunden.",
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            configuration = Global.Konfig("Abschnitt", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("Schulnummer", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("StatistikDatum", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("Kursarten", Global.Modus.Read, configuration);

                            m.Unterrichte = new Unterrichte(configuration, m, Global.Zweck.Statistik, Global.Art.Kurse);
                            m.Unterrichte.AddRange(new Unterrichte(configuration, m, Global.Zweck.Statistik, Global.Art.NichtKursUnterrichte));

                            m.Zieldateien =
                            [
                                m.Lernabschnittsdaten(configuration, Global.Zweck.Statistik, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLernabschnittsdaten.dat"), "|", '\0', new UTF8Encoding(true), false),
                                m.Kurse(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Kurse.dat"), "|", '\0', new UTF8Encoding(true), false),
                                m.LeistungsdatenStatistik(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerLeistungsdaten.dat"), Global.Zweck.Statistik),
                                m.LehrkraefteSonderzeiten(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat"),
                                    ["Lehrkraft", "Zeitart", "Grund"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false),
                                m.Lehrkraefte(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Lehrkraefte.dat"),
                                    ["InternKrz"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false),
                                m.Faecher(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Faecher.dat"),
                                    ["InternKrz"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false)
                            ];
                            m.Zieldateien.ExportAusSchildVerschieben(configuration);
                            m.Zieldateien.VergleichenFilternErstellen(quelldateien);
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
                            m.Zieldateien =
                            [
                                m.LehrkraefteSonderzeiten(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "LehrkraefteSonderzeiten.dat"),
                                    ["Lehrkraft", "Zeitart", "Grund"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false, null, "200", Global.Modus.ReadSilent),
                                m.Lehrkraefte(
                                    configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Lehrkraefte.dat"),
                                    ["InternKrz"],
                                    [],
                                    "|", '\0', new UTF8Encoding(true), false),
                            ];
                            m.Zieldateien.ExportAusSchildVerschieben(configuration);
                            m.Zieldateien.VergleichenFilternErstellen(quelldateien);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Fotos erstellen: Schüler*innen klassenweise fotografieren",
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Erstellen Sie jetzt Fotos der vor Ihnen stehenden Klasse (z.B. mit dem Handy). Dabei ist die [{Global.GetColor(Global.ColorInfoBox)}]Reihenfolge & Anzahl[/] laut folgender Tabelle exakt einzuhalten. ",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #1:[/] Wenn jemand fehlt, dann die weiße Wand fotografieren, damit [{Global.GetColor(Global.ColorInfoBox)}]Reihenfolge & Anzahl[/] stimmen.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #2:[/] Wenn ein Foto nicht gelungen ist, dann löschen und neu erstellen.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #3:[/] Wenn mehr als eine Klasse ausgewählt wird, wird nur die erste Klasse berücksichtigt",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis #4:[/] Das Kriterium für die Reihenfolge ist der Dateiname.",
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration, "Klasse", "Geben Sie den Namen der Klasse an, die jetzt vor Ihnen steht.");
                            m.IStudents.KlassenordnerErstellen(configuration);
                            m.IStudents.KlassenListenAnzeigen(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Fotos hochladen: Erstellte Schüler*innenfotos nach SchILD2 hochladen",
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat"]),
                        students,
                        klassen,
                        [
                            $"Es werden jetzt die Fotos nach SchILD2 hochgeladen.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Vorarbeiten:[/] Fotos aller Schüler wurden bereits erstellt und liegen nun in der richtigen [bold aqua]Reihenfolge[/] und [bold aqua]Anzahl[/] in Unterordnern unter [{Global.GetColor(Global.ColorPfadInDateien)}]" + Path.Combine(configuration["PfadDownloads"], "Fotos") +  "[/]."
                        ],
                        m =>
                        {
                            m.IKlassen = m.Students.KlassenAuswählen(configuration);
                            m.Students.FotosVerarbeiten(configuration, m.IKlassen);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                    "Klassen: Neue Klassen von Untis nach SchILD übergeben und Eigenschaften anpassen",
                    quelldateien.Notwendige(configuration, ["klassen,dat", "GPU003,txt"]),
                    students,
                    klassen,
                    [
                        $"Für diese Funktion wird angenommen, dass sich der Klassenname während der gesamten Laufbahn einer/eines jeden Schüler*in nicht ändert. Dazu müssen die Klassennamen so gebildet werden, dass",
                        $"1. das Einschulungsjahr in der Notation JJ im Klassennamen enthalten ist,",
                        $"2. sie auf einen Buchstaben enden. Mit jeder weiteren Parallelklasse erhöht sich der Buchstabe.",
                        $"Die exportierten Klassen aus SchILD werden nun ergänzt und gefiltert, um dann nach SchILD reimportiert zu werden.",
                        $"Bei vorhandenen Klassen werden abweichende Eigenschaften (z.B. Klassenleitung) angepasst.",
                        $"Stellv. Klassenleitung und Prüfungsordnung müssen ggfs. manuell angepasst werden.",
                    ],
                    m =>
                    {
                        m.Zieldateien =
                        [
                            m.KlassenErstellen(
                                configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "Klassen.dat"),
                                ["InternBez"],
                                ["SonstigeBez", "Folgeklasse"],
                                "|", '\0', new UTF8Encoding(true), false),
                        ];
                        m.Zieldateien.ExportAusSchildVerschieben(configuration);
                        m.Zieldateien.VergleichenFilternErstellen(quelldateien);
                    },
                    Global.Rubrik.Allgemein,
                    Global.NurBeiDiesenSchulnummern.Nur177659
                ),
                new Menüeintrag(
                    "Wiki: Diverse SQLite-Dateien (Organigramm, Praktikum etc.) erstellen",
                    quelldateien.Notwendige(configuration, ["schuelerzusatzdaten,dat", "absenceperstudent,csv", "exportlesson,csv", "GPU020,txt"]),
                    students,
                    klassen,
                    [
                        $"Das Organigramm wird aus Untisanrechnungen gebildet. Beispiele: {{...}} > KATEGORIE; [[...]] > HINWEIS, Text ohne Klammern wird zur ROLLE; A14, A15, A16 ohne Klammern > AMT; Untis-Beschreibung > AUFGABE. Im Organigramm wird nach Kategorie, Aufgabe oder Beschreibung gruppiert.",
                        $"Untisanrechnungen: 1.Struct Schema Editor > Untisanrechnungen > Löschen/Leeren > 'untisanrechnungen' eingeben, dann Leeren",
                        $"Untisanrechnungen: 2.Struct Schema Editor > Untisanrechnungen > Importieren/Exportieren > Importieren von Rohdaten > Global > Durchsuchen"
                    ],
                    m =>
                    {
                        var anrechnungen = new Anrechnungen(lehrers, configuration);

                        m.Zieldateien =
                        [
                            m.GetGruppen(configuration, anrechnungen, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-gruppen.csv"), lehrers, ",", '\"', new UTF8Encoding(false), true),
                            m.GetLehrer(configuration, Path.Combine(Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-lul-utf8OhneBom-einmalig-vor-SJ-Beginn.csv")), ",", '\'', new UTF8Encoding(false), false),
                            m.Zieldatei = m.Praktikanten(
                                [
                                    "BW,1", "BT,1", "BS,1", "BS,2", "HBG,1", "HBT,1", "HBW,1", "GG,1", "GT,1", "GW,1", "IFK,1"
                                ],
                                Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-praktikanten-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"), ",", '\'', new UTF8Encoding(false), false),
                            m.KlassenAnlegen(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + @"-klassen-utf8OhneBom-einmalig-vor-SJ-Beginn.csv"), ",", '\"', new UTF8Encoding(false), true)
                        ];

                        m.Schulpflichtüberwachung(configuration);

                        m.Zieldateien.Add(m.GetFaecher(configuration, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-faecher.csv"), ",", '\'', new UTF8Encoding(false), false));

                        m.Zieldateien.VergleichenFilternErstellen(quelldateien);
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
                            foreach (var kalender in new List<string>(){"termine_berufliches_gymnasium", "termine_kollegium", "termine_verwaltung", "termine_fhr" })
                            {
                                m.Zieldateien =
                                [
                                    m.Kalender2Wiki(configuration, kalender, Path.Combine(pfadDownloads ?? "", DateTime.Now.ToString("yyyyMMdd-HHmm") + "-ImportNachWiki-" + kalender), ",", '\"', new UTF8Encoding(false), true)
                                ];

                                m.Zieldateien[0].Erstellen();
                            }
                        },
                        Global.Rubrik.Wiki,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "Schnellmeldung: Relationsgruppen im September aufbereiten",
                        quelldateien,
                        students,
                        klassen,
                        [
                            "Dokumentation siehe Schips.webuntis2schildGui.nrw.de",
                            "Realtionen gemäß §93 SchulG"
                        ],
                        _ => { new Relationsgruppen(klassen, students); },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659                        
                    ),                
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
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Falls mehrere Kollegen dasselbe Fach zeitgleich unterrichten, dann muss ein Zähler an das Fach angehangen werden. Bsp: Zwei LuL unterrichten Mathe: Dann M und M1. Beide Fächer müssen in SchILD existieren. Damit M1 in den Leistungsdaten erscheint, aber nicht auf dem Zeugnis gedruckt wird, muss die Eigenschaft 'Nicht auf Zeugnis drucken' in SchILD gesetzt werden.",                            
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
                        "Zeugnisse #2: Säumige Lehrer*innen im Teams-Chat an die Noten-Eintragung erinnern",
                        anrechnungen,
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
                        "Teams-Chat: Teams-Chat mit Gruppe von Lehrkräften beginnen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["exportlessons,csv"]),
                        students,
                        klassen,
                        [
                            "Lehrkräfte können über Teams angeschrieben werden.",
                            $"Dazu wird die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadDownloads ?? "", "ExportLessons.csv")}[/] ausgewertet.",
                            "Es wird jetzt ein Link nach Teams erstellt, um einen Chat mit den Lehrkräften zu beginnen."
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

                            configuration = Global.Konfig("TeamsChatAuswahl", Global.Modus.Update, configuration);

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
                        Global.NurBeiDiesenSchulnummern.Nur177659
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
                        "Teilleistungen: SchuelerTeilleistungen.dat für SchILD erstellen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["schuelerbasisdaten,dat", "exportlessons,csv", "marksperlesson,csv"]),
                        students,
                        klassen,
                        [
                            $"Es wird jetzt die Datei [bold {Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerTeilleistungen.dat")}[/] erstellt.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Damit der Import nach SchILD reibungslos funktioniert, müssen zuvor die Teilleistungsarten in SchILD ([{Global.GetColor(Global.ColorPfadInProgrammen)}]Schulverwaltung > Teilleistungsarten[/]) gleichlautend mit dem Langnamen in Webuntis ([{Global.GetColor(Global.ColorPfadInProgrammen)}]Stammdaten > Prüfungsarten[/]) angelegt werden.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Es empfiehlt sich, dass die Lernabschnitssdaten und Leistungsdaten zuerst in SchILD importiert bzw. angelegt werden."
                        ],
                        m =>
                        {
                            m.FilterInteressierendeStudentsUndKlassen(configuration);
                            m.Zieldatei = m.Teilleistungen(configuration, Path.Combine(pfadSchilddatenaustausch ?? "", "SchuelerTeilleistungen.dat"));
                            //m.Zieldatei = m.Zieldatei.VergleichenUndFiltern(quelldateien, configuration, ["Nachname", "Vorname", "Geburtsdatum", "Jahr", "Abschnitt"], []);
                            m.Zieldatei.Erstellen("|", '\0', new UTF8Encoding(true), false);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Alle
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
                            "Mit weniger als 10 offenen Eintragungen wird nicht gemahnt. Ab 20 oder mehr Stunden wird die Schulleitung in CC informiert.",
                            $"Die Anzahl der offenen Klassenbucheinträge wird aus der Datei [{Global.GetColor(Global.ColorPfadInDateien)}]OpenPeriods[/] ausgelesen.",
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
                        $"PDF-Dateien #1: Von PDF-Dateien in {configuration["PfadDownloads"]} verschlüsselte Kopien erstellen",
                        anrechnungen,
                        new Dateien(),
                        students,
                        klassen,
                        [
                            $"Von allen PDF-Dateien in [{Global.GetColor(Global.ColorPfadInDateien)}]" + configuration["PfadDownloads"] + "[/] wird eine verschlüsselte Kopie erstellt.",
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
                        "PDF-Dateien #2: PDF-Seiten an darauf enthaltene E-Mail-Adressen mailen",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, ["lehrkraefte,dat"]),
                        students,
                        klassen,
                        [
                            "1. Die zuletzt bearbeitete PDF-Datei wird eingelesen.",
                            "2. Jede Seite der Datei wird nach E-Mail-Adressen durchsucht.",
                            "3. Die betreffenden Seiten werden an die E-Mail-Adressen gemailt.",
                            "4. Optional wird verschlüsselt."
                        ],
                        m =>
                        {
                            var pdfDatei = Directory.GetFiles(pfadDownloads, "*.pdf").OrderByDescending(File.GetLastWriteTime).FirstOrDefault();
                            Global.ZeileSchreiben("Die neueste PDF-Datei wird versendet:", pdfDatei, ConsoleColor.White, ConsoleColor.Black);
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
                                    seite?.PdfDocumentEncrypt(pdfKennwort);
                                    seite?.Mailen(configuration);
                                }
                            });
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
                    ),
                    new Menüeintrag(
                        "PDF-Zeugnisse: Von Atlantis in die SchILD-Dokumentenverwaltung kopieren",
                        anrechnungen,
                        quelldateien.Notwendige(configuration, []),
                        students,
                        klassen,
                        [
                            $"PDF-Zeugnisse und andere PDF-Dateien werden in die Schüler*innen-Ordner der SchILD-Dokumentenverwaltung einsortiert.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung #1[/]: Zu kopierende PDF-Dateien nach [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(configuration["PfadDownloads"], "PDF-Input")}[/] kopieren.",
                            $"[{Global.GetColor(Global.ColorHinweise)}]Vorbereitung #2[/]: Eine UTF8-CSV-Datei mit Spalten: Nachname, Vorname, Geburtsdatum und Klasse aus Atlantis exportieren und in [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.Combine(configuration["PfadDownloads"], "PDF-Input")}[/] ablegen.",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]Durchführung #1[/]: (Einzelne) Klasse(n) oder 'alle' auswählen.",
                            $"[{Global.GetColor(Global.ColorTextHervorheben)}]Durchführung #2[/]: Geben Sie die Schlüsselwörter an, um die interessierenden PDF-Dateien einzugrenzen.",
                        ],
                        m =>
                        {
                            configuration = Global.Konfig("PfadDownloads", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("PfadDokumentenverwaltung", Global.Modus.Read, configuration);
                            configuration = Global.Konfig("Schlüsselwörter", Global.Modus.Update, configuration);

                            m.IStudents.GetStudentsVonAtlantisCsv(configuration);
                            m.IStudents.PdfDateienVerarbeiten(configuration);
                        },
                        Global.Rubrik.Allgemein,
                        Global.NurBeiDiesenSchulnummern.Nur177659
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
                    )*/
                ]
            );
        }
        catch (Exception ex)
        {
            throw ex;
        }   
    }
}