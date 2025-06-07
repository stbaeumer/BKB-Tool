// Stefan Bäumer 2025
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Common;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using Path = System.IO.Path;
using dotenv.net;


public static class Global
{
    public static List<(string Meldung, ConsoleColor Farbe)> Zeilen = new();
    public static int Kalenderwoche { get; set; }    
    public static string? WikiJsonUserKennwort { get; set; }
    public static string? WikiJsonUser { get; set; }
    public static string? WikiUrl { get; set; }    
    public static string? SchipsUrl { get; set; }    
    public static string? NetmanMailReceiver { get; set; }
    public static string? SmtpServer { get; set; }
    public static string? SmtpPort { get; set; }
    public static string? SmtpPassword { get; set; }
    public static string? SmtpUser { get; set; }
    public static DateTime Sprechtagsdatum { get; set; }    
    public static string? AktuellerPfad { get; set; }    
    public static string? OutputFolder { get; set; }    
    public static string? User { get; set; }
    public static int FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt { get; set; }
    public static object? WikiSprechtagKleineAenderung { get; set; }
    public static List<string>? Protokoll { get; set; }    

    public enum Art
    {
        Mahnung,
        Statistik,
        Zeugnis
    }

    public enum NurBeiDiesenSchulnummern
    {
        Alle,
        NurPrivilegiert
    }

    public enum Rubrik
    {
        Allgemein,
        Leistungsdaten,
        WöchtentlicheArbeiten,
        Wiki
    }

    public enum Datentyp
    {
        String,
        Int,
        DateTime,
        Pfad,
        Url,
        JaNein,
        Mail,
        Klassen,
        Datei,
        Abschnitt,
        ListInt,
        Float,
        Auswahl
    }

    public enum Modus
    {
        Create,
        Update,
        Read
    }

    public static List<string> AktSj = new List<string>()
    {
        (DateTime.Now.Month >= 7 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
        (DateTime.Now.Month >= 7 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString()
    };

    public static string? ConnectionStringUntis { get; set; }
    public static string? ZipKennwort { get; set; }  
    public static string? PfadSchilddatenaustausch { get; private set; }
    public static List<string>? PrivilegierteSchulnummern { get; set; }
    public static string AppVersion { get; set; }

    public static string? SafeGetString(SqlDataReader reader, int colIndex)
    {
        if (!reader.IsDBNull(colIndex))
            return reader.GetString(colIndex);
        return string.Empty;
    }

    internal static void Dateischreiben(IConfiguration configuration, string name)
    {
        string tempPfadUndDatei = Path.Combine(Path.GetTempPath(), Path.GetFileName(name));
        string pfadDownloads = configuration["pfadDownloads"] ?? throw new ArgumentNullException(nameof(configuration), "pfadDownloads cannot be null");
        string pfadUndDatei = Path.Combine(pfadDownloads, name);
        UTF8Encoding utf8NoBom = new UTF8Encoding(false);

        if (File.Exists(pfadUndDatei) && File.Exists(tempPfadUndDatei))
        {
            string contentNeu = File.ReadAllText(tempPfadUndDatei, utf8NoBom);

            // Lese den Inhalt der Dateien
            string contentAlt = File.ReadAllText(pfadUndDatei, utf8NoBom);

            // Vergleiche die Inhalte der Dateien
            if (contentAlt != contentNeu)
            {
                // Überschreibe alt mit dem Inhalt von neu
                File.WriteAllText(pfadUndDatei, contentNeu, utf8NoBom);
                ZeileSchreiben(name, "überschrieben", ConsoleColor.Yellow, ConsoleColor.Gray);
            }
            else
            {
                ZeileSchreiben(name, "Identisch. Keine Änderungen", ConsoleColor.White, ConsoleColor.Black);
            }
        }

        if (!File.Exists(pfadUndDatei))
        {
            string directoryPath = Path.GetDirectoryName(pfadUndDatei) ?? string.Empty;

            if (directoryPath != null)
            {
                // Fehlende Verzeichnisse anlegen
                Directory.CreateDirectory(directoryPath);
            }

            string contentNeu = File.ReadAllText(tempPfadUndDatei, utf8NoBom);
            File.WriteAllText(pfadUndDatei, contentNeu, utf8NoBom);
            ZeileSchreiben(name, ": Datei neu erstellt.", ConsoleColor.Green, ConsoleColor.Gray);
        }
    }

    public static void DisplayHeader(IConfiguration configuration)
    {
        Console.Clear();
        AnsiConsole.Write(new FigletText(configuration["AppName"] ?? "AppName").Centered().Color(Color.Aquamarine1_1));
        AnsiConsole.Write(new Rule($"[aquamarine1_1] {AppVersion} | [link=https://github.com/stbaeumer/BKB-Tool]https://github.com/stbaeumer/BKB-Tool[/] | GPLv3 [/]").RuleStyle("seagreen1_1").Centered());
    }

    public static string InsertLineBreaks(string text, int maxLineLength)
    {
        if (string.IsNullOrEmpty(text) || maxLineLength <= 0)
            return text;

        var currentIndex = 0;
        var length = text.Length;
        var result = new StringBuilder();

        while (currentIndex < length)
        {
            // Calculate the length of the next segment
            int nextSegmentLength = Math.Min(maxLineLength, length - currentIndex);
            // Append the segment and a line break
            result.Append(text.Substring(currentIndex, nextSegmentLength));
            result.Append(Environment.NewLine + "   ");
            // Move to the next segment
            currentIndex += nextSegmentLength;
        }

        return result.ToString();
    }

    internal static void OrdnerAnlegen(object name)
    {
        throw new NotImplementedException();
    }

    public static void ZeileSchreiben(string linkeSeite, string rechteSeite, ConsoleColor foreground = ConsoleColor.Black, ConsoleColor background = ConsoleColor.White)
    {
        var gesamtbreite = Console.WindowWidth;
        var punkte = gesamtbreite - linkeSeite.Length - rechteSeite.Length - 1;
        var mitte = " .".PadRight(Math.Max(3, punkte), '.') + " ";

        // Wenn linkeSeite auf einen Punkt endet, dann wird das Leerzeichen links durch einen Punkt ersetzt

        if (linkeSeite.Length > 1 && linkeSeite.Substring(linkeSeite.Length - 1, 1) == ".")
        {
            mitte = "." + mitte.Substring(1);
        }

        // Wenn die linke Seite ein Pfad ist, dann wird ein Panel erstellt.

        if (File.Exists(linkeSeite))
        {
            var path = new TextPath(linkeSeite);

            path.RootStyle = new Style(foreground: Color.Red);
            path.SeparatorStyle = new Style(foreground: Color.Green);
            path.StemStyle = new Style(foreground: Color.Blue);
            path.LeafStyle = new Style(foreground: Color.Yellow);

            var panel = new Panel(path)
                .Header("[bold blue]  Neu:  [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Aquamarine1_1);

            AnsiConsole.Write(panel);
        }
        else
        {
            AnsiConsole.MarkupLine("[blue]" + linkeSeite + "[/]" + "[blue]" + mitte + "[/]" + "[blue]" + rechteSeite + "[/]");
        }
    }

    public static void Speichern(string key, string value)
    {
        var json = File.ReadAllText(Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json"));
        var jsonDoc = JsonDocument.Parse(json);
        var jsonRoot = jsonDoc.RootElement;

        string finalValue = Verschluesseln(value);

        // Neuen Wert setzen
        using (var stream = new MemoryStream())
        {
            using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true }))
            {
                writer.WriteStartObject();
                foreach (var property in jsonRoot.EnumerateObject())
                {
                    if (property.NameEquals(key))
                    {
                        writer.WriteString(key, finalValue);
                    }
                    else
                    {
                        property.WriteTo(writer);
                    }
                }

                writer.WriteEndObject();
            }

            // Neue JSON-Daten in die Datei schreiben
            File.WriteAllText(Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json"), Encoding.UTF8.GetString(stream.ToArray()));
        }
    }

    // Hilfsmethode zur Verschlüsselung
    public static string Verschluesseln(string value)
    {
        // Beispiel für eine einfache Verschlüsselung (Base64)
        byte[] data = Encoding.UTF8.GetBytes(value);
        string encryptedValue = Convert.ToBase64String(data);
        return encryptedValue;
    }
    
    public static void OpenWebseite(string url)
    {
        if (!url.StartsWith("http"))
        {
            url = "https://wiki.svws.nrw.de/mediawiki/index.php?title=Schnittstellenbeschreibung#";
            Console.WriteLine("     Die Seite " + url + " wird geöffnet.");
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url.TrimEnd('#'),
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine("    Fehler beim Öffnen der Webseite:");
            Console.WriteLine(ex.Message);
        }
    }

    public static string CheckFile(IConfiguration configuration, string Pfad, string endung)
    {
        var pfadDownloads = configuration["pfadDownloads"] ?? throw new ArgumentNullException(nameof(configuration), "pfadDownloads cannot be null");
        Pfad = Path.Combine(pfadDownloads, Pfad);

        if (!Path.Exists(Path.GetDirectoryName(Pfad)))
        {
            var directoryPath = Path.GetDirectoryName(Pfad);
            if (!string.IsNullOrEmpty(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
        }
        
        var directoryName = Path.GetDirectoryName(Pfad);
        string sourceFile = string.Empty;

        if (!string.IsNullOrEmpty(directoryName))
        {
            sourceFile =
                (from f in Directory.GetFiles(directoryName, endung, SearchOption.AllDirectories)
                 where Path.GetFileName(f).StartsWith(Path.GetFileName(Pfad))
                 orderby File.GetLastWriteTime(f)
                 select f).LastOrDefault();
        }

        return sourceFile ?? string.Empty;
    }

    public static IConfiguration Konfig(string parameter, Modus modus, IConfiguration configuration, string aufforderung = "", string hinweise = "", Datentyp datentyp = Datentyp.String, string defaultValue = "alle", Students? students = null, string zulässigeAuswahlOptionen = "")
    {
        object userInput = "";

        var panel = new Panel(hinweise)
            //.Header("[bold blue]  Hinweis:  [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand()
            .BorderColor(Color.Green);

        // Der Wert aus der JSON hat Vorrang vor dem defaultwert. Nur wenn die JSON keinen Wert enthält, wird der defaultwert verwendet.
        defaultValue = !string.IsNullOrEmpty(configuration[parameter])
            ? configuration[parameter] ?? defaultValue
            : defaultValue;

        if (datentyp == Datentyp.JaNein)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && !string.IsNullOrEmpty(defaultValue) && defaultValue.ToLower().StartsWith("j"))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
            new TextPrompt<string>(aufforderung)
                .PromptStyle("green")
                .ShowDefaultValue(true)
                .Validate(n =>
                {
                    if (zulässigeAuswahlOptionen.ToLower() != "ja")
                        return ValidationResult.Error($"Sie müssen [bold green]Ja[/] eintippen, um {configuration["AppName"]} nutzen zu können.");
                    return ValidationResult.Success();
                }));


            if (userInput.ToString()?.ToLower() == "ja" || userInput.ToString()?.ToLower() == "j")
                DisplayHeader(configuration);
        }
        if (datentyp == Datentyp.Auswahl)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && !string.IsNullOrEmpty(defaultValue))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            var größteAuswahlZahl = GetMaxNumberFromList(zulässigeAuswahlOptionen.Split(','));

            userInput = AnsiConsole.Prompt(
            new TextPrompt<string>(aufforderung)
                .PromptStyle("green")
                .DefaultValue(defaultValue.ToLower() == "x" ? "x" : defaultValue.ToString())
                .ShowDefaultValue(true)
                .Validate(n =>
                {
                    if (!zulässigeAuswahlOptionen.Split(",").Contains(n))
                        return ValidationResult.Error($"Zulässige Auswahl: [bold blue]{zulässigeAuswahlOptionen}[/]");
                    return ValidationResult.Success();
                }));
        }
        if (datentyp == Datentyp.String)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && !string.IsNullOrEmpty(defaultValue))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (string.IsNullOrEmpty(n))
                            return ValidationResult.Error("Eingabe darf nicht leer sein.");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Datei)
        {
            defaultValue = File.Exists(defaultValue) ? defaultValue : string.Empty;

            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && !string.IsNullOrEmpty(defaultValue) && File.Exists(defaultValue))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (string.IsNullOrEmpty(n))
                            return ValidationResult.Error("Eingabe darf nicht leer sein.");
                        if (!File.Exists(n))
                            return ValidationResult.Error($"Die Datei [bold red]{n}[/] existiert nicht.");
                        return ValidationResult.Success();
                    })
                    .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Klassen)
        {
            var verschiedeneKlassen = students.Select(s => s.Klasse).Distinct().ToList();
            var interessierendeKlassen = new List<string>();

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!string.IsNullOrEmpty(n) && n == "alle")
                            return ValidationResult.Success();
                        if (n.Split(',').Any(teil => verschiedeneKlassen.Any(klasse => klasse.ToLower().Contains(teil.Trim().ToLower()))))
                            return ValidationResult.Success();
                        if (verschiedeneKlassen.Any(s => s.ToLower().StartsWith(n.ToLower())))
                            return ValidationResult.Success();
                        return ValidationResult.Error("Eingabe ist ungültig. Geben Sie eine Klasse an oder 'alle'.");
                    })
                .DefaultValue<string>(defaultValue));

            var x = userInput.ToString() ?? string.Empty;

            foreach (var klasse in verschiedeneKlassen)
            {
                foreach (var item in x.Trim().Split(','))
                {
                    if (klasse.ToLower().StartsWith(item.ToLower()))
                    {
                        if (!interessierendeKlassen.Select(s => s.ToLower()).Contains(klasse.ToLower()))
                        {
                            if (klasse != "")
                            {
                                interessierendeKlassen.Add(klasse);
                            }
                        }
                    }
                }
            }
            userInput = string.Join(",", interessierendeKlassen);
        }
        if (datentyp == Datentyp.Url)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && (!string.IsNullOrEmpty(defaultValue) || defaultValue.StartsWith("https://")))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!n.StartsWith("https://") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("Eingabe muss mit https:// beginnen.");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Mail)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && (!string.IsNullOrEmpty(defaultValue) && defaultValue.StartsWith("@") && defaultValue.Contains(".")))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!n.StartsWith("@") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("Eingabe muss mit @ beginnen und einen Punkt enthalten.");
                        if (!n.Contains(".") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("Eingabe muss mit @ beginnen und einen Punkt enthalten.");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Pfad)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && Path.Exists(defaultValue))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }
            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!Path.Exists(n))
                            return ValidationResult.Error($"Der Pfad {n} existiert nicht.");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(string.IsNullOrEmpty(defaultValue) || !Path.Exists(defaultValue) ? Environment.CurrentDirectory : defaultValue));
        }
        if (datentyp == Datentyp.Int)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && int.TryParse(defaultValue, out _))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!int.TryParse(n.ToString(), out _))
                        {
                            return ValidationResult.Error($"{n} ist keine Zahl");
                        }
                        if (zulässigeAuswahlOptionen != "" && !zulässigeAuswahlOptionen.Contains(n.ToString()))
                        {
                            return ValidationResult.Error($"Die Zahl {n} außerhalb des zulässigen Bereichs. Zulässige Werte: {zulässigeAuswahlOptionen}");
                        }
                                
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.Float)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && float.TryParse(defaultValue, out _))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!float.TryParse(n.ToString(), out _))
                            return ValidationResult.Error($"{n} ist keine Zahl");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.ListInt)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && int.TryParse(defaultValue, out _))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        var teile = n.ToString().Split(',');
                        if (!teile.All(t => int.TryParse(t.Trim(), out _)))
                            return ValidationResult.Error($"{n} ist keine Liste aus Zahlen");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.Abschnitt)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && int.TryParse(defaultValue, out _))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!int.TryParse(n.ToString(), out _))
                        {
                            return ValidationResult.Error($"{n} ist keine Zahl");
                        }
                        else
                        {
                            if (n != "1" && n != "2")
                                return ValidationResult.Error($"{n} ist ungültig. Erlaubt sind 1 und 2.");
                        }

                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.DateTime)
        {
            // Wenn im READ-Modus der Wert plausibel ist, dann wird er nicht erneut abgefragt
            if (modus == Modus.Read && DateTime.TryParse(defaultValue, out _))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>(aufforderung)
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if (!DateTime.TryParse(n.ToString(), out _))
                        {
                            return ValidationResult.Error($"{n} ist kein Datum (TT.MM.JJJJ)");
                        }

                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }

        // Speichern des Klartextwerts in der Konfiguration
        configuration[parameter] = userInput.ToString();

        // Verschlüsselten Wert in der JSON-Datei speichern
        Speichern(parameter, userInput?.ToString() ?? string.Empty);

        return configuration;
    }

    public static void EditorOeffnen(string pfad)
    {
        try
        {
            System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Notepad++\Notepad++.exe", pfad);
        }
        catch (Exception)
        {
            System.Diagnostics.Process.Start("Notepad.exe", pfad);
        }
    }

    public static string PrüfeAufNullOrEmpty(IDictionary<string, object> dict, string s)
    {
        if (dict.TryGetValue(s, out var nameObj) && nameObj is string name && !string.IsNullOrWhiteSpace(name))
        {
            return name;
        }
        else
        {
            return "";
        }
    }
    
    public static IConfiguration EinstellungenDurchlaufen(Modus modus, IConfiguration configuration)
    {
        // Wenn User.json noch nicht existiert, dann erstellen
        if (!File.Exists(Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json")))
        {
            // User.json mit Standardinhalten füllen
            var bkbJsonContent = CreateBkbJsonContent();
            var json = JsonSerializer.Serialize(bkbJsonContent, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json"), json);
        }

        // Konfiguration aus User.json laden
        configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(Global.User + ".json", optional: false, reloadOnChange: false)
            .Build();

        // Alle Values entschlüsseln
        foreach (var key in configuration.AsEnumerable())
        {
            if (key.Value != null)
            {
                configuration[key.Key] = Entschluesseln(key.Value);
            }
        }

        while (string.IsNullOrEmpty(configuration["ZustimmungLizenz"]) || 
               (configuration["ZustimmungLizenz"]?.ToLower() != "ja" && configuration["ZustimmungLizenz"]?.ToLower() != "j"))
        {
            configuration = Global.Konfig("ZustimmungLizenz", Global.Modus.Update, configuration, "Ich stimme den Lizenzbedingungen der GPLv3 zu. (Ja/Nein)", "Sie müssen den Lizenzbedingungen der GPLv3 zustimmen.", Global.Datentyp.JaNein, "", null, "Ja");
        }

        if (modus == Modus.Update)
            DisplayHeader(configuration);
        
        if (modus == Modus.Create)
            DisplayHeader(configuration);
        var panel = new Panel("Ihre Einstellungen werden verschlüsselt in der Datei [blue]" + Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json[/]") + " gespeichert." +
        $"\nDateien (aus Webuntis etc.), die {configuration["AppName"]} importieren soll, werden aus [blue]" + configuration["PfadDownloads"] + "[/] eingelesen.")
                        .Header($" [bold blue]  Einstellungen  [/]")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Blue);

        if (modus != Modus.Read)
            AnsiConsole.Write(panel);

        configuration = Konfig("PfadDownloads", modus, configuration, @"Downloads-Verzeichnis", "Geben Sie im Folgenden den Pfad des Downloads-Verzeichnisses an. In der Regel wird das Verzeichnis bereits richtig vorgeschlagen. Dann einfach [bold green]ENTER[/] drücken:", Datentyp.Pfad, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"));
        configuration = Konfig("Schulnummer", modus, configuration, @"Schulnummer", "Geben Sie Im Folgenden Ihre Schulnummer an. Je nach Schulnummer werden evtl. unterschiedliche Funktionen angeboten.", Datentyp.Int);
        configuration = Konfig("PfadSchilddatenaustausch", modus, configuration, @"SchILD-Datenaustausch-Ausgabeverzeichnis", "Geben Sie im Folgenden das Verzeichnis an, das in SchILD unter [green bold]Datenaustausch > Schnittstelle SchILD-NRW > Export[/] als [green bold]Ausgabeverzeichnis[/] eingetragen ist. Wenn dort kein Verzeichnis steht, tragen Sie dort das selbe Verzeichnis ein, das Sie auch im Folgenden angeben:", Datentyp.Pfad, "/home/stefan/Windows/SchILD-NRW/Datenaustausch");
        configuration = Konfig("MaxDateiAlter", modus, configuration, "Wie viele Tage dürfen Dateien höchstens alt sein?", $"Geben Sie im Folgenden an, wie viele Tage Dateien höchstens alt sein dürfen, um vom {configuration["AppName"]} für das Einlesen akzeptiert zu werden. Die Angabe einer (möglichst niedrigen) Zahl soll sicherstellen, dass nicht versehntlich veraltete Dateien eingelesen werden.", Datentyp.Int);
        configuration = Konfig("AppName", modus, configuration, "Wie soll die App heißen?", $"Sie können die App [bold green]{configuration["AppName"]}[/] umbennen.", Datentyp.String);

        if (modus == Modus.Update && PrivilegierteSchulnummern.Contains(configuration["Schulnummer"]))
        {
            configuration = Konfig("MailDomain", modus, configuration, "Mail-Domain für Schüler*innen", "Geben Sie im Folgenden die Mail-Domain für Ihre Schüler*innen an. Ihre Eingabe muss mit [green bold]@[/] beginnen und einen [green bold]Punkt[/] enthalten. Beispiel: [green bold]@students.meine-schule.de[/]", Datentyp.Mail);
            configuration = Konfig("ConnectionStringUntis", modus, configuration, "ConnectionStringUntis (optional)");
            configuration = Konfig("ZipKennwort", modus, configuration, "Kennwort zum Verschlüsseln von Zip-Dateien");
            configuration = Konfig("SmtpUser", modus, configuration, "Mail-Benutzer");
            configuration = Konfig("SmtpPassword", modus, configuration, "Mail-Kennwort");
            configuration = Konfig("SmtpPort", modus, configuration, "SMTP-Port");
            configuration = Konfig("SmtpServer", modus, configuration, "SMTP-Server");
            configuration = Konfig("NetmanMailReceiver", modus, configuration, "Wem soll die Netman-Mail geschickt werden?");
        }

        return configuration;
    }

    static object CreateBkbJsonContent()
    {
        DotEnv.Load();

        return new
        {
            ConnectionStringUntis = Verschluesseln(Environment.GetEnvironmentVariable("UNTIS_CONNECTION_STRING") ?? ""),
            Schulnummer = Verschluesseln("000000" ?? ""),
            SchipsPasswort = Verschluesseln(Environment.GetEnvironmentVariable("SCHIPS_PASSWORD") ?? ""),
            ZeugnisPasswort = Verschluesseln(Environment.GetEnvironmentVariable("ZEUGNIS_PASSWORD") ?? ""),
            SmtpPassword = Verschluesseln(Environment.GetEnvironmentVariable("SMTP_PASSWORD") ?? ""),
            AccessPassword = Verschluesseln(Environment.GetEnvironmentVariable("ACCESS_PASSWORD") ?? ""),
            SmtpPasswordMassenmail = Verschluesseln(Environment.GetEnvironmentVariable("SMTP_PASSWORD_MASSENMAIL") ?? ""),
            ZeugnisUrl = Verschluesseln("https://bkb.wiki/notenlisten:start"),
            SchipsUrl = Verschluesseln("https://bkb.wiki/statistik:schips:start"),
            PdfKennwort = Verschluesseln(Environment.GetEnvironmentVariable("PDF_PASSWORD") ?? ""),
            SchipsOderZeugnisseOderAnderePdfs = Verschluesseln("1"),
            PfadDownloads = Verschluesseln(Environment.GetEnvironmentVariable("DOWNLOADS_PATH") ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")),
            PfadSchilddatenaustausch,
            Kalenderfilter = Verschluesseln(""),
            Auswahl = "x",
            OnlineHilfeURL = Verschluesseln("https://github.com/stbaeumer/BKB-Tool/wiki"),
            ZustimmungLizenz = "nein",
            AppName = Verschluesseln("BKB-Tool"),
            AppVersion = Verschluesseln("1.0.0"),
            AppDescription = Verschluesseln("BKB-Tool - Ein Werkzeug an der Schnittstelle zwischen SchILD und Untis."),
            Klassen = Verschluesseln("HBG"),
            Vergleich = Verschluesseln("n"),
            Kennwort = Verschluesseln(""),
            InputFolder = Verschluesseln(""),
            OutputFolder = Verschluesseln(""),
            Halbjahreszeugnisdatum = Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
            Halbjahreskonferenzdatum = Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
            Jahreszeugnisdatum = Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
            Jahreskonferenzdatum = Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
            Sprechtagsdatum = Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
            WikiUrl = Verschluesseln("http://192.168.134.10/lib/exe/xmlrpc.php"),
            WikiJsonUser = Verschluesseln("root"),
            WikiJsonUserKennwort = Verschluesseln(""),
            WikiSprechtagKleineAenderung = Verschluesseln(""),
            Zaehlerfach = Verschluesseln("j"),
            MariaUser = Verschluesseln(""),
            MariaIp = Verschluesseln(""),
            MariaPort = Verschluesseln(""),
            MariaDb = Verschluesseln(""),
            MariaPw = Verschluesseln(""),
            FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt = Verschluesseln(""),
            MaximaleAnzahlFehlstundenProTag = Verschluesseln("8"),
            Abschnitt = Verschluesseln("1"),
            Chat = Verschluesseln(""),
            AusbuchenNachWievielTagen = Verschluesseln(""),
            DatenimportLetztesDatum = Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
            MaxDateiAlter = Verschluesseln("6"),
            AktSj = Verschluesseln(""),
            Klasse = Verschluesseln(""),
            MailDomain = Verschluesseln("@students.berufskolleg-borken.de"),
            ZipKennwort = Verschluesseln("Geheim123"),
            EinstellungenVorgenommen = Verschluesseln("n"),
            SmtpServer = Verschluesseln("smtp.office365.com"),
            SmtpUser = Verschluesseln("webuntis@berufskolleg-borken.de"),
            SmtpPort = Verschluesseln("587"),
            NetmanMailReceiver = Verschluesseln("stefan.baeumer@berufskolleg-borken.de"),
            NetmanMailBccReceiver = Verschluesseln("catrin.stakenkoetter@berufskolleg-borken.de"),
            Betreff = Verschluesseln("Betreff"),
            Body = Verschluesseln("Guten Morgen [Lehrer],\n\nbitte beachten Sie den Anhang.\n\nErläuterungen dazu finden Sie hier: https://bkb.wiki/konzepte:stundenplanungskonzept#information_aller_lehrkraefte_per_mail \n\nViele Grüße aus der Schulverwaltung"),
            AccessPfad = Verschluesseln(@"\\fs01\SchILD-NRW\DB\Test.mdb"),
            PdfInputFolder = Verschluesseln(@"PDF-Input"),
            PdfOutputFolder = Verschluesseln(@"PDF-Output"),
            PfadDokumentenverwaltung = Verschluesseln(@"/home/stefan/Windows/Dokumentenverwaltung"),
            BodyMassenmail = Verschluesseln("Guten Morgen [Lehrer],\n\nbitte beachten Sie den Anhang.\n\nErläuterungen dazu finden Sie hier: https://bkb.wiki/konzepte:stundenplanungskonzept#information_aller_lehrkraefte_per_mail \n\nViele Grüße aus der Schulverwaltung"),
            Verbose = Verschluesseln("false"),
            SmtpUserMassenmail = Verschluesseln("campusfest@berufskolleg-borken.de"),
            SmtpServerMassenmail = Verschluesseln("smtp-hve.office365.com"),
            BetreffMassenmail = Verschluesseln("Save the date: Campusfest am Berufskolleg Borken"),
            Schlüsselwörter = Verschluesseln("Jahreszeugnis, Abschlusszeugnis, Abgangszeugnis, Zeugnis"),
            Teilleistungsarten = Verschluesseln("Vornote"),
            LehrkraefteSonderzeiten = Verschluesseln(""),
            VolleStelle = Verschluesseln("25,5")
        };
    }

    /// <summary>
    /// Prüft, ob die Anwendung in einem Codespace läuft.
    /// </summary>
    /// <returns></returns>
    public static bool RunningInCodeSpace()
    {
        if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("CODESPACES")))
        {
            return true;
        }
        return false;
    }

    public static void WeiterMitAnykey(IConfiguration configuration, Menüeintrag menüeintrag = null)
    {
        if (!RunningInCodeSpace())
        {
            if (menüeintrag != null && menüeintrag.Zieldatei != null && menüeintrag.Zieldatei.Ordner != null)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = menüeintrag.Zieldatei.Ordner,
                    UseShellExecute = true
                });
            }
            else
            {
                if (menüeintrag != null && menüeintrag.Zieldatei != null && !string.IsNullOrEmpty(Path.GetDirectoryName(menüeintrag.Zieldatei.AbsoluterPfad)))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = Path.GetDirectoryName(menüeintrag.Zieldatei.AbsoluterPfad),
                        UseShellExecute = true
                    });
                }
            }
        }

        var panel = new Panel($"Weiter mit [bold green]Anykey[/] oder mit [bold blue]x[/] Einstellungen durchlaufen oder mit [bold blue]y[/] Onlinehilfe öffnen.")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Grey);

        AnsiConsole.Write(panel);

        var weiter = Console.ReadKey(true); // true unterdrückt die Ausgabe des Zeichens im Terminal

        if (weiter.Key == ConsoleKey.Y)
        {
            configuration = EinstellungenDurchlaufen(Modus.Update, configuration);
            return;
        }

        if (weiter.Key == ConsoleKey.H)
        {
            OpenWebseite(configuration["OnlineHilfeURL"]);
            return;
        }

        DisplayHeader(configuration);
    }
    public static string Entschluesseln(string encryptedValue)
    {
        // Beispiel für eine einfache Entschlüsselung (Base64)
        try
        {
            byte[] data = Convert.FromBase64String(encryptedValue);
            return Encoding.UTF8.GetString(data);
        }
        catch
        {
            // Falls der Wert nicht entschlüsselt werden kann, wird er unverändert zurückgegeben
            return encryptedValue;
        }
    }

    internal static DataAccess DataAccessHerstellen(IConfiguration configuration)
    {
        do
        {
            configuration = Global.Konfig("AccessPfad", Global.Modus.Read, configuration, "Pfad zur Access-Datenbank", "", Global.Datentyp.Datei);
            configuration = Global.Konfig("AccessPassword", Global.Modus.Read, configuration, "Passwort zur Access-Datenbank");

            DateTime[] releases =
            {
                new DateTime(2017, 2, 11),
                new DateTime(2017, 5, 11),
                new DateTime(2017, 11, 14),
                new DateTime(2017, 10, 14),
                new DateTime(2018, 02, 02),
                new DateTime(2018, 06, 10),
                new DateTime(2025, 01, 18),
                new DateTime(2025, 01, 29)
            };

            var SchildVersionExpected = releases;

            var dataAccess = new DataAccess(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + configuration["AccessPfad"] + "; Persist Security Info = False;Jet OLEDB:Database Password = " + configuration["AccessPassword"] + ";");

            var schildVersionActual = dataAccess.GetSchildVersion();

            if (schildVersionActual.Year > 1)
            {
                if (SchildVersionExpected.All(x => x.Date < schildVersionActual.Date))
                {
                    Console.WriteLine(
                        "Diese Version von SchuelerFoto-NRW ist aktuell nur freigegeben für die SchILD-Version vom " +
                        SchildVersionExpected.Max().Date.ToShortDateString() +
                        ".\n\rIhre SchILD-Version: " + schildVersionActual.Date.ToShortDateString() +
                        ". Schauen Sie hier: https://github.com/stbaeumer/schuelerfoto");
                    Console.ReadKey(true);
                    Environment.Exit(0);
                }
                else
                {
                    if (dataAccess != null)
                    {
                        return dataAccess;
                    }
                    else
                    {
                        Console.WriteLine("Fehler beim Erstellen der Datenbankverbindung.");
                        Console.ReadKey(true);
                        Environment.Exit(0);
                    }
                }
            }

            else
            {
                Console.WriteLine("Fehler beim Erstellen der Datenbankverbindung.");
            }

        } while (true);
    }
    // Gibt die höchste Zahl aus einer List<string> zurück (ignoriert Nicht-Zahlen)
    public static int GetMaxNumberFromList(string[] liste)
    {
        int max = int.MinValue;
        foreach (var item in liste)
        {
            if (int.TryParse(item, out int zahl))
            {
                if (zahl > max)
                    max = zahl;
            }
        }
        return max == int.MinValue ? 0 : max; // 0, falls keine Zahl gefunden wurde
    }    
}    