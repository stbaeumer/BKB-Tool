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
using static KonfigHelper;

#pragma warning disable CS0252 // Unbeabsichtigter Verweisvergleich. Wandeln Sie die linke Seite in den Typ "string" um, um einen Wertvergleich durchzuführen.
#pragma warning disable CA2200
#pragma warning disable CS8618
#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise nullverweisenden Objekts.
#pragma warning disable CS8604 // Möglicherweise wird ein nullverweisendes Argument an eine nicht-nullbare Parameterreferenz übergeben.
#pragma warning disable CS8073
#pragma warning disable CS0472
#pragma warning disable CS8625
#pragma warning disable CS8600


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

    public enum Zweck
    {
        Mahnung,
        Statistik,
        Zeugnis
    }

    public enum Art
    {
        Kurse,
        NichtKursUnterrichte,
    }

    public enum NurBeiDiesenSchulnummern
    {
        Alle,
        AlleBisAufGesperrte,
        NurPrivilegiert,
        Nur177659,
        Nur000000
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
        Auswahl,
        Maildomain,
        EnterOderAbbrechen,
        ListString,
        MultiSelect,
        FirstRun
    }

    public enum ZipModus
    {
        Webuntis,
        Geevoo
    }

    public enum Modus
    {
        Create,
        Update,
        Read, // Wird verwendet, um Einstellungen zu lesen und den Benutzer zu fragen
        ReadSilent // Wird verwendet, um Einstellungen zu lesen, ohne den Benutzer zu fragen
    }

    public static List<string> AktSj = new List<string>()
    {
        (DateTime.Now.Month > 7 ? DateTime.Now.Year : DateTime.Now.Year - 1).ToString(),
        (DateTime.Now.Month > 7 ? DateTime.Now.Year + 1 : DateTime.Now.Year).ToString()
    };

    public static string? ConnectionStringUntis { get; set; }
    public static string? ZipKennwort { get; set; }
    public static string? PfadSchilddatenaustausch { get; private set; }
    public static List<string>? SchulnummernPrivilegiert { get; set; }
    public static string AppVersion { get; set; }
    public static List<string> SchulnummernGesperrt { get; set; }
    public static List<string> SchulnummernJedermann { get; set; }
    public static List<string> SchulnummernDebug { get; set; }
    public static List<string> Schulnummer177659 { get; set; }
    public static Color ColorPfadInProgrammen { get; set; }
    public static Color ColorPfadInDateien { get; set; }
    public static Color ColorActionInMenüs { get; set; }
    public static Color ColorEinstellungenRahmen { get; set; }
    public static Color ColorHinweise { get; set; }
    public static Color ColorKopfzeileInCSV { get; set; }
    public static Color ColorHyperlink { get; set; }
    public static Color ColorZahlen { get; set; }
    public static Color ColorÜberschrift { get; set; }
    public static Color ColorBeschreibung { get; set; }
    public static Color ColorUnterschrift { get; set; }
    public static Color ColorTextHervorheben { get; set; }
    public static Color ColorFehler { get; set; }
    public static Color ColorInfoBox { get; set; }
    public static string HilfeUrl { get; set; }

    public static string? SafeGetString(SqlDataReader reader, int colIndex)
    {
        if (!reader.IsDBNull(colIndex))
            return reader.GetString(colIndex);
        return string.Empty;
    }

    public static void DisplayHeader(IConfiguration configuration, List<string> content = null)
    {
        Console.Clear();
        AnsiConsole.Write(new FigletText("BKB-Tool").Centered().Color(ColorÜberschrift));

        var unterschrift = GetColor(ColorUnterschrift);
        var contentString = ""; //configuration["AppDescription"] ?? "BKB-Tool - Ein Werkzeug an der Schnittstelle zwischen SchILD & WebUntis";
        var header = $"[{unterschrift} link=https://github.com/stbaeumer/BKB-Tool] https://github.com/stbaeumer/BKB-Tool[/] | [{unterschrift}]GPLv3[/] | [{unterschrift}]v{AppVersion} [/]";

        if (content != null && content.Count > 0)
        {
            contentString = content.LastOrDefault() ?? string.Empty;   
        }

        //contentString = "\n"; // $"\n\n";
        var panel = new Panel(contentString)
                .Header(header)
                .HeaderAlignment(Justify.Center)
                .RoundedBorder()//.SquareBorder()
                .Expand()
                .BorderColor(Global.ColorÜberschrift);

        AnsiConsole.Write(panel);
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
        var gesamtbreite = Console.WindowWidth - 2;
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

            path.RootStyle = new Style(foreground: Color.White, background: Color.Black);
            path.SeparatorStyle = new Style(foreground: Color.White, background: Color.Black);
            path.StemStyle = new Style(foreground: Color.White, background: Color.Black);
            path.LeafStyle = new Style(foreground: Color.SpringGreen2, background: Color.Black);

            var s = linkeSeite;

            if (rechteSeite != "")
            {
                s = $"[{Global.GetColor(Global.ColorPfadInDateien)}]{s}[/]\n\n[{Global.GetColor(Global.ColorHinweise)}]Importhinweise:  [/]\n" + rechteSeite;
            }

            var panel = new Panel(s)
                .Header($"[bold {Global.GetColor(ColorInfoBox)}] Bereit für den Import [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Global.ColorInfoBox);

            AnsiConsole.Write(panel);

            if (!RunningInCodeSpace())
            {
                var ordner = Path.GetDirectoryName(linkeSeite);
                if (!string.IsNullOrEmpty(ordner))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = ordner,
                        UseShellExecute = true
                    });
                }
            }
        }
        else
        {
            AnsiConsole.MarkupLine("[aqua] " + linkeSeite + "[/]" + "[aqua]" + mitte + "[/]" + "[aqua]" + rechteSeite + " [/]");
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

                // Falls der Key nicht existiert, fügen wir ihn hinzu
                if (!jsonRoot.TryGetProperty(key, out _))
                {
                    writer.WriteString(key, finalValue);
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
            return;
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

    public static IConfiguration Konfig(string parameter, Modus modus, IConfiguration configuration, string aufforderung = "", int lfdNrVon = -1, int bis = -1, string hinweise = "", string defaultValue = "", Students? students = null, string zulässigeAuswahlOptionen = "")
    {
        object userInput = "";
        Datentyp datentyp = Datentyp.String;

        if (KonfigMetadaten.TryGetValue(parameter, out var meta))
        {
            if (string.IsNullOrWhiteSpace(aufforderung)) aufforderung = meta.Aufforderung;
            if (string.IsNullOrWhiteSpace(hinweise)) hinweise = meta.Hinweise;
            if (string.IsNullOrWhiteSpace(defaultValue) || defaultValue == "") defaultValue = meta.DefaultValue;
            // Der Datentyp wird aus der KonfigMeta entnommen
            datentyp = meta.Datentyp;
        }

        var panel = new Panel(hinweise)
            //.Header($"[bold]  {parameter}  [/]")
            .HeaderAlignment(Justify.Left)
            .SquareBorder()
            .Expand();

        if (lfdNrVon > 0 && bis > 0)
            panel.Header($"[bold]  {lfdNrVon} von {bis}  [/]");

        if (parameter == "Auswahl")
            {
                panel.BorderColor(Global.ColorÜberschrift);
            }
            else
            {
                panel.BorderColor(Global.ColorActionInMenüs);
            }

        // Der Wert aus der JSON hat Vorrang vor dem defaultwert. Nur wenn die JSON keinen Wert enthält oder der Wert nicht zulässig ist, wird der defaultwert verwendet.
        
        defaultValue = !string.IsNullOrEmpty(configuration[parameter])
            ? configuration[parameter] ?? defaultValue
            : defaultValue;

        if (datentyp == Datentyp.EnterOderAbbrechen)
        {
            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            var key = Console.ReadKey(true).Key;
            if (key != ConsoleKey.Enter)
            {
                throw new RestartException("Abbruch durch den Benutzer.");
            }
        }
        if (datentyp == Datentyp.FirstRun)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && !string.IsNullOrEmpty(defaultValue))
            {
                configuration[parameter] = defaultValue;
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            var key = Console.ReadKey(true).Key;
            if (key != ConsoleKey.Enter)
            {
                throw new RestartException("Abbruch durch den Benutzer.");
            }
            userInput = "true";
        }
        if (datentyp == Datentyp.JaNein)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && !string.IsNullOrEmpty(defaultValue) && defaultValue.ToLower().StartsWith("j") && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if (modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
            new TextPrompt<string>($"[] {aufforderung}[/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                .DefaultValue(defaultValue.ToLower() == "j" ? "j" : defaultValue.ToString())
                .ShowDefaultValue(true)
                .Validate(n =>
                {
                    if(n == "x")
                        throw new Exception("Sie haben abgebrochen.");
                    if (n.ToLower() != "j" && n.ToLower() != "n" && n.ToLower() != "ja" && n.ToLower() != "nein")
                        return ValidationResult.Error($"  Sie müssen [{Global.GetColor(Global.ColorActionInMenüs)}]ja[/] oder [{Global.GetColor(Global.ColorFehler)}]nein[/] eintippen.");                    
                    return ValidationResult.Success();
                }));
        }
        if (datentyp == Datentyp.Auswahl)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && !string.IsNullOrEmpty(defaultValue) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;                
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            var größteAuswahlZahl = GetMaxNumberFromList(zulässigeAuswahlOptionen.Split(','));

            userInput = AnsiConsole.Prompt(
            new TextPrompt<string>($"[] {aufforderung}[/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                .DefaultValue(defaultValue.ToLower() == "e" ? "e" : defaultValue.ToString())
                .ShowDefaultValue(true)
                .Validate(n =>
                {
                    if(n == "x")
                        throw new Exception("Sie haben abgebrochen.");
                    if (!zulässigeAuswahlOptionen.Split(",").Contains(n))
                        return ValidationResult.Error($"[]  Zulässige Auswahl: [bold aqua]{zulässigeAuswahlOptionen}[/][/]");                    
                    return ValidationResult.Success();
                }));
        }
        if (datentyp == Datentyp.String)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && !string.IsNullOrEmpty(defaultValue) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")                
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (string.IsNullOrEmpty(n))
                            return ValidationResult.Error("[]  Eingabe darf nicht leer sein.[/]");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Datei)
        {
            defaultValue = File.Exists(defaultValue) ? defaultValue : string.Empty;

            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && !string.IsNullOrEmpty(defaultValue) && File.Exists(defaultValue)  && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[]  {aufforderung}[/]")
                    .ShowDefaultValue(true)
                    .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (string.IsNullOrEmpty(n))
                            return ValidationResult.Error("[]  Eingabe darf nicht leer sein.[/]");
                        if (!File.Exists(n))
                            return ValidationResult.Error($"[]  Die Datei [bold red]{n}[/] existiert nicht.[/]");
                        return ValidationResult.Success();
                    })
                    .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Klassen)
        {
            var verschiedeneKlassen = students.Where(s => s.Status == "2" || s.Status == "6").Select(s => s.Klasse).Distinct().ToList();
            var interessierendeKlassen = new List<string>();

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[]  {aufforderung}[/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!string.IsNullOrEmpty(n) && n.ToLower() == "alle")
                            return ValidationResult.Success();
                        if (n.Split(',').Any(teil => verschiedeneKlassen.Any(klasse => klasse.ToLower().Contains(teil.Trim().ToLower()))))
                            return ValidationResult.Success();
                        if (verschiedeneKlassen.Any(s => s.ToLower().StartsWith(n.ToLower())))
                            return ValidationResult.Success();
                        return ValidationResult.Error($"Die Eingabe ist ungültig. \n Haben Sie alle *.dat-Dateien aus SchILD exportiert? \n Geben Sie eine Klasse an oder [{GetColor(Global.ColorActionInMenüs)}]alle[/].");
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

            if (x.ToLower() == "alle")
                userInput = string.Join(",", verschiedeneKlassen);
            else
                userInput = string.Join(",", interessierendeKlassen);
        }
        if (datentyp == Datentyp.MultiSelect)
        {
            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            var vorausgewählt = new[] { defaultValue }
                .SelectMany(s => s.Split(','))
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();

            if (vorausgewählt.Count == 0)
            {
                vorausgewählt = zulässigeAuswahlOptionen
                    .Split(',')
                    .Select(s => s.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToList();
            }

            var prompt = new MultiSelectionPrompt<string>()
                    .Title($"{aufforderung}:")
                    .NotRequired() // Not required to have a favorite fruit
                    .PageSize(10)
                    .MoreChoicesText("[grey](Mit Pfeiltasten navigieren)[/]")
                    .InstructionsText(
                        "(Mit [blue]<Leertaste>[/] auswählen, " +
                        "Mit [green]ENTER[/] bestätigen)")
                    .AddChoiceGroup("alle", zulässigeAuswahlOptionen.Split(','));

            foreach (var element in vorausgewählt) // vorausgewähltListe ist z.B. List<string>
            {
                prompt.Select<string>(element);
            }
            // Prompt anzeigen und Ergebnis als kommagetrennte Liste speichern
            var userInputList = AnsiConsole.Prompt(prompt);
            userInput = string.Join(",", userInputList);
            ZeileSchreiben($"[]  {aufforderung}[/]", userInput.ToString() ?? string.Empty);
        }
        if (datentyp == Datentyp.Url)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && (!string.IsNullOrEmpty(defaultValue) || defaultValue.StartsWith("https://")) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                    .ShowDefaultValue(true)
                    .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!n.StartsWith("https://") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("[]  Eingabe muss mit https:// beginnen.[/]");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Mail)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && (!string.IsNullOrEmpty(defaultValue) && defaultValue.StartsWith("@") && defaultValue.Contains("."))  && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                    .ShowDefaultValue(true)
                    .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!n.Contains("@") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("[]  Eingabe muss mit @ beginnen und einen Punkt enthalten.[/]");
                        if (!n.Contains(".") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("[]  Eingabe muss mit @ beginnen und einen Punkt enthalten.[/]");
                        if (!(n.Contains(".de") || !n.Contains(".org")))
                            return ValidationResult.Error("[]  Zulässige TLDs: .de und .org.[/]");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Maildomain)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && (!string.IsNullOrEmpty(defaultValue) && defaultValue.StartsWith("@") && defaultValue.Contains("."))  && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                    .ShowDefaultValue(true)
                    .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!n.StartsWith("@") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("[]  Eingabe muss mit @ beginnen und einen Punkt enthalten.[/]");
                        if (!n.Contains(".") && !string.IsNullOrEmpty(n))
                            return ValidationResult.Error("[]  Eingabe muss mit @ beginnen und einen Punkt enthalten.[/]");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue));
        }
        if (datentyp == Datentyp.Pfad)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && Path.Exists(defaultValue) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }
            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                    .ShowDefaultValue(true)
                    .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!Path.Exists(n.TrimEnd(Path.DirectorySeparatorChar)))
                            return ValidationResult.Error($" Der Pfad [{Global.GetColor(Global.ColorFehler)}]{n}[/] existiert nicht.");                        
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(string.IsNullOrEmpty(defaultValue) || !Path.Exists(defaultValue) ? Environment.CurrentDirectory : defaultValue));
            userInput = userInput.ToString()?.TrimEnd(Path.DirectorySeparatorChar) ?? string.Empty;
        }
        if (datentyp == Datentyp.Int)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && int.TryParse(defaultValue, out _) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                    .ShowDefaultValue(true)
                    .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!int.TryParse(n.ToString(), out _))
                        {
                            return ValidationResult.Error($"[{Global.GetColor(Global.ColorFehler)}]  {n}[/] ist keine zulässige Zahl. Bitte eine Zahl eingeben.");
                        }
                        if (zulässigeAuswahlOptionen != "" && !zulässigeAuswahlOptionen.Contains(n.ToString()))
                        {
                            return ValidationResult.Error($"Die Zahl {n} außerhalb des zulässigen Bereichs. Zulässige Werte: [{Global.GetColor(Global.ColorActionInMenüs)}]{zulässigeAuswahlOptionen}[/]");
                        }

                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.Float)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && float.TryParse(defaultValue, out _) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!float.TryParse(n.ToString(), out _))
                            return ValidationResult.Error($"  [{Global.GetColor(Global.ColorFehler)}]  {n}[/] ist keine zulässige Zahl.");                        
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.ListInt)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && int.TryParse(defaultValue, out _) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        var teile = n.ToString().Split(',');
                        if (!teile.All(t => int.TryParse(t.Trim(), out _)))
                            return ValidationResult.Error($"[]  {n} ist keine Liste aus Zahlen[/]");
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.ListString)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && !string.IsNullOrEmpty(defaultValue) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Nur die zulässigen Auswahloptionen werden als Defaultwert verwendet
            string default1 = "";                        
            foreach (var z in zulässigeAuswahlOptionen.Split(','))
            {
                if (!string.IsNullOrEmpty(z) && defaultValue.Split(",").Contains(z))
                {
                    default1 += z + ",";
                }
            }
            // Wennaus dem Defaultwert nichts matcht, dann werden die zulässigen Werte übernommen.
            if (default1.Length == 0)
            {
                default1 = zulässigeAuswahlOptionen;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
                AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung} ({zulässigeAuswahlOptionen}) [/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        var teile = n.ToString().Trim().Split(',');
                        if (!teile.All(t => zulässigeAuswahlOptionen.Split(',').Contains(t.Trim())))
                        {
                            return ValidationResult.Error($"[{Global.GetColor(Global.ColorFehler)}]  {n}[/] ist keine kommagetrennte Liste aus zulässigen Werten. Zulässige Werte: [{Global.GetColor(Global.ColorActionInMenüs)}]{zulässigeAuswahlOptionen}[/]");
                        }
                        
                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(default1.ToString().TrimEnd(',')));
        }
        if (datentyp == Datentyp.Abschnitt)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && int.TryParse(defaultValue, out _)  && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!int.TryParse(n.ToString(), out _))
                        {
                            return ValidationResult.Error($"[{Global.GetColor(Global.ColorFehler)}]  {n}[/] ist keine zulässige Zahl. Bitte eine Zahl eingeben.");
                        }
                        else
                        {
                            if (n != "1" && n != "2")
                                return ValidationResult.Error($"[]  {n} ist ungültig. Erlaubt sind 1 und 2.[/]");
                        }

                        return ValidationResult.Success();
                    })
                .DefaultValue<string>(defaultValue.ToString()));
        }
        if (datentyp == Datentyp.DateTime)
        {
            // Der Wert wird im Read-Modus nicht erneut abgefragt, wenn der Wert plausibel ist und schon ein Wert in der configuration existiert.
            if ((modus == Modus.ReadSilent || modus == Modus.Read) && DateTime.TryParse(defaultValue, out _) && !string.IsNullOrEmpty(configuration[parameter]))
            {
                configuration[parameter] = defaultValue;
                if(modus != Modus.ReadSilent)
                    ZeileSchreiben(aufforderung, defaultValue);
                return configuration;
            }

            // Wenn der Wert abgefragt wird, dann wird ein Panel mit dem Hinweis angezeigt
            AnsiConsole.Write(panel);

            userInput = AnsiConsole.Prompt(
                new TextPrompt<string>($"[] {aufforderung}[/]")
                .PromptStyle(Global.GetColor(Global.ColorActionInMenüs))
                    .ShowDefaultValue(true)
                    .Validate(n =>
                    {
                        if(n == "x")
                            throw new Exception("Sie haben abgebrochen.");
                        if (!DateTime.TryParse(n.ToString(), out _))
                        {
                            return ValidationResult.Error($"[]  {n} ist kein Datum (TT.MM.JJJJ)[/]");
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

    public static IConfiguration EinstellungenDurchlaufen(IConfiguration configuration, Global.Modus modus = Global.Modus.Read)
    {
        // Wenn User.json noch nicht existiert, dann erstellen
        if (!File.Exists(Path.Combine(Directory.GetCurrentDirectory(), Global.User + ".json")))
        {
            var existiertnichtOderNichtBeschreibbar = true;

            do
            {
                if (!Directory.Exists(Directory.GetCurrentDirectory()) || !IsDirectoryWritable(Directory.GetCurrentDirectory()))
                {
                    AnsiConsole.MarkupLine($"[red]Das Verzeichnis [bold {Global.GetColor(Global.ColorPfadInDateien)}]" + Directory.GetCurrentDirectory() + "[/] existiert nicht oder ist nicht beschreibbar. Das muss korrigiert werden.[/]");
                    AnsiConsole.MarkupLine($"[red]Drücken Sie eine beliebige Taste, um fortzufahren...[/]");
                    while (Console.KeyAvailable) Console.ReadKey(true);

                    Console.ReadKey();
                    return configuration;
                }
                else
                {
                    existiertnichtOderNichtBeschreibbar = false;
                }
            } while (existiertnichtOderNichtBeschreibbar);

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
            DisplayHeader(configuration);
            configuration = Global.Konfig("ZustimmungLizenz", Global.Modus.Read, configuration);            
        }
        
        var kon = KonfigMetadaten.Where(e =>
            e.Value.InitialAbfragen == true || // nur solche, die initial abgefragt werden sollen
            (
                e.Value.InGrundeinstellungAbfragen == true && // nur solche, die in Grundeinstellung abgefragt werden sollen
                //e.Value.NurBeiDiesenSchulnummern != null && // nur solche, die für bestimmte Schulnummern abgefragt werden sollen
                //e.Value.NurBeiDiesenSchulnummern.Contains(configuration["Schulnummer"]) &&
                configuration[e.Key] != null && // nur solche, die in der Konfiguration gespeichert sind
                configuration[e.Key] != ""
            )).ToList();

        // Durchlaufe alle Einstellungen gemäß KonfigHelper
        for (var i = 0; i < kon.Count(); i++)
        {   
            DisplayHeader(configuration);         
            configuration = Konfig(kon[i].Key, modus, configuration, kon[i].Value.Aufforderung, i + 1, kon.Count(), kon[i].Value.Hinweise, kon[i].Value.DefaultValue);
        }

        return configuration;
    }

    static object CreateBkbJsonContent()
    {
        DotEnv.Load();

        return new
        {
            ZustimmungLizenz = "nein",
/*
                        ConnectionStringUntis = Verschluesseln(Environment.GetEnvironmentVariable("UNTIS_CONNECTION_STRING") ?? ""),
                        Dokumentenverwaltung = Verschluesseln(@"\\fs01\SchILD-NRW\Dokumentenverwaltung"),
                        Schulnummer = Verschluesseln("177659" ?? ""),
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
                        PfadFotosAusSchILD = "",
                        Kalenderfilter = Verschluesseln(""),
                        Auswahl = "e",
                        OnlineHilfeURL = Verschluesseln("https://github.com/stbaeumer/BKB-Tool/wiki"),

                        AppName = Verschluesseln("BKB-Tool"),
                        AppVersion = Verschluesseln("1.0.0"),
                        AppDescription = Verschluesseln("BKB-Tool - Ein Werkzeug an der Schnittstelle zwischen SchILD und Webuntis."),
                        Klassen = Verschluesseln("HBG"),
                        Vergleich = Verschluesseln("n"),
                        Kennwort = Verschluesseln(""),
                        InputFolder = Verschluesseln(""),
                        OutputFolder = Verschluesseln(""),
                        Halbjahreszeugnisdatum = Verschluesseln(new DateTime(Convert.ToInt32(Global.AktSj[1]), 02, 01).ToString("dd.MM.yyyy")),
                        Halbjahreskonferenzdatum = Verschluesseln(new DateTime(Convert.ToInt32(Global.AktSj[1]), 02, 01).ToString("dd.MM.yyyy")),
                        Abschnittswechsel = Verschluesseln(new DateTime(Convert.ToInt32(Global.AktSj[1]), 02, 01).ToString("dd.MM.yyyy")),
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
                        FehlzeitenWaehrendDerLetztenTagBleibenUnberuecksichtigt = Verschluesseln("3"),
                        MaximaleAnzahlFehlstundenProTag = Verschluesseln("8"),
                        Abschnitt = Verschluesseln("1"),
                        Chat = Verschluesseln(""),
                        AusbuchenNachWievielTagen = Verschluesseln(""),
                        DatenimportLetztesDatum = Verschluesseln(DateTime.Now.ToString("dd.MM.yyyy")),
                        MaxDateiAlter = Verschluesseln("3"),
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
                        Body = Verschluesseln("Guten Morgen #Lehrer#,\n\nbitte beachten Sie den Anhang.\n\nErläuterungen dazu finden Sie hier: https://bkb.wiki/konzepte:stundenplanungskonzept#information_aller_lehrkraefte_per_mail \n\nViele Grüße aus der Schulverwaltung"),
                        AccessPfad = Verschluesseln(@"\\fs01\SchILD-NRW\DB\Test.mdb"),
                        PdfInputFolder = Verschluesseln(@"PDF-Input"),
                        PdfOutputFolder = Verschluesseln(@"PDF-Output"),
                        PfadDokumentenverwaltung = Verschluesseln(@"\\fs01\SchILD-NRW\Dokumentenverwaltung"),
                        BodyMassenmail = Verschluesseln("Guten Morgen #Lehrer#,\n\nbitte beachten Sie den Anhang.\n\nErläuterungen dazu finden Sie hier: https://bkb.wiki/konzepte:stundenplanungskonzept#information_aller_lehrkraefte_per_mail \n\nViele Grüße aus der Schulverwaltung"),
                        Verbose = Verschluesseln("false"),
                        SmtpUserMassenmail = Verschluesseln("campusfest@berufskolleg-borken.de"),
                        SmtpServerMassenmail = Verschluesseln("smtp-hve.office365.com"),
                        BetreffMassenmail = Verschluesseln("Save the date: Campusfest am Berufskolleg Borken"),
                        Schlüsselwörter = Verschluesseln("Jahreszeugnis, Abschlusszeugnis, Abgangszeugnis, Zeugnis"),
                        Teilleistungsarten = Verschluesseln("Vornote,Abschluss-Schriftl."),
                        LehrkraefteSonderzeiten = Verschluesseln(""),
                        VolleStelle = Verschluesseln("25,5"),
                        Lk1faecher = Verschluesseln("D,BI,M,E")
              */
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
        var panel = new Panel($"Weiter mit [bold {Global.GetColor(Global.ColorActionInMenüs)}]Anykey[/] oder mit [bold {Global.GetColor(Global.ColorActionInMenüs)}]e[/] Einstellungen durchlaufen oder mit [bold {Global.GetColor(Global.ColorActionInMenüs)}]h[/] Onlinehilfe öffnen.")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Global.ColorUnterschrift);

        AnsiConsole.Write(panel);

        var weiter = Console.ReadKey(true); // true unterdrückt die Ausgabe des Zeichens im Terminal

        DisplayHeader(configuration);
        if (weiter.Key == ConsoleKey.E)
        {
            configuration = EinstellungenDurchlaufen(configuration, Global.Modus.Update);
            return;
        }

        if (weiter.Key == ConsoleKey.H)
        {
            var url = "";
            if (menüeintrag != null)
            {
                url = $"#{menüeintrag.Titel.Split(':')[0].ToLower()}";
            }

            OpenWebseite($"{HilfeUrl}{url}");
            return;
        }
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
            configuration = Global.Konfig("AccessPfad", Global.Modus.Read, configuration);
            configuration = Global.Konfig("AccessPassword", Global.Modus.Read, configuration);

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
                    while (Console.KeyAvailable) Console.ReadKey(true);

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
                        while (Console.KeyAvailable) Console.ReadKey(true);

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

    // Hilfsmethode: Prüft, ob ein Verzeichnis beschreibbar ist
    private static bool IsDirectoryWritable(string dirPath)
    {
        try
        {
            string testFile = Path.Combine(dirPath, Path.GetRandomFileName());
            using (FileStream fs = File.Create(testFile, 1, FileOptions.DeleteOnClose)) { }
            return true;
        }
        catch
        {
            return false;
        }
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

    public static string GetColor(Color farbe)
    {
        // Wandel den ersten Buchtaben in einen Kleinbchstaben um:        
        var s = farbe.ToString();
        if (string.IsNullOrEmpty(s)) return s;
        return char.ToLowerInvariant(s[0]) + s.Substring(1);
    }

    internal static bool IsExplorerOpen(string pfad)
    {
        var letzterTeil = Path.GetFileName(pfad.TrimEnd(Path.DirectorySeparatorChar));

        // Überprüfe, ob der Explorer für den angegebenen Pfad bereits geöffnet ist
        var processes = System.Diagnostics.Process.GetProcessesByName("explorer");
        foreach (var process in processes)
        {
            try
            {
                var mainWindowTitle = process.MainWindowTitle;
                if (!string.IsNullOrEmpty(mainWindowTitle) && mainWindowTitle.Contains(letzterTeil))
                {
                    return true;
                }
            }
            catch
            {
                // Ignoriere Prozesse, auf die nicht zugegriffen werden kann
            }
        }
        return false;
    }
}    


public class KonfigMeta
{
    public string Key { get; set; }
    public string DefaultValue { get; set; }
    public string Aufforderung { get; set; }
    public string Hinweise { get; set; }
    public Global.Datentyp Datentyp { get; set; }
    public List<string> NurBeiDiesenSchulnummern { get; internal set; }

    /// <summary>
    /// Sobald der Wert in die json aufgenommen wurde, wird er in den Grundenstellungen immer wieder abgefragt.
    /// Wenn der Wert nicht in den Grundeinstellungen abgefragt werden soll, dann auf false     
    /// </summary>
    public bool InGrundeinstellungAbfragen { get; internal set; }
    public bool InitialAbfragen { get; internal set; }
}

public static class KonfigHelper
{
    
    /*
    configuration = Konfig("PfadDownloads", modus, configuration, @"Downloads-Verzeichnis", "Geben Sie den Pfad des Downloads-Verzeichnisses an. In der Regel wird das Verzeichnis bereits richtig vorgeschlagen. Dann einfach [bold springGreen2]ENTER[/] drücken:", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"));
        configuration = Konfig("Schulnummer", modus, configuration, @"Schulnummer", "Geben Sie Ihre Schulnummer an. Je nach Schulnummer werden evtl. unterschiedliche Funktionen angeboten.");
        configuration = Konfig("PfadSchilddatenaustausch", modus, configuration, @"SchILD-Ausgabeverzeichnis", $"Geben Sie das Verzeichnis an, das in SchILD unter [{Global.GetColor(Global.ColorPfadInProgrammen)}]Datenaustausch > Schnittstelle SchILD-NRW > Export[/] als [{Global.GetColor(Global.ColorPfadInProgrammen)}]Ausgabeverzeichnis[/] eingetragen ist. Wenn dort nichts eingetragen ist, tragen Sie dort das selbe Verzeichnis ein, das Sie auch hier eintragen.\n[{Global.GetColor(Global.ColorHinweise)}]Hinweis:[/] Falls an mehreren Arbeitsplätzen parallel mit [{Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] gearbeitet wird, kann es sinnvoll sein, dass jede*r ein eigenes Ausgabeverzeichnis in SchILD bekommt.", @"\\fs01\SchILD-NRW\Ausgabeverzeichnis");
        configuration = Konfig("PfadDokumentenverwaltung", modus, configuration, "Pfad zur Dokumentenverwaltung", $"Geben Sie das Verzeichnis an, das in SchILD unter [{Global.GetColor(Global.ColorPfadInProgrammen)}]Extras > Programmeinstellungen > Globale Einstellungen > Dokumentenverwaltung[/] als [{Global.GetColor(Global.ColorPfadInProgrammen)}]Dokumentenverzeichnis[/] eingetragen ist. Die Dokumentenverwaltung muss eingeschaltet sein. Wenn dort nichts eingetragen ist, tragen Sie dort das selbe Verzeichnis ein, das Sie auch hier angeben:", @"\\fs01\SchILD-NRW\Dokumentenverwaltung");
        configuration = Konfig("MaxDateiAlter", modus, configuration, "Wie viele Tage dürfen Dateien höchstens alt sein?", $"Geben Sie an, wie viele Tage Dateien höchstens alt sein dürfen, um vom [{Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] für das Einlesen akzeptiert zu werden. Die Angabe einer (möglichst niedrigen) Zahl soll sicherstellen, dass nicht versehntlich veraltete Dateien eingelesen werden.");
        configuration = Konfig("LK1-Fächer", modus, configuration, "Fächer des 1.LKs", $"Geben Sie kommasepariert die möglichen Fächer des 1.LKs an. Bitte nur die Fächer angeben, ohne 'LK' etc. Beispiel: D,BI,M,E", "D,BI,M,E");

        if (modus == Modus.Update && SchulnummernPrivilegiert.Contains(configuration["Schulnummer"]))
        {
            //configuration = Konfig("MailDomain", modus, configuration, "Mail-Domain für Schüler*innen", $"Geben Sie die Mail-Domain für Ihre Schüler*innen an. Ihre Eingabe muss mit [{Global.GetColor(Global.ColorZahlen)}]@[/] beginnen und einen [{Global.GetColor(Global.ColorZahlen)}]Punkt[/] enthalten. Beispiel: [springGreen2 bold]@students.meine-schule.de[/]", Datentyp.Mail);
            configuration = Konfig("ConnectionStringUntis", modus, configuration, "ConnectionStringUntis (optional)");
            configuration = Konfig("SmtpUser", modus, configuration, "Mail-Benutzer");
            configuration = Konfig("SmtpKennwort", modus, configuration, "Mail-Kennwort");
            configuration = Konfig("SmtpPort", modus, configuration, "SMTP-Port");
            configuration = Konfig("SmtpServer", modus, configuration, "SMTP-Server");
            configuration = Konfig("NetmanMailReceiver", modus, configuration, "Wem soll die Netman-Mail geschickt werden?");
        }
    
    */


    public static readonly Dictionary<string, KonfigMeta> KonfigMetadaten = new()
    {        
        ["AppDescription"] = new KonfigMeta        
        {
            Key = "AppDescription",
            DefaultValue = "BKB-Tool - Ein Werkzeug an der Schnittstelle zwischen SchILD und Webuntis.",
            Aufforderung = "",
            Hinweise = "Kurze Beschreibung der Anwendung.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Abschnitt"] = new KonfigMeta
        {
            Key = "Abschnitt",
            DefaultValue = "1",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]Lernabschnitt[/] an. Das Schuljahr beginnt immer mit Abschnitt [{Global.GetColor(Global.ColorZahlen)}]1[/]. \nI.d.R. wechselt der Abschnitt nach den Halbjahreszeugnissen auf Abschnitt [{Global.GetColor(Global.ColorZahlen)}]2[/].",
            Datentyp = Global.Datentyp.Abschnitt,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = true,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Abschnittswechsel"] = new KonfigMeta
        {
            Key = "Abschnittswechsel",
            DefaultValue = new DateTime(DateTime.Now.Month > 7 ? DateTime.Now.Year + 1 : DateTime.Now.Year, 2, 1).ToShortDateString(),
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Datum des Abschnittswechsels[/] an. I.d.R. ist der [{Global.GetColor(Global.ColorZahlen)}]{new DateTime(DateTime.Now.Month > 7 ? DateTime.Now.Year + 1 : DateTime.Now.Year, 2, 1).ToShortDateString()}[/] (oder ein anderes Datum im Februar) der richtige Wert.",
            Datentyp = Global.Datentyp.DateTime,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["AccessPfad"] = new KonfigMeta
        {
            Key = "AccessPfad",
            DefaultValue = @"\\fs01\SchILD-NRW\DB\Test.mdb",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]Pfad zur Access-Datenbank[/] an. Beispiel:[{Global.GetColor(Global.ColorPfadInProgrammen)}]{@"\\fs01\SchILD-NRW\DB\Test.mdb"}[/]\nGeben Sie stets den kompletten Pfad inklusive Dateinname an.",
            Datentyp = Global.Datentyp.Pfad,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["AccessPassword"] = new KonfigMeta
        {
            Key = "AccessPassword",
            DefaultValue = Environment.GetEnvironmentVariable("ACCESS_PASSWORD") ?? "",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Kennwort[/] der Access-Datenbank an. Es geht nicht um Ihr persönliches Kennwort in SchILD, sondern um das Kennwort der Access-Datenbank selbst.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Auswahl"] = new KonfigMeta
        {
            Key = "Auswahl",
            DefaultValue = "1",
            Aufforderung = "Ihre Auswahl bitte",
            Hinweise = $"Geben Sie eine [bold {Global.GetColor(Global.ColorTextHervorheben)}]Zahl[/] ein oder [bold {Global.GetColor(Global.ColorTextHervorheben)}]e[/] für Einstellungen oder [bold {Global.GetColor(Global.ColorTextHervorheben)}]h[/] für Onlinehilfe",
            Datentyp = Global.Datentyp.Auswahl,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Betreff"] = new KonfigMeta
        {
            Key = "Betreff",
            DefaultValue = "Betreff",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]Betreff[/] an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["BccAdresse"] = new KonfigMeta
        {
            Key = "BccAdresse",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie [{Global.GetColor(Global.ColorInfoBox)}]BCC-Adresse[/] an. Geben Sie ein Leerzeichen ein, wenn Sie keine BCC-Adresse wünschen.",
            Datentyp = Global.Datentyp.Mail,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Body"] = new KonfigMeta
        {
            Key = "Body",
            DefaultValue = $"Guten Morgen,\n\n...",
            Aufforderung = "",
            Hinweise = $"Geben Sie [{Global.GetColor(Global.ColorInfoBox)}]Nachricht[/] ein. Geben Sie \\n ein, um einen Zeilenumbruch zu erzeugen. geben Sie #Lehrer# als Platzhalter für den Namen des Lehrers ein. Der Name wird dann automatisch ersetzt.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["ConnectionStringSchild"] = new KonfigMeta
        {
            Key = "ConnectionStringSchild",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]ConnectionString[/] für SchILD an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.Schulnummer177659
        },
        ["ConnectionStringWebuntis"] = new KonfigMeta
        {
            Key = "ConnectionStringWebuntis",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]ConnectionString[/] für Webuntis an (optional).",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernPrivilegiert
        },
        ["ConnectionStringUntis"] = new KonfigMeta
        {
            Key = "ConnectionStringUntis",
            DefaultValue = Environment.GetEnvironmentVariable("UNTIS_CONNECTION_STRING") ?? "",
            Aufforderung = "",
            Hinweise = $"Bitte geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]ConnectionString[/] für Untis an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernPrivilegiert
        },
        ["FirstRun"] = new KonfigMeta
        {
            Key = "FirstRun",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = "",
            Datentyp = Global.Datentyp.FirstRun,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["FehlzeitenVorDemAbschnittswechselBeruecksichtigen"] = new KonfigMeta
        {
            Key = "FehlzeitenVorDemAbschnittswechselBeruecksichtigen",
            DefaultValue = "Nein",
            Aufforderung = "",
            Hinweise = $"Sollen die [{Global.GetColor(Global.ColorInfoBox)}]Fehlzeiten[/] vor dem Abschnittswechsel auf dem Zeugnis hinzugefügt werden?",
            Datentyp = Global.Datentyp.JaNein,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["InteressierendeUnterrichtsgruppen"] = new KonfigMeta
        {
            Key = "InteressierendeUnterrichtsgruppen",
            DefaultValue = "1.HJ,U",
            Aufforderung = "",
            Hinweise = $"Geben Sie alle [{Global.GetColor(Global.ColorInfoBox)}]Unterrichtsgruppen[/] an, die am Stichtag anwesend sein werden. Unterrichte ohne Unterrichtsgrupe werden immer berücksichtigt.\nGroß- und Kleinschreibung beachten!\nWenn Sie alle Unterrichtsgruppen berücksichtigen wollen, schreiben wie das Wort [{Global.GetColor(Global.ColorActionInMenüs)}]alle[/] gewählt.",
            Datentyp = Global.Datentyp.MultiSelect,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Klassen"] = new KonfigMeta
        {
            Key = "Klassen",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie die [{Global.GetColor(Global.ColorInfoBox)}]interessierende(n) Klasse(n)[/] an. Mehrere Klassen sind mit Komma zu trennen. Es können auch Namensteile von Klassen angegeben werden, wordurch alle Klassen gewählt werden, deren Klassenname den Namensteil enthält. Alle Klassen werden mit dem Wort [bold springGreen2]alle[/] gewählt.",
            Datentyp = Global.Datentyp.Klassen,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Konferenzdatum"] = new KonfigMeta
        {
            Key = "Konferenzdatum",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Konferenzdatum[/] an. Das kann später in SchILD (mit einem Gruppenprozess) erneut geändert werden.",
            Datentyp = Global.Datentyp.DateTime,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Kursarten"] = new KonfigMeta
        {
            Key = "Kursarten",
            DefaultValue = " G, L, A, Z, V, P",
            Aufforderung = "",
            Hinweise = $"Geben Sie im Folgenden an, wie sich bei Ihnen die [{Global.GetColor(Global.ColorInfoBox)}]Kursarten[/] aus den Fächern ableiten lassen. Wenn Ihre LKs beispielsweise so aussehen: [{Global.GetColor(Global.ColorHinweise)}]M L1, D L2, E L1[/], dann geben Sie hier für LKs (inkl. führendes Leerzeichen) '[{Global.GetColor(Global.ColorHinweise)}] L[/]' an." +
            $"\nDie Kursarten in SchILD sind: [{Global.GetColor(Global.ColorHinweise)}]GK, LK, AB, ZK, VTF, PJK[/]" +
            $"\nIhre vollständige Eingabe könnte also so aussehen: [{Global.GetColor(Global.ColorActionInMenüs)}] G, L, A, Z, V, P[/]." +
            $"\nWenn Sie die automatische Zuordnung nicht nutzen möchten, weil Sie z.B. keine Kurse nutzen, geben Sie weniger als 5 Kommas ein.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Lk1faecher"] = new KonfigMeta
        {
            Key = "Lk1faecher",
            DefaultValue = "D,BI,M,E",
            Aufforderung = "",
            Hinweise = $"Geben Sie kommasepariert die möglichen [{Global.GetColor(Global.ColorInfoBox)}]Fächer des 1.LKs[/] an. Bitte nur die Fächer angeben, ohne 'LK' etc. Beispiel: [{Global.GetColor(Global.ColorActionInMenüs)}]D,BI,M,E[/]",
            Datentyp = Global.Datentyp.Maildomain,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["LehrkraefteSonderzeiten"] = new KonfigMeta
        {
            Key = "LehrkraefteSonderzeiten",
            DefaultValue = "098,099,007,360,160",
            Aufforderung = "",
            Hinweise = "Welche [{Global.GetColor(Global.ColorInfoBox)}]Anrechnungsgründe[/] sollen ignoriert bzw. auf 0 gesetzt werden?",
            Datentyp = Global.Datentyp.ListInt,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["MailDomain"] = new KonfigMeta
        {
            Key = "MailDomain",
            DefaultValue = "@students.berufskolleg-borken.de",
            Aufforderung = "",
            Hinweise = $"Geben Sie die schulische [{Global.GetColor(Global.ColorInfoBox)}]Mail-Domain[/] für Mailadressen der Schüler*innen an. Bsp: [{Global.GetColor(Global.ColorHyperlink)}]@students.berufskolleg-borken.de[/]. Ihre Eingabe muss mit [{Global.GetColor(Global.ColorHyperlink)}]@[/] beginnen und mit [{Global.GetColor(Global.ColorHyperlink)}].de[/] etc. enden.",
            Datentyp = Global.Datentyp.Maildomain,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["NetmanMailReceiver"] = new KonfigMeta
        {
            Key = "",
            DefaultValue = Environment.GetEnvironmentVariable("NETMAN_MAIL_RECEIVER"),
            Aufforderung = "NetmanMailReceiver",
            Hinweise = $"Geben Sie die [{Global.GetColor(Global.ColorInfoBox)}]Empfänger-E-Mail-Adresse[/] der Netman-Datei an.",
            Datentyp = Global.Datentyp.Mail,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["NetmanMailBccReceiver"] = new KonfigMeta
        {
            Key = "NetmanMailBccReceiver",
            DefaultValue = Environment.GetEnvironmentVariable("NETMAN_MAIL_BCC_RECEIVER"),
            Aufforderung = "",
            Hinweise = $"Geben Sie die [{Global.GetColor(Global.ColorInfoBox)}]BCC-Empfänger-E-Mail-Adresse[/] der Netman-Datei an.",
            Datentyp = Global.Datentyp.Mail,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["PfadDownloads"] = new KonfigMeta
        {
            Key = "PfadDownloads",
            DefaultValue = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
            Aufforderung = "",
            Hinweise = $"Geben Sie den Pfad des [{Global.GetColor(Global.ColorInfoBox)}]Downloads-Verzeichnisses[/] an. In der Regel wird das Verzeichnis bereits richtig vorgeschlagen. Dann einfach [bold springGreen2]ENTER[/] drücken:",
            Datentyp = Global.Datentyp.Pfad,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = true,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["PfadFotosAusSchILD"] = new KonfigMeta
        {
            Key = "PfadFotosAusSchILD",
            DefaultValue = $"\\\\fs01\\SchILD-NRW\\Fotos_aus_SchILD",
            Aufforderung = "",
            Hinweise = $"Wo liegen die [{Global.GetColor(Global.ColorInfoBox)}]Fotos[/], die Sie aus SchILD exportiert haben ([{Global.GetColor(Global.ColorPfadInProgrammen)}]Datenaustausch > Fotos > Fotos exportieren[/])? Der Export ist wichtig, um nur fehlende Fotos an SchILD zu übergeben. Geben Sie den Pfad zu den Fotos aus SchILD an. Beispiel: [{Global.GetColor(Global.ColorPfadInProgrammen)}]\\\\fs01\\SchILD-NRW\\Fotos_aus_SchILD[/]",
            Datentyp = Global.Datentyp.Pfad,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.Schulnummer177659
        },
        ["PfadSchilddatenaustausch"] = new KonfigMeta
        {
            Key = "PfadSchilddatenaustausch",
            DefaultValue = $"\\\\fs01\\SchILD-NRW\\Ausgabeverzeichnis",
            Aufforderung = "",
            Hinweise = $"Geben Sie den Pfad und das Verzeichnis an, das in SchILD eingetragen ist als [{Global.GetColor(Global.ColorInfoBox)}]Ausgabeverzeichnis[/] unter: \n[{Global.GetColor(Global.ColorPfadInProgrammen)}]Datenaustausch > Schnittstelle SchILD-NRW > Export[/]",
            Datentyp = Global.Datentyp.Pfad,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = true,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["PdfKennwort"] = new KonfigMeta
        {
            Key = "PdfKennwort",
            DefaultValue = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
            Aufforderung = "",
            Hinweise = $"Geben Sie optional ein [{Global.GetColor(Global.ColorInfoBox)}]Kennwort[/] für PDF-Dateien an, die Sie mit [{Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] erstellen möchten. Das Kennwort wird für alle PDF-Dateien verwendet, die Sie mit [{Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] erstellen. Wenn Sie kein Kennwort wünschen, dann ein Leerzeichen eingeben.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["MaxDateiAlter"] = new KonfigMeta
        {
            Key = "MaxDateiAlter",
            DefaultValue = "3",
            Aufforderung = "",
            Hinweise = $"Geben Sie an, [{Global.GetColor(Global.ColorInfoBox)}]wie viele Tage[/] Dateien höchstens alt sein dürfen, um vom BKB-Tool für das Einlesen akzeptiert zu werden. Die Angabe einer (möglichst niedrigen) Zahl soll sicherstellen, dass nicht versehentlich veraltete Dateien eingelesen werden.",
            Datentyp = Global.Datentyp.Int,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = true,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["NurDieseGruende"] = new KonfigMeta
        {
            Key = "NurDieseGruende",
            DefaultValue = "200",
            Aufforderung = "",
            Hinweise = $"Geben Sie die [{Global.GetColor(Global.ColorInfoBox)}]interessierenden Gründe[/] an.",
            Datentyp = Global.Datentyp.ListInt,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["NurNeueFotosExportieren"] = new KonfigMeta
        {
            Key = "NurNeueFotosExportieren",
            DefaultValue = "Ja",
            Aufforderung = "Nur neue und veränderte Fotos nach Webuntis & Co. exportieren?",
            Hinweise = $"Sie können auswählen, ob alle [{Global.GetColor(Global.ColorInfoBox)}]Fotos[/] exportiert werden sollen oder nur neue und veränderte.",
            Datentyp = Global.Datentyp.JaNein,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["OffeneFehlstunden"] = new KonfigMeta
        {
            Key = "OffeneFehlstunden",
            DefaultValue = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
            Aufforderung = "",
            Hinweise = $"Es gibt [{Global.GetColor(Global.ColorInfoBox)}]offene Fehlstunden[/]. Offene Fehlstunden sind weder entschuldigt noch unentschuldigt und werden im Zeugnis nicht berücksichtigt. Mit [{Global.GetColor(Global.ColorActionInMenüs)}]ENTER[/] geht es ohne Berücksichtigung der offenen Fehlstunden weiter. Abbruch mit [{Global.GetColor(Global.ColorFehler)}]ANYKEY[/].",
            Datentyp = Global.Datentyp.JaNein,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },         
        ["PfadDokumentenverwaltung"] = new KonfigMeta
        {
            Key = "PfadDokumentenverwaltung",
            DefaultValue = @"\\fs01\SchILD-NRW\Dokumentenverwaltung",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Verzeichnis[/] an, das in SchILD unter [{Global.GetColor(Global.ColorPfadInProgrammen)}]Extras > Programmeinstellungen > Globale Einstellungen > Dokumentenverwaltung[/] als [{Global.GetColor(Global.ColorPfadInProgrammen)}]Dokumentenverzeichnis[/] eingetragen ist.",
            Datentyp = Global.Datentyp.Pfad,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["PfadFotosImSchILD-Ordner"] = new KonfigMeta
        {
            Key = "PfadFotosImSchILD-Ordner",
            DefaultValue = "\\\\fs01\\SchILD-NRW\\Fotos",
            Aufforderung = "",
            Hinweise = $"Wo werden die aufbereiteten Schüler*innenfotos in Klassenordnern gespeichert? Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]Pfad[/] an! Beispiel: [aqua]\\\\fs01\\SchILD-NRW\\Fotos[/]. Beachten Sie, dass alle Fotos im Zielordner den Benutzernamen (der Teil vor dem @ der Mail-Adresse) als Namen haben müssen.",
            Datentyp = Global.Datentyp.Pfad,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["PfadLitteraImport"] = new KonfigMeta
        {
            Key = "PfadLitteraImport",
            DefaultValue = @"\\fs01\Littera\Atlantis Import Daten",
            Aufforderung = "",
            Hinweise = $"Wohin soll die neue Datei nach dem Erstellen verschoben werden?Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]Pfad[/] an.",
            Datentyp = Global.Datentyp.Pfad,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["RotateFotos"] = new KonfigMeta
        {
            Key = "RotateFotos",
            DefaultValue = "0",
            Aufforderung = "",
            Hinweise = $"Sie können vor dem Hochladen die Fotos drehen. Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]Drehwinkel in Grad[/] an. Beispiel: [{Global.GetColor(Global.ColorZahlen)}]90[/] für 90 Grad im Uhrzeigersinn. [{Global.GetColor(Global.ColorZahlen)}]0[/] bedeutet keine Drehung.",
            Datentyp = Global.Datentyp.Int,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Schlüsselwörter"] = new KonfigMeta
        {
            Key = "Schlüsselwörter",
            DefaultValue = "Jahreeszeugnis, Abschlusszeugnis, Abgangszeugnis, Zeugnis",
            Aufforderung = "",
            Hinweise = $"Geben Sie kommagetrennt interessierende [{Global.GetColor(Global.ColorInfoBox)}]Schlüsselwörter[/] an (z.B. Abgangszeugnis, Abschlusszeugnis, Jahreszeugnis). BKB-Tool durchsucht die PDF-Dateien im Ordner nach den Wörtern. Sobald ein Schlüsselwort matcht, wird die Datei in das Dokumentenverzeichnis kopiert.",
            Datentyp = Global.Datentyp.ListString,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["SchipsOderZeugnisseOderAnderePdfs"] = new KonfigMeta
        {
            Key = "SchipsOderZeugnisseOderAnderePdfs",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Wählen Sie zwischen den [{Global.GetColor(Global.ColorInfoBox)}]Optionen[/]:\n1: Schips\n2: Notenlisten\n3: andere PDFs",
            Datentyp = Global.Datentyp.Int,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Schulnummer"] = new KonfigMeta
        {
            Key = "Schulnummer",
            DefaultValue = "177659",
            Aufforderung = "",
            Hinweise = $"Geben Sie Ihre [{Global.GetColor(Global.ColorInfoBox)}]Schulnummer[/] an. Je nach Schulnummer werden evtl. unterschiedliche Funktionen angeboten.",
            Datentyp = Global.Datentyp.Int,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = true,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["SmtpKennwort"] = new KonfigMeta
        {
            Key = "SmtpKennwort",
            DefaultValue = Environment.GetEnvironmentVariable("SMTP_KENNWORT") ?? "",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]SMTP-Kennwort[/] an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["SmtpPort"] = new KonfigMeta
        {
            Key = "SmtpPort",
            DefaultValue = Environment.GetEnvironmentVariable("SMTP_PORT") ?? "",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]SMTP-Port[/] an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },        
        ["SmtpServer"] = new KonfigMeta
        {
            Key = "SmtpServer",
            DefaultValue = Environment.GetEnvironmentVariable("SMTP_SERVER") ?? "",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]SMTP-Server[/] an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["SmtpUser"] = new KonfigMeta
        {
            Key = "SmtpUser",
            DefaultValue = Environment.GetEnvironmentVariable("SMTP_USER") ?? "",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]SMTP-Benutzer[/] an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["SprechtagsDatum"] = new KonfigMeta
        {
            Key = "SprechtagsDatum",
            DefaultValue = $"01.09.{DateTime.Now.Year}",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Datum des Sprechtags[/] an (tt.mm.jjjj).",
            Datentyp = Global.Datentyp.DateTime,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["StatistikDatum"] = new KonfigMeta
        {
            Key = "StatistikDatum",
            DefaultValue = $"01.09.{DateTime.Now.Year}",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Datum der Abgabe[/] an. Das ist wichtig, um diejenigen Unterrichte auszuschließen, die befristet sind und nicht am Stichtag stattfinden.",
            Datentyp = Global.Datentyp.DateTime,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["TeamsChatAuswahl"] = new KonfigMeta
        {
            Key = "TeamsChatAuswahl",
            DefaultValue = "1",
            Aufforderung = "",
            Hinweise = $"Bitte eine [{Global.GetColor(Global.ColorInfoBox)}]Zahl[/] auswählen:",
            Datentyp = Global.Datentyp.Int,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Teilleistungsarten"] = new KonfigMeta
        {
            Key = "Teilleistungsarten",
            DefaultValue = "Vornote,Abschluss-Schriftl.,Abschluss-Mündl.",
            Aufforderung = "Welche Teilleistungsarten (kommagetrennt) sollen gezogen werden?",
            Hinweise = $"Die [{Global.GetColor(Global.ColorInfoBox)}]Teilleistungsart(en)[/] in Webuntis und in SchILD müssen identisch heißen. Ansonsten werden keine Teilleistungen nach SchILD importiert.",
            Datentyp = Global.Datentyp.ListString,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },        
        ["VolleStelle"] = new KonfigMeta
        {
            Key = "VolleStelle",
            DefaultValue = "25,5",
            Aufforderung = "",
            Hinweise = $"Geben Sie an, wie viele Stunden einer [{Global.GetColor(Global.ColorInfoBox)}]vollen Stelle[/] entsprechen. Das ist wichtig, um die Altersermäßigung berechnen zu können. Beispiel: [{Global.GetColor(Global.ColorZahlen)}]25,5[/].",
            Datentyp = Global.Datentyp.Float,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["WikiUrl"] = new KonfigMeta
        {
            Key = "WikiUrl",
            DefaultValue = "https://wiki.berufskolleg-borken.de/xmlrpc.php",
            Aufforderung = "",
            Hinweise = $"Geben Sie die [{Global.GetColor(Global.ColorInfoBox)}]URL[/] zum dokuwiki xmlrpc an.",
            Datentyp = Global.Datentyp.Url,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },        
        ["WikiJsonUser"] = new KonfigMeta
        {
            Key = "WikiJsonUser",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie den [{Global.GetColor(Global.ColorInfoBox)}]Benutzernamen[/] für den Zugriff auf das Wiki JSON an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["WikiJsonUserKennwort"] = new KonfigMeta
        {
            Key = "WikiJsonUserKennwort",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Kennwort[/] für den Zugriff auf das Wiki JSON an.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["WikiSprechtagKleineAenderung"] = new KonfigMeta
        {
            Key = "WikiSprechtagKleineAenderung",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Handelt es sich um eine [{Global.GetColor(Global.ColorInfoBox)}]kleine Änderung[/]? Kleine Änderungen erzeugen keine neue Version (j/n)",
            Datentyp = Global.Datentyp.JaNein,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["Zeugnisdatum"] = new KonfigMeta
        {
            Key = "Zeugnisdatum",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Geben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Zeugnisdatum[/] an. Das kann später in SchILD (mit einem Gruppenprozess) erneut geändert werden.",
            Datentyp = Global.Datentyp.DateTime,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["ZeugnisUrl"] = new KonfigMeta
        {
            Key = "ZeugnisUrl",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Soll eine [{Global.GetColor(Global.ColorInfoBox)}]bestimmte Webseite[/] geöffnet werden, um die verschlüsselte(n) Datei(en) dort hochzuladen?\nFalls Sie keine Seite öffnen wollen, geben Sie ein Leerzeichen ein.",
            Datentyp = Global.Datentyp.DateTime,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["ZeugnisPasswort"] = new KonfigMeta
        {
            Key = "ZeugnisPasswort",
            DefaultValue = "",
            Aufforderung = "",
            Hinweise = $"Mit welchem [{Global.GetColor(Global.ColorInfoBox)}]Kennwort[/] sollen die Dateien verschlüsselt werden?",
            Datentyp = Global.Datentyp.DateTime,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["ZipKennwort"] = new KonfigMeta
        {
            Key = "ZipKennwort",
            DefaultValue = Environment.GetEnvironmentVariable("ZIP_KENNWORT") ?? "",
            Aufforderung = "",
            Hinweise = $"Die Datei wird nun gezippt.\nGeben Sie das [{Global.GetColor(Global.ColorInfoBox)}]Kennwort[/] ein, mit dem Sie die Zip-Datei verschlüsseln wollen. Geben Sie ein Leerzeichen ein, wenn kein Kennwort gesetzt werden soll.",
            Datentyp = Global.Datentyp.String,
            InGrundeinstellungAbfragen = true,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        },
        ["ZustimmungLizenz"] = new KonfigMeta
        {
            Key = "ZustimmungLizenz",
            DefaultValue = "nein",
            Aufforderung = $"Sie müssen der Lizenz mit [bold {Global.GetColor(Global.ColorÜberschrift)}]ja[/] zustimmen",
            Hinweise = $"[bold {Global.GetColor(Global.ColorÜberschrift)}]BKB-Tool[/] steht unter der GNU General Public License Version 3 (GPLv3). " +
            "Die GPLv3 ist eine freie Softwarelizenz, die es Ihnen erlaubt, die Software zu verwenden, zu modifizieren und weiterzugeben, solange Sie die Bedingungen der Lizenz einhalten. " +
            "Die wichtigsten Bedingungen dieser Lizenz sind:\n" +
            $"[{Global.GetColor(Global.ColorHinweise)}]Freiheit zur Nutzung, Änderung und Weiterverbreitung:[/] Sie dürfen diese Software frei verwenden, anpassen und weitergeben, solange alle abgeleiteten Werke ebenfalls unter der GPLv3 stehen.\n" +
            $"[{Global.GetColor(Global.ColorHinweise)}]Keine Garantie:[/] Diese Software wird \"wie sie ist\" bereitgestellt, ohne jede ausdrückliche oder stillschweigende Gewährleistung, insbesondere ohne Garantie auf Fehlerfreiheit oder Eignung für einen bestimmten Zweck.\n" +
            $"[{Global.GetColor(Global.ColorHinweise)}]Keine Haftung:[/] Der Entwickler haftet nicht für direkte oder indirekte Schäden, Datenverluste oder andere Konsequenzen, die durch die Nutzung oder Fehlfunktion dieser Software entstehen.\n" +
            $"[{Global.GetColor(Global.ColorHinweise)}]Verwendung auf eigene Gefahr:[/] Die Nutzung erfolgt ausschließlich auf eigenes Risiko.\n" +
            $"[{Global.GetColor(Global.ColorHinweise)}]Vollständige Lizenz unter:[/] [lightskyblue3_1 link=https://www.gnu.org/licenses/gpl-3.0.de.html]https://www.gnu.org/licenses/gpl-3.0.de.html[/].",
            Datentyp = Global.Datentyp.JaNein,
            InGrundeinstellungAbfragen = false,
            InitialAbfragen = false,
            NurBeiDiesenSchulnummern = Global.SchulnummernJedermann
        }
    };
}