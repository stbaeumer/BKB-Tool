using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Xml;
using Spectre.Console.Rendering;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using Color = Spectre.Console.Color;
using System.Diagnostics;
using CookComputing.XmlRpc;
using System.Dynamic;
using System.Reflection;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Pdf;
using iText.Layout;

//using iText.Layout.Element;
namespace Common;

#pragma warning disable CS8603 // Mögliche Null-Verweis-Rückgabe
#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise null-Objekts.
#pragma warning disable CS8604 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8600 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8618 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8619 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0219 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8625


public class Datei : List<dynamic>
{
    public string Endung { get; set; } = null!;
    public string Delimiter { get; set; } = null!;
    public char Quote { get; set; }
    public Encoding Encoding { get; set; }
    public bool ShouldAllQuote { get; private set; }
    public bool ShouldQuote { get; set; }
    public List<string> Importhinweise { get; set; }
    public Lehrers Lehrers { get; set; } = null!;
    public bool DarfLeerSein { get; set; } = false;
    public bool HasHeader { get; set; }
    public string Name { get; set; } = null!;
    public string UnterordnerUndDateiname { get; set; } = null!;
    private Datei Vergleichsdatei { get; } = null!;
    public string[] Hinweise { get; } = null!;
    public DateTime Erstelldatum { get; set; }
    public bool Vorhanden { get; set; }
    public string Beschreibung { get; } = null!;
    public Func<Datei, List<dynamic>> Funktion { get; } = null!; // Rückgabewert definiert
    public List<string> KlassenNamen { get; set; } = null!;
    public Students IStudents { get; set; } = null!;

    /// <summary>
    /// Am Ende werden neu erstellte Dateien mit vorhandenen verglichen. Diese Eigenschaften werden für den Dateivergleich ignoriert.
    /// </summary>
    private string[] ZuIgnorierendeEigenschaften { get; set; } = null!;

    public string? AbsoluterPfad { get; set; } = "";
    public List<Action<Datei>> Funktionen { get; set; }
    public IConfiguration Konfiguration { get; private set; }
    public string Dateiname { get; set; } = null!;
    public string ZipPfad { get; private set; }
    public string Fehlermeldung { get; set; }
    public List<Global.Modus> Modus { get; set; }
    public string Ordner { get; internal set; }
    public bool IstOptional { get; internal set; }
    public bool Nur177659 { get; internal set; }
    /// <summary>
    /// KeyAttribute, anhand derer die Datei mit anderen Dateien verglichen wird.    
    /// </summary>
    public string[] AnhandDieserSchlüsselAttributeWirdVerglichen { get; set; }
    public string[] DieseAttributeWerdenBeimVergleichIgnoriert { get; set; }
    public object Value1 { get; }
    public string V1 { get; }
    public char V2 { get; }
    public UTF8Encoding UTF8Encoding { get; }
    public bool V3 { get; }
    public object Value2 { get; }
    public DokuwikiZugriff DokuwikiZugriff { get; private set; }
    public string UrlMitte { get; set; }
    public string UrlRechts { get; set; }

    //public Datei(string name, Global.Modus modus, string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert)
    public Datei(
        string name,        
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise)
    {
        Name = name;
        Dateiname = Path.GetFileName(name);
        AbsoluterPfad = name;        
        AnhandDieserSchlüsselAttributeWirdVerglichen = anhandDieserAttributeWirdVerglichen;
        DieseAttributeWerdenBeimVergleichIgnoriert = dieseAttributeWerdenBeimVergleichIgnoriert;
        UnterordnerUndDateiname = name;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
    }
    
    public Datei(
        string name,        
        List<Action<Datei>> funktionen,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise)
    {
        Name = name;
        Dateiname = Path.GetFileName(name);
        AbsoluterPfad = name;        
        Funktionen = funktionen;
        AnhandDieserSchlüsselAttributeWirdVerglichen = anhandDieserAttributeWirdVerglichen;
        DieseAttributeWerdenBeimVergleichIgnoriert = dieseAttributeWerdenBeimVergleichIgnoriert;
        UnterordnerUndDateiname = name;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
    }

    public Datei(
        string name,
        List<Action<Datei>> funktionen,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise)
    {
        Name = name;
        Funktionen = funktionen;
        Dateiname = Path.GetFileName(name);
        AbsoluterPfad = name;
        UnterordnerUndDateiname = name;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
    }

    public Datei(
        string name,
        List<Action<Datei>> funktionen,
        IConfiguration configuration)
    {
        Name = name;
        AbsoluterPfad = name;
        Funktionen = funktionen;
        Konfiguration = configuration;
    }

    public Datei(string zieldateiname, Datei vergleichsdatei)
    {
        AbsoluterPfad = zieldateiname;
        Vergleichsdatei = vergleichsdatei;
    }

    public Datei()
    {
    }

    public Datei(List<dynamic> records)
    {
        AddRange(records);
    }

    public Datei(string absoluterPfad)
    {
        AbsoluterPfad = absoluterPfad;        
    }

    public Datei(Datei datei)
    {
        UnterordnerUndDateiname = datei.UnterordnerUndDateiname;
        Erstelldatum = datei.Erstelldatum;
        ZuIgnorierendeEigenschaften = datei.ZuIgnorierendeEigenschaften;
        Hinweise = datei.Hinweise;
    }

    public Datei(string name, string beschreibung, string[] hinweise, string[] zuIgnorierendeEigenschaften,
        bool hasHeader, Func<Datei, List<dynamic>> funktion, string endung = "*.dat", string delimiter = "|", bool darfLeerSein = false)
    {
        Name = name;
        Dateiname = Path.GetFileName(name);
        Beschreibung = beschreibung;
        Hinweise = hinweise;
        ZuIgnorierendeEigenschaften = zuIgnorierendeEigenschaften;
        HasHeader = hasHeader;
        Funktion = funktion;
        Endung = endung;
        Delimiter = delimiter;
        DarfLeerSein = darfLeerSein;
    }

    public Datei(
        string absoluterPfad,
        IConfiguration configuration,
        List<Action<Datei>> funktionen,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote,
        Lehrers lehrers,
        List<string> importhinweise) : this(absoluterPfad)
    {
        AbsoluterPfad = absoluterPfad;
        Konfiguration = configuration;
        Funktionen = funktionen;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
        Lehrers = lehrers;
    }

    public Datei(
        string absoluterPfad,
        string[] anhandDieserAttributeWirdVerglichen,
        string[] dieseAttributeWerdenBeimVergleichIgnoriert,
        string delimiter, char quote, Encoding encoding, bool shouldAllQuote,
        List<Action<Datei>> funktionen,
        List<string> importhinweise) : this(absoluterPfad)
    {
        AbsoluterPfad = absoluterPfad;
        Funktionen = funktionen;
        AnhandDieserSchlüsselAttributeWirdVerglichen = anhandDieserAttributeWirdVerglichen;
        DieseAttributeWerdenBeimVergleichIgnoriert = dieseAttributeWerdenBeimVergleichIgnoriert;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
    }

    public Datei(string absoluterPfad, IConfiguration configuration, string delimiter, char quote, Encoding encoding, bool shouldAllQuote, Lehrers lehrers, List<string> importhinweise) : this(absoluterPfad)
    {
        AbsoluterPfad = absoluterPfad;
        Konfiguration = configuration;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
        Lehrers = lehrers;
    }

    public List<dynamic> Filtern(Students students, Klassen klassen)
    {
        IStudents = students;
        if(klassen != null)
        {
            KlassenNamen = klassen.Where(x => !string.IsNullOrEmpty(x.Name)).Select(x => x.Name).ToList();
        }

        return Funktion?.Invoke(this) ?? new List<dynamic>(); // Falls `Funktion` null ist, leere Liste zurückgeben
    }

    public List<dynamic> FilternDatDatei()
    {
        if(IStudents.Count == 0){
            return this;
        }

        var liste = new List<dynamic>();
        
        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (IStudents.Any(student => student.Nachname == dict["Nachname"].ToString() &&
                                         student.Vorname == dict["Vorname"].ToString() &&
                                         student.Geburtsdatum == dict["Geburtsdatum"].ToString()))
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilterExportLessons()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            
            if (dict["klassen"].ToString().Split('~').Any(klasse => KlassenNamen.Contains(klasse)))
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilterGost()
    {
        throw new NotImplementedException();
    }

    public List<dynamic> FilterKurse()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilterDatumsAusAtlantis()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternAdressenAtlantis()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (IStudents.Where(student =>
                    student.Nachname == dict["Schüler: Nachname"].ToString() &&
                    student.Vorname == dict["Schüler: Vorname"].ToString() &&
                    student.Geburtsdatum == dict["Schüler: Geburtsdatum"].ToString())
                .Any())
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternTermineKollegium()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternTermineFhr()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternTermineVerwaltung()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternTermineBeruflichesGymnasium()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternAtlantisZeugnisseNoten()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (IStudents.Any(x => x.Nachname == dict["Field1"].ToString().Replace("'", "") &&
                                   x.Geburtsdatum == dict["Field3"].ToString().Replace("'", "")))
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternSchildKlassen()
    {
        if (KlassenNamen.Count() == 0)
        {
            return this;
        }
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;

            if (KlassenNamen.Any(k => k == dict["InternBez"].ToString()))
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternSchildFaecher()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternWebuntisStudent()
    {
        if(IStudents.Count == 0)
        {
            return this;
        }
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if(true)
            /*if (IStudents.Where(student =>
                    student.Nachname == dict["longName"].ToString() &&
                    student.Vorname == dict["foreName"].ToString() &&
                    student.Geburtsdatum == dict["birthDate"].ToString())
                .Any())*/
            {
                liste.Add(rec);
            }
            else
            {
                string s = "";
            }   
        }

        return liste;
    }

    public List<dynamic> FilternMarksPerLessons()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (KlassenNamen.Any(k => k == dict["Klasse"].ToString()))
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternAbsencePerLessons()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (KlassenNamen.Any(k => k == dict["Klasse"].ToString()))
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public List<dynamic> FilternStudentgroupStudents()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public void Erstellen()
    {
        if (AbsoluterPfad == null)
        {
            // Wenn der Pfad leer ist oder die Liste leer ist, wird die Datei nicht erstellt.
            var panel = new Panel($"[red]Datei nicht erstellt: Pfad leer: [/]{AbsoluterPfad}")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Red);
            
            AnsiConsole.Write(panel);
            return;
        }else if (Count == 0)
        {
            // Wenn der Pfad leer ist oder die Liste leer ist, wird die Datei nicht erstellt.
            var panel = new Panel($"[red]Datei nicht erstellt (0 Zeilen): [/]{AbsoluterPfad}")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Red);
            
            AnsiConsole.Write(panel);
            return;
        }
        /*
        new UTF8Encoding(true), // UTF-8 mit BOM
        new UTF8Encoding(false),   // UTF-8 ohne BOM
        Encoding.Default,          // ANSI        
        */

        // Wenn der Dateiname auf XML endet, wird die Datei als XML-Datei erstellt.
        if (AbsoluterPfad.ToLower().EndsWith(".xml"))
        {
            // Es wird aus der Liste eine XML-Datei erstellt.
            try
            {
                var xmlDoc = new XmlDocument();
                var xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "iso-8859-1", null);
                xmlDoc.AppendChild(xmlDeclaration);
                var root = xmlDoc.CreateElement("Leserdaten");
                xmlDoc.AppendChild(root);

                foreach (var record in this)
                {
                    var recordDict = record as IDictionary<string, object>;
                    if (recordDict != null)
                    {
                        var element = xmlDoc.CreateElement("Leser");
                        foreach (var kvp in recordDict)
                        {
                            var childElement = xmlDoc.CreateElement(kvp.Key);
                            childElement.InnerText = kvp.Value?.ToString() ?? string.Empty;
                            element.AppendChild(childElement);
                        }
                        root.AppendChild(element);
                    }
                }

                File.Delete(AbsoluterPfad);
                xmlDoc.Save(AbsoluterPfad);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fehler beim Schreiben der XML-Datei: {ex.Message}");
            }
            finally
            {
                Global.ZeileSchreiben(AbsoluterPfad, "", ConsoleColor.White, ConsoleColor.Blue);
            }
        }
                else if (AbsoluterPfad.ToLower().EndsWith(".html"))
        {
            try
            {
                var sb = new StringBuilder();
        
                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("<html>");
                sb.AppendLine("<head>");
                sb.AppendLine("  <meta charset=\"utf-8\" />");
                sb.AppendLine("  <title>Export</title>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.AppendLine("  <table border=\"1\">");
        
                var firstRecord = this.FirstOrDefault() as IDictionary<string, object>;
                if (firstRecord != null)
                {
                    // Kopfzeile
                    sb.AppendLine("    <tr>");
                    foreach (var header in firstRecord.Keys)
                    {
                        var adjustedHeader = header
                            .Replace("DOPPELPUNKT", ":")
                            .Replace("PUNKT", ".")
                            .Replace("MINUS", "-")
                            .Replace("ZWEI", "2")
                            .Replace("EINS", "1")
                            .Replace("UNTERSTRICH", "_")
                            .Replace("SLASH", "/")
                            .Replace("LEERZEICHEN", " ")
                            .Replace("KLAMMERAUF", "(")
                            .Replace("KLAMMERZU", ")");
        
                        sb.AppendLine($"      <th>{System.Net.WebUtility.HtmlEncode(adjustedHeader)}</th>");
                    }
                    sb.AppendLine("    </tr>");
        
                    // Datenzeilen
                    foreach (var record in this)
                    {
                        var recordDict = record as IDictionary<string, object>;
                        if (recordDict == null) continue;
        
                        sb.AppendLine("    <tr>");
                        foreach (var header in firstRecord.Keys)
                        {
                            recordDict.TryGetValue(header, out var value);
                            sb.AppendLine($"      <td>{System.Net.WebUtility.HtmlEncode(value?.ToString() ?? string.Empty)}</td>");
                        }
                        sb.AppendLine("    </tr>");
                    }
                }
        
                sb.AppendLine("  </table>");
                sb.AppendLine("</body>");
                sb.AppendLine("</html>");
        
                File.Delete(AbsoluterPfad);
                File.WriteAllText(AbsoluterPfad, sb.ToString(), this.Encoding ?? Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fehler beim Schreiben der HTML-Datei: {ex.Message}");
            }
            finally
            {
                Global.ZeileSchreiben(AbsoluterPfad, "", ConsoleColor.White, ConsoleColor.Blue);
            }
        }
        else
        {
            try
            {
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = this.Delimiter,
                    Quote = this.Quote,
                    ShouldQuote = args => this.ShouldAllQuote
                };

                File.Delete(AbsoluterPfad);

                if (this != null && this.Any())
                {
                    using var writer = new StreamWriter(AbsoluterPfad, false, this.Encoding);
                    using var csv = new CsvWriter(writer, config);

                    // Header manuell extrahieren
                    var firstRecord = this[0] as IDictionary<string, object>;
                    var adjustedHeaders = new List<string>();

                    foreach (var header in firstRecord.Keys)
                    {
                        // Anpassen der Header
                        var adjustedHeader = header
                            .Replace("DOPPELPUNKT", ":")
                            .Replace("PUNKT", ".")
                            .Replace("MINUS", "-")
                            .Replace("ZWEI", "2")
                            .Replace("EINS", "1")
                            .Replace("UNTERSTRICH", "_")
                            .Replace("SLASH", "/")
                            .Replace("LEERZEICHEN", " ")
                            .Replace("KLAMMERAUF", "(")
                            .Replace("KLAMMERZU", ")");

                        adjustedHeaders.Add(adjustedHeader);
                    }

                    // Schreiben der angepassten Header
                    foreach (var header in adjustedHeaders)
                    {
                        csv.WriteField(header);
                    }

                    csv.NextRecord();

                    // Schreiben der Datensätze
                    foreach (var record in this)
                    {
                        var recordDict = record as IDictionary<string, object>;
                        foreach (var value in recordDict.Values)
                        {
                            csv.WriteField(value);
                        }

                        csv.NextRecord();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Fehler beim Schreiben der Datei: {ex.Message}");
            }
            finally
            {
                var rechteSeite = Importhinweise != null && Importhinweise.Any() ? string.Join("\n", Importhinweise) : "";
                Global.ZeileSchreiben(AbsoluterPfad, rechteSeite, ConsoleColor.White, ConsoleColor.Blue);
            }
        }
    }

        public void SpectreTabelleErstellen()
    {
        if (string.IsNullOrWhiteSpace(AbsoluterPfad))
        {
            AnsiConsole.Write(new Panel($"[red]Datei nicht erstellt: Pfad leer: [/]{AbsoluterPfad}")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Red));
            return;
        }
    
        if (Count == 0)
        {
            AnsiConsole.Write(new Panel($"[Spectre.Console.Color.Orange1]Datei nicht erstellt (0 Zeilen): [/]{AbsoluterPfad}")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Red));
            return;
        }
    
        try
        {
            var firstRecord = this[0] as IDictionary<string, object>;
            if (firstRecord == null || firstRecord.Count == 0)
                return;
    
            var originalHeaders = firstRecord.Keys.ToList();
            var adjustedHeaders = originalHeaders
                .Select(header => header
                    .Replace("DOPPELPUNKT", ":")
                    .Replace("PUNKT", ".")
                    .Replace("MINUS", "-")
                    .Replace("ZWEI", "2")
                    .Replace("EINS", "1")
                    .Replace("UNTERSTRICH", "_")
                    .Replace("SLASH", "/")
                    .Replace("LEERZEICHEN", " ")
                    .Replace("KLAMMERAUF", "(")
                    .Replace("KLAMMERZU", ")"))
                .ToList();
    
            var table = new Table().Border(TableBorder.Ascii);
            foreach (var header in adjustedHeaders)
                table.AddColumn(header);
    
            foreach (var record in this)
            {
                var recordDict = record as IDictionary<string, object>;
                if (recordDict == null) continue;
    
                var row = originalHeaders
                    .Select(h => recordDict.TryGetValue(h, out var value) ? value?.ToString() ?? "" : "")
                    .ToArray();
    
                table.AddRow(row);
            }
    
            var sw = new StringWriter();
            var console = AnsiConsole.Create(new AnsiConsoleSettings
            {
                Out = new AnsiConsoleOutput(sw),
                Ansi = AnsiSupport.No,
                ColorSystem = ColorSystemSupport.NoColors
            });
    
            AnsiConsole.Write(table);
    
            if (File.Exists(AbsoluterPfad))
                File.Delete(AbsoluterPfad);
    
            File.WriteAllText(AbsoluterPfad, sw.ToString(), this.Encoding ?? Encoding.UTF8);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Fehler beim Schreiben der Tabelle: {ex.Message}");
        }
        finally
        {
            var rechteSeite = Importhinweise != null && Importhinweise.Any() ? string.Join("\n", Importhinweise) : "";
            Global.ZeileSchreiben(AbsoluterPfad, rechteSeite, ConsoleColor.White, ConsoleColor.Blue);
        }
    }


    public Datei? VergleichenUndFiltern(Dateien quelldateien)
    {
        var neueDatei = new Datei(AbsoluterPfad);
        bool skipProcessing = false;
        
        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Vergleichen & Filtern ...", ctx =>
        {
            var dateiendung = Path.GetExtension(AbsoluterPfad);
            var tableRows = new List<Text>();
            var table = TableErstellen(
                $"Vergleich von [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.GetFileName(AbsoluterPfad)}[/] mit [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.GetFileName(AbsoluterPfad)}[/]",
                AnhandDieserSchlüsselAttributeWirdVerglichen);

            var vorhandeneRec = GetVorhandeneRec(quelldateien);
            if (vorhandeneRec == null || vorhandeneRec.Count == 0)
            {
                skipProcessing = true;
                return;
            }

            // Wenn in der vorhandenen eine Spalte namens Nachname existiert und die Dateiendung .dat ist,
            // dann muss in jeder Zeile der Nachname um #Klasse ergänzt werden.
            if (AnhandDieserSchlüsselAttributeWirdVerglichen.Contains("Nachname") && dateiendung == ".dat")
            {
                // In jeder Zeile den Nachnamen um #Klasse ergänzen
                foreach (var vRec in vorhandeneRec)
                {
                    var vorhandeneDict = (IDictionary<string, object>)vRec;
                    if (vorhandeneDict.ContainsKey("Nachname"))
                    {
                        var nachname = vorhandeneDict["Nachname"].ToString();
                        if (vorhandeneDict.ContainsKey("Klasse"))
                        {
                            nachname += "#" + vorhandeneDict["Klasse"];
                        }
                        vorhandeneDict["Nachname"] = nachname;
                    }
                }
            }

            // Die Tabelle soll begrenzt werden. Falls mehr als maxRows Zeilen gefunden werden, wird eine weitere Zeile mit "..." eingefügt.

            var maxRows = 100;
            var rows = 0;

            // Für jede neue Zeile ...
            foreach (var neueRec in this)
            {
                var neueDict = (IDictionary<string, object>)neueRec;

                // ... wird geprüft, ob es eine vorhandene Zeile gibt, die auf die Schlüsselattribute matcht.
                var zeileMitIdentischenSchlüsselattributen = GetZeileMitIdentischenSchlüsselattributen(vorhandeneRec, neueDict, AnhandDieserSchlüsselAttributeWirdVerglichen);

                // Fall1: Wenn keine Zeile in den Vergleichsattributen auf die vorhandenen matcht, wird die Zeile neu angelegt.
                if (zeileMitIdentischenSchlüsselattributen == null) { neueDatei.Add(neueRec); continue; } // und die Schleife übersprungen

                // Fall2: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht, werden abweichende Nicht-Schlüssel-Attributwerte gesucht, ...
                var nichtIdentischeSonstigeAttribute = GetNichtIdentischeSonstigeAttribute(zeileMitIdentischenSchlüsselattributen, neueDict);

                // Fall2a: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht abweichen, ...
                if (nichtIdentischeSonstigeAttribute.Count == 0) continue; // ... überspringe den Rest der Schleife

                // Fall2b: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht identisch sind, ...
                if (nichtIdentischeSonstigeAttribute.Count <= 0) continue;

                // Für die ersten 2 abweichenden Attribute dieser Zeile wird eine Zeile in der Tabelle erstellt.
                for (int i = 0; i < nichtIdentischeSonstigeAttribute.Count; i++)
                {
                    if (rows > maxRows) break; // Wenn die maximale Anzahl an Zeilen erreicht ist, breche die Schleife ab.                    
                    if (rows == maxRows)
                        table.AddRow(new Text("..."), new Text("..."), new Text("..."), new Text("..."));
                    if (rows < maxRows)
                        table.AddRow(RenderZeile(neueDict, vorhandeneRec, nichtIdentischeSonstigeAttribute[i], zeileMitIdentischenSchlüsselattributen));
                    rows++;
                }

                neueDatei.Add(neueRec);
            }

            if (rows == 0)
            {
                table.AddRow(new Text("keine Änderungen, nichts anzuzeigen"), new Text("..."), new Text("..."), new Text("..."));
            }

            AnsiConsole.Write(table);
            AnsiConsole.Write(new Align(
                new Text($"Summe: {neueDatei.Count} "),
                HorizontalAlignment.Right,
                VerticalAlignment.Top));
        });

        if (skipProcessing)
            return this;

        // Entferne alle Zeilen aus this und ersetzte durch neueDatei.
        Clear();
        AddRange(neueDatei);
        return this;
    }

    private IEnumerable<IRenderable> RenderZeile(IDictionary<string, object> neueDict, List<dynamic> vorhandeneRec, string nichtIdentischesSonstigesAttribut, IDictionary<string, object> zeileMitIdentischenVergleichsattributen)
    {
        return
        [
            new Text(GetLinkeSeite(neueDict, AnhandDieserSchlüsselAttributeWirdVerglichen)),
            new Text(nichtIdentischesSonstigesAttribut.Replace("PUNKT", ".").Replace("LEERZEICHEN", " ").Replace("MINUS", "-").Replace("UNTERSTRICH", "_").Replace("SCHRÄGSTRICH", "/")),
            new Text(GetAlterWert(zeileMitIdentischenVergleichsattributen, nichtIdentischesSonstigesAttribut)),
            new Text(GetNeuerWert(neueDict, nichtIdentischesSonstigesAttribut))
        ];
    }

    private string GetAlterWert(IDictionary<string, object> zeileMitIdentischenVergleichsattributen, string nichtIdentischesSonstigesAttribut)
    {
        // Suche aus allen Spalten aus zeileMitIdentischenVergleichsattributen denjenigen Spaltenwert, dessen Spaltenname nichtIdentischesSonstigesAttribut entspricht.
        if (
            zeileMitIdentischenVergleichsattributen != null &&
            zeileMitIdentischenVergleichsattributen.TryGetValue(nichtIdentischesSonstigesAttribut, out var value)
            )
        {
            return value?.ToString() ?? string.Empty;
        }
        return string.Empty;
    }

    private Table TableErstellen(string title, string[] anhandDieserSchlüsselAttributeWirdVerglichen)
    {
        var table = new Table();
        table.Expand();
        table.Border(TableBorder.Rounded);
        if (!string.IsNullOrEmpty(title))
            table.Title = new TableTitle(title);
        
        table.Expand();
        table.AddColumn((string.Join(", ", anhandDieserSchlüsselAttributeWirdVerglichen)).Replace("PUNKT", ".").Replace("LEERZEICHEN", " ").Replace("MINUS", "-").Replace("UNTERSTRICH", "_").Replace("SCHRÄGSTRICH", "/"));
        table.AddColumn("Attribut");
        table.AddColumn("alt");
        table.AddColumn("neu");
        return table;
    }

    /// <summary>
    /// Die vorhandene hat denselben Namen wie die zieldatei wird im Downloadsordner gesucht
    /// </summary>
    /// <param name="quelldateien"></param>
    /// <returns></returns>
    private List<dynamic> GetVorhandeneRec(Dateien quelldateien)
    {
        var vorhandeneRec = new List<dynamic>();

        foreach (var vorhandeneDatei in quelldateien)
        {
            // Bei .dat-Dateien matcht der Dateiname
            if (Path.GetFileName(vorhandeneDatei.AbsoluterPfad) == Path.GetFileName(AbsoluterPfad))
            {
                vorhandeneRec = vorhandeneDatei;
                break; // Schleife abbrechen, wenn die Datei gefunden wurde
            }
            // bei .csv-Dateien matcht der Dateiname vor dem ersten Unterstrich
            if (
                Path.GetFileName(AbsoluterPfad).Contains("_") &&
                Path.GetFileName(vorhandeneDatei.AbsoluterPfad).Split('_')[0] == Path.GetFileName(AbsoluterPfad).Split('_')[0])
            {
                vorhandeneRec = vorhandeneDatei;
                break; // Schleife abbrechen, wenn die Datei gefunden wurde
            }
        }
        return vorhandeneRec ?? new List<dynamic>(); // Rückgabe der gefundenen Datei oder leere Liste, wenn keine Datei gefunden wurde
    }

    private static IDictionary<string, object> GetVorhandeneZeile(IDictionary<string, object> neueDict,
        List<dynamic> vorhandene, string[] anhandDieserAttributeWirdVerglichen)
    {
        foreach (var rec in vorhandene)
        {
            var vorhDict = (IDictionary<string, object>)rec;

            // Prüfen, ob alle relevanten Schlüssel-Werte-Paare übereinstimmen
            bool isMatch = anhandDieserAttributeWirdVerglichen.All(attr =>
                vorhDict.ContainsKey(attr) &&
                neueDict.ContainsKey(attr) &&
                Equals(vorhDict[attr], neueDict[attr]));

            if (isMatch)
            {
                return vorhDict; // Gefundene übereinstimmende Zeile zurückgeben
            }
        }

        return null; // Falls keine passende Zeile gefunden wurde
    }


    private string GetNeuerWert(IDictionary<string, object> neueDict, string nichtIdentischesSonstigesAttribut)
    {
        var x = "";

        x = neueDict.TryGetValue(nichtIdentischesSonstigesAttribut, out var neuerWert)
            ? (neuerWert?.ToString()?.Length > 20
                ? neuerWert.ToString().Substring(0, 17) + "..."
                : neuerWert?.ToString() ?? string.Empty)
            : string.Empty;

        if (x.Length > 0)
            return x;

        // Wenn nichts matcht, versuche nochmal mit ersetzten Sonderzeichen

        nichtIdentischesSonstigesAttribut = nichtIdentischesSonstigesAttribut.Replace(".", "PUNKT").Replace(" ", "LEERZEICHEN").Replace("-", "MINUS").Replace("_", "UNTERSTRICH").Replace("/", "SLASH");

        return neueDict.TryGetValue(nichtIdentischesSonstigesAttribut, out var neuerWert1)
            ? (neuerWert1?.ToString()?.Length > 20
                ? neuerWert1.ToString().Substring(0, 17) + "..."
                : neuerWert1?.ToString() ?? string.Empty)
            : string.Empty;
            


    }

    private string GetLinkeSeite(IDictionary<string, object> neueDict, string[] anhandDieserAttributeWirdVerglichen)
    {
        var linkeSeite = "";

        for (int i = 0; i < anhandDieserAttributeWirdVerglichen.Length; i++)
        {
            var wert = Global.PrüfeAufNullOrEmpty(neueDict, anhandDieserAttributeWirdVerglichen[i]);
            if (wert.Length > 0)
            {
                linkeSeite += wert + ", ";
                if (linkeSeite.Length > 60)
                {
                    linkeSeite = linkeSeite.Substring(0, 60) + "...";
                    break;
                }
            }
        }

        return linkeSeite.TrimEnd(',').TrimEnd(' ').TrimEnd(',').TrimEnd(' ');
    }

    private List<string> GetNichtIdentischeSonstigeAttribute(IDictionary<string, object> vorhDict,
        IDictionary<string, object> neueDict)
    {
        var sonderzeichen = false;
        // Es wird sichergestellt, dass in neueDict und vorhandeneDict die gleichen Schlüssel verwendet werden.
        foreach (var key in neueDict.Keys)
        {
            if(key.Contains(".") || key.Contains(" ") || key.Contains("-") || key.Contains("_") || key.Contains("/"))
            {
                sonderzeichen = true;
                break;
            }
        }

        List<string> nichtIdentischeSonstige = new List<string>();
        foreach (var key in vorhDict.Keys)
        {
            var k = key;
            // Wenn die neueDict keine Sonderzeichen enthält, dann wird auch bei vorhandeneDict die Sonderzeichen ersetzt.
            if (!sonderzeichen)
                k = key.Replace(".", "PUNKT").Replace(" ", "LEERZEICHEN").Replace("-", "MINUS").Replace("_", "UNTERSTRICH").Replace("/", "SCHRÄGSTRICH");

            // Die Felder, die mit Field beginnen, sind nicht relevant
            if (DieseAttributeWerdenBeimVergleichIgnoriert.Contains(key)) continue;
            if (AnhandDieserSchlüsselAttributeWirdVerglichen.Contains(key)) continue; // Die Vergleichsattribute werden nicht berücksichtigt

            // Es wird geprüft, ob das Dictionary neueDict einen Eintrag mit dem Schlüssel key enthält.
            // Falls nicht (!), wird die aktuelle Iteration der Schleife übersprungen (continue).
            // Falls doch, wird der Wert zu diesem Schlüssel in die Variable value geschrieben und der Code läuft weiter.            
            if (!neueDict.TryGetValue(k, out var value)) continue;
            // Wenn die Werte identisch sind, wird die aktuelle Iteration der Schleife übersprungen (continue).
            if (vorhDict[key].Equals(value?.ToString())) continue;
            // Z.B. bei Fehlstunden bleibt die neue Zelle leer. In der alten steht 0
            if (vorhDict[key].ToString() == "0" && neueDict[k].ToString() == "") continue;
            nichtIdentischeSonstige.Add(key);
        }

        return nichtIdentischeSonstige;
    }

    public IDictionary<string, object> GetZeileMitIdentischenSchlüsselattributen(
        List<dynamic> vorhandene,
        IDictionary<string, object>
        neueDict,
        string[] anhandDieserAttributeWirdVerglichen)
    {
        foreach (var vorhDict in vorhandene.Select(vorhRec => (IDictionary<string, object>)vorhRec))
        {
            bool match = true;
            foreach (var key in anhandDieserAttributeWirdVerglichen)
            {
                var neueDictWert = "";
                if ((neueDict.ContainsKey(key)) && (neueDict[key] != null))
                {
                    neueDictWert = neueDict[key].ToString().Split('#')[0];
                }

                var vorhDictWert = "";
                if (vorhDict.ContainsKey(key) && vorhDict[key] != null)
                {
                    vorhDictWert = vorhDict[key].ToString().Split('#')[0];
                }

                if (neueDictWert != vorhDictWert)
                {
                    match = false;
                    break;
                }
            }
            if (match)
            {
                return vorhDict;
            }
        }
        return null;
    }
  

    public List<dynamic> FilterLehrkraefte()
    {
        var liste = new List<dynamic>();

        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (true)
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    /// <summary>
    /// Liest die Datei ein und füllt die Liste mit den Zeilen.
    /// Wenn die Datei eine PDF ist, wird sie übersprungen.
    /// </summary>
    public void GetZeilen()
    {
        this.Clear();

        if (AbsoluterPfad != null && AbsoluterPfad.EndsWith(".pdf"))
        {
            return;
        }

        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            HeaderValidated = null,
            MissingFieldFound = null,
            HasHeaderRecord = HasHeader,
            Delimiter = Delimiter,
            BadDataFound = null,
            IgnoreBlankLines = true
        };

        using (var reader = new StreamReader(AbsoluterPfad))
        using (var csv = new CsvReader(reader, config))
        {
            foreach (var record in csv.GetRecords<dynamic>())
            {
                var anzahlNichtLeererRecords = 0;

                // Iteriere über die Key-Value-Paare des dynamischen Records
                foreach (var item in (IDictionary<string, object>)record)
                {
                    // Prüfe, ob der Wert nicht leer ist
                    if (item.Value != null && !string.IsNullOrWhiteSpace(item.Value.ToString()))
                    {
                        anzahlNichtLeererRecords++;
                    }
                }

                // Wenn nur eine einzige Zelle Inhalt hat, deutet das auf mehrzeilige Records hin. Weitere Zeilen werden ignoriert.
                if (anzahlNichtLeererRecords > 2)
                {
                    this.Add(record);
                }
            }

            this.AddRange(csv.GetRecords<dynamic>());
        }
    }

    public List<dynamic> FilternKlassenGPU003()
    {
        return this;
    }

    public List<dynamic> FilternKlassenGPU002()
    {
        var listeUnterrichtsIds = new List<string>();

        // Es werden allen IDs herausgesucht, an denen sich eine interessierende Klasse beteiligt.
        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (KlassenNamen.Any(k => k == dict["Field5"].ToString()))
            {
                listeUnterrichtsIds.Add(dict["Field1"].ToString());
            }
        }

        var liste = new List<dynamic>();

        // Es werden allen Zeilen (auch von anderen Klassen) herausgesucht, die über die ID mit den interessierenden Klassen verbunden sind.        
        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (listeUnterrichtsIds.Any(k => k == dict["Field1"].ToString()))
            {
                liste.Add(rec);
            }
        }

        return liste;
    }

    public string? GetAbsoluterPfad()
    {
        return AbsoluterPfad;
    }

    public void FotosZippen(IConfiguration configuration, string kennwort = "", int kompressionsLevel = 0, Students students = null!)
    {
        configuration = Global.Konfig("RotateFotos", Global.Modus.Update, configuration);

        try
        {
            var studentsMitNeuenFotos = students.Where(s => !string.IsNullOrEmpty(s.ZielFotoPfad) && File.Exists(s.ZielFotoPfad)).ToList();
            if(studentsMitNeuenFotos.Count == 0)
            {
                throw new Exception("Es wurden keine Fotos gefunden, die gezippt werden könnten.");
            }

            // Erste Zip-Datei: Webuntis-Kurzname
            var webuntisZipPfad = Path.Combine(Path.GetDirectoryName(AbsoluterPfad), 
                Path.GetFileNameWithoutExtension(AbsoluterPfad) + "_Webuntis-Kurzname.zip");
            
            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Fotos zippen (Webuntis-Kurzname) ...", ctx =>
            {
                using (FileStream zipStream = File.Create(webuntisZipPfad))
                using (ZipOutputStream zip = new ZipOutputStream(zipStream))
                {
                    zip.SetLevel(kompressionsLevel);

                    if (!string.IsNullOrEmpty(kennwort) && kennwort != " ")
                    {
                        zip.Password = kennwort;
                    }

                    foreach (var student in studentsMitNeuenFotos)
                    {
                        var tempDatei = Path.Combine(Path.GetTempPath(), student.MailSchulisch.Split('@')[0] + ".jpg");

                        using (var image = Image.Load(student.ZielFotoPfad))
                        {
                            image.Mutate(x => x.Resize(160, 160));
                            image.Mutate(x => x.Rotate(Convert.ToInt32(configuration["RotateFotos"])));
                            image.Save(tempDatei);
                        }

                        byte[] buffer = new byte[4096];
                        string dateiName = Path.GetFileName(tempDatei);

                        ZipEntry entry = new ZipEntry(dateiName)
                        {
                            DateTime = DateTime.Now,
                            CompressionMethod = CompressionMethod.Stored
                        };

                        zip.PutNextEntry(entry);

                        using (FileStream dateiStream = File.OpenRead(tempDatei))
                        {
                            int bytesRead;
                            while ((bytesRead = dateiStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                zip.Write(buffer, 0, bytesRead);
                            }
                        }
                        
                        File.Delete(tempDatei);
                    }

                    try
                    {
                        zip.CloseEntry();
                    }
                    catch
                    {
                        throw new Exception("Fehler beim Erstellen des Zip-Archivs (Webuntis). Möglicherweise wurden keine Fotos gefunden.");
                    }
                    
                    zip.IsStreamOwner = true;
                }
            });
            Global.ZeileSchreiben("Fotos gezippt (Webuntis-Kurzname)", webuntisZipPfad, ConsoleColor.Green, ConsoleColor.White);

            // Zweite Zip-Datei: Geevoo-Mail
            var geevooZipPfad = Path.Combine(Path.GetDirectoryName(AbsoluterPfad), 
                Path.GetFileNameWithoutExtension(AbsoluterPfad) + "_Geevoo-Mail.zip");
            
            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Fotos zippen (Geevoo-Mail) ...", ctx =>
            {
                using (FileStream zipStream = File.Create(geevooZipPfad))
                using (ZipOutputStream zip = new ZipOutputStream(zipStream))
                {
                    zip.SetLevel(kompressionsLevel);

                    if (!string.IsNullOrEmpty(kennwort) && kennwort != " ")
                    {
                        zip.Password = kennwort;
                    }

                    foreach (var student in studentsMitNeuenFotos)
                    {
                        var tempDatei = Path.Combine(Path.GetTempPath(), student.MailSchulisch + ".jpg");

                        using (var image = Image.Load(student.ZielFotoPfad))
                        {
                            image.Mutate(x => x.Resize(160, 160));
                            image.Mutate(x => x.Rotate(Convert.ToInt32(configuration["RotateFotos"])));
                            image.Save(tempDatei);
                        }

                        byte[] buffer = new byte[4096];
                        string dateiName = Path.GetFileName(tempDatei);

                        ZipEntry entry = new ZipEntry(dateiName)
                        {
                            DateTime = DateTime.Now,
                            CompressionMethod = CompressionMethod.Stored
                        };

                        zip.PutNextEntry(entry);

                        using (FileStream dateiStream = File.OpenRead(tempDatei))
                        {
                            int bytesRead;
                            while ((bytesRead = dateiStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                zip.Write(buffer, 0, bytesRead);
                            }
                        }
                        
                        File.Delete(tempDatei);
                    }

                    try
                    {
                        zip.CloseEntry();                        
                    }
                    catch
                    {
                        throw new Exception("Fehler beim Erstellen des Zip-Archivs (Geevoo). Möglicherweise wurden keine Fotos gefunden.");
                    }
                    
                    zip.IsStreamOwner = true;
                }
            });
            Global.ZeileSchreiben("Fotos gezippt (Geevoo-Mail)", geevooZipPfad, ConsoleColor.Green, ConsoleColor.White);
            // Zweite Zip-Datei: Netman-Kurzname
            var netmanZipPfad = Path.Combine(Path.GetDirectoryName(AbsoluterPfad),
                Path.GetFileNameWithoutExtension(AbsoluterPfad) + "_Netman-Kurzname.zip");

            /*AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Fotos zippen (Netman-Kurzname) ...", ctx =>
            {
                using (FileStream zipStream = File.Create(geevooZipPfad))
                using (ZipOutputStream zip = new ZipOutputStream(zipStream))
                {
                    zip.SetLevel(kompressionsLevel);

                    if (!string.IsNullOrEmpty(kennwort) && kennwort != " ")
                    {
                        zip.Password = kennwort;
                    }

                    foreach (var student in studentsMitNeuenFotos)
                    {
                        var tempDatei = Path.Combine(Path.GetTempPath(), student.MailSchulisch + ".jpg");

                        using (var image = Image.Load(student.ZielFotoPfad))
                        {
                            image.Mutate(x => x.Resize(160, 160));
                            image.Mutate(x => x.Rotate(Convert.ToInt32(configuration["RotateFotos"])));
                            image.Save(tempDatei);
                        }

                        byte[] buffer = new byte[4096];
                        string dateiName = Path.GetFileName(tempDatei);

                        ZipEntry entry = new ZipEntry(dateiName)
                        {
                            DateTime = DateTime.Now,
                            CompressionMethod = CompressionMethod.Stored
                        };

                        zip.PutNextEntry(entry);

                        using (FileStream dateiStream = File.OpenRead(tempDatei))
                        {
                            int bytesRead;
                            while ((bytesRead = dateiStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                zip.Write(buffer, 0, bytesRead);
                            }
                        }

                        File.Delete(tempDatei);
                    }

                    try
                    {
                        zip.CloseEntry();
                    }
                    catch
                    {
                        throw new Exception("Fehler beim Erstellen des Zip-Archivs (Netman). Möglicherweise wurden keine Fotos gefunden.");
                    }

                    zip.IsStreamOwner = true;
                }
            });
            Global.ZeileSchreiben("Fotos gezippt (Geevoo-Mail)", geevooZipPfad, ConsoleColor.Green, ConsoleColor.White);*/
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public void ZippenMitKennwort(IConfiguration configuration)
    {
        ZipPfad = Path.Combine(AbsoluterPfad + ".zip");

        try
        {
            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Zippen ...", ctx =>
            {
                using (FileStream zipStream = File.Create(ZipPfad))
                using (ZipOutputStream zip = new ZipOutputStream(zipStream))
                {
                    zip.SetLevel(0); // Kompressionslevel (0-9, 9 = beste Kompression)

                    if (!string.IsNullOrEmpty(configuration["ZipKennwort"]) && configuration["ZipKennwort"] != " ")
                    {
                        zip.Password = configuration["ZipKennwort"]; // Passwort setzen
                    }

                    byte[] buffer = new byte[4096];
                    string dateiName = Path.GetFileName(AbsoluterPfad);

                    // Zippe ohne Komprimierung
                    ZipEntry entry = new ZipEntry(dateiName)
                    {
                        DateTime = DateTime.Now,
                        CompressionMethod = CompressionMethod.Stored // Keine Komprimierung
                    };

                    zip.PutNextEntry(entry);

                    using (FileStream dateiStream = File.OpenRead(AbsoluterPfad))
                    {
                        int bytesRead;
                        while ((bytesRead = dateiStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            zip.Write(buffer, 0, bytesRead);
                        }
                    }

                    // Durchlaufe die Liste der absoluten Pfade und füge sie dem Zip-Archiv hinzu

                    zip.CloseEntry();
                    zip.IsStreamOwner = true;
                }
            });
            Global.ZeileSchreiben("Datei gezippt", AbsoluterPfad, ConsoleColor.Green, ConsoleColor.White);
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

 internal void Mailen(IConfiguration configuration, string subject, string body, List<string> to, List<string> cc, List<string> bcc, List<string> attachments)
 {
  if (this.Count == 0)
   return;
  
  configuration = Global.Konfig("SmtpUser", Global.Modus.Update, configuration);
  configuration = Global.Konfig("SmtpKennwort", Global.Modus.Update, configuration);
  configuration = Global.Konfig("SmtpPort", Global.Modus.Update, configuration);
  configuration = Global.Konfig("SmtpServer", Global.Modus.Update, configuration);
  var mail = new Mail();
  mail.Senden(configuration, subject, body, to, cc, bcc, attachments);
 }

    

    internal List<dynamic> FilterOpenPeriod()
    {
        return this;
    }

    internal void Verschieben(string zielVerzeichnis)
    {
        if (string.IsNullOrEmpty(AbsoluterPfad) || !File.Exists(AbsoluterPfad))
        {
            Console.WriteLine("Die Datei existiert nicht oder der Pfad ist ungültig.");
            return;
        }

        try
        {                 
            var zielPfad = Path.Combine(zielVerzeichnis, Path.GetFileName(AbsoluterPfad));      
            
            // Lösche die Zieldatei, falls vorhanden
            if (File.Exists(zielPfad))
            {
                File.Delete(zielPfad);
            }

            // Verschiebe die Datei
            File.Move(AbsoluterPfad, zielPfad);                        
            Global.ZeileSchreiben(zielPfad, "", ConsoleColor.Green, ConsoleColor.White);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Verschieben der Datei: {ex.Message}");
        }

    }

    internal void OeffneWebseite(string urlBeginn, string urlMitte = "", string urlEnde = "")
    {
        var mitgliederMail = "";

        if (!string.IsNullOrEmpty(urlMitte))
        {
            mitgliederMail = this
            .Where(rec =>
            {
                if (rec == null) return false;
                var dict = (IDictionary<string, object>)rec;
                return dict != null && dict["MitgliederMail"] != null && !string.IsNullOrWhiteSpace(dict["MitgliederMail"].ToString());
            })
            .Select(rec => ((IDictionary<string, object>)rec)["MitgliederMail"].ToString())
            .LastOrDefault();

            // Wenn MitgliederMail vorhanden ist, verwende es in der URL
            if (!string.IsNullOrEmpty(mitgliederMail))
            {
                urlMitte = mitgliederMail;
            }
        }
        

        // Wenn der URL insgesamt länger als 300 Zeichen ist, wird der urlMitte solange gekürzt, bis die URL passt.
        // Das Kürzen geschieht immer an den Kommas. Diejenigen E-Mail-Adressen, die am Ende übrig bleiben, werden in einem Panel angezeigt.
        var hinweise = new List<string>();
        while ((urlBeginn + urlMitte + urlEnde).Length > 1000)
        {
            var letzteKommaPosition = urlMitte.LastIndexOf(',');
            if (letzteKommaPosition > 0)
            {
                var uebrigeAdressen = urlMitte.Substring(letzteKommaPosition + 1).Trim();
                hinweise.Add(uebrigeAdressen);
                urlMitte = urlMitte.Substring(0, letzteKommaPosition).Trim();
            }
            else
            {
                // Kein Komma mehr gefunden, Abbruch der Schleife
                break;
            }
        }

        if (hinweise.Count > 0)
        {
            var panel = new Panel($"[bold {Global.GetColor(Global.ColorHinweise)}]Die URL ist zu lang. Es konnten nicht alle E-Mail-Adressen berücksichtigt werden.[/]\n[gray]{string.Join("\n", hinweise)}[/]")
                .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Global.ColorHinweise);

            AnsiConsole.Write(panel);
        }

        try
        {
            if (OperatingSystem.IsWindows())
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = urlBeginn + urlMitte + urlEnde,
                    UseShellExecute = true
                });
            }
            else if (OperatingSystem.IsLinux())
            {
                Process.Start("xdg-open", urlBeginn + urlMitte + urlEnde);
            }
            else if (OperatingSystem.IsMacOS())
            {
                Process.Start("open", urlBeginn + urlMitte + urlEnde);
            }

            AnsiConsole.MarkupLine($"[green]Webseite geöffnet: {urlBeginn + urlMitte + urlEnde}[/]");
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Fehler beim Öffnen der Webseite: {ex.Message}[/]");
        }
    }

    internal List<dynamic> FilternGPU020()
    {
        return this;
    }

    internal List<dynamic> FilternLehrkraefteGPU004()
    {
        return this;
    }

    internal bool IstVeraltet(IConfiguration configuration)
    {
        int maxDateiAlter = 6;
        if (int.TryParse(configuration["MaxDateiAlter"], out int parsedMaxDateiAlter))
        {
            maxDateiAlter = parsedMaxDateiAlter;
        }
        if (Erstelldatum.Year > 1 && Erstelldatum.Date.AddDays(maxDateiAlter) < DateTime.Now.Date)
        {
            return true; // Datei ist veraltet
        }
        return false; // Datei ist nicht veraltet
    }

    internal void FehlermeldungRendern(IConfiguration configuration)
    {
        if (!string.IsNullOrEmpty(Fehlermeldung) && !IstOptional)
        {
            if (IstOptional)
            {
                var panel2 = new Panel($"{Fehlermeldung}\n[gray]{string.Join("\n", Hinweise)}[/]")
                .Header("[bold aqua] Optionale Datei [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.DodgerBlue1);

                AnsiConsole.Write(panel2);
            }
            else
            {
                throw new Exception($"[bold {Global.GetColor(Global.ColorHinweise)}]{Fehlermeldung}[/]\n[gray]{string.Join("\n", Hinweise)}[/]");
                var panel2 = new Panel($"[bold {Global.GetColor(Global.ColorHinweise)}]{Fehlermeldung}[/]\n[gray]{string.Join("\n", Hinweise)}[/]")
                .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] !? [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Global.ColorHinweise);

                AnsiConsole.Write(panel2);
            }
        }
    }

    internal List<string> GetFotosAusSchildPfade(IConfiguration configuration, Students students, object webuntis)
    {
        throw new NotImplementedException();
    }

    internal Datei Verarbeiten(Dateien quelldateien, Global.Modus modus)
    {
        var modusString = "Vergleichen";

        if (modus == Global.Modus.Filtern)
            modusString = "Filtern";        

        var neueDatei = new Datei(AbsoluterPfad);
        bool skipProcessing = false;

        var vorhandeneDatei = GetVorhandeneDatei(quelldateien);

        // Wenn keine vorhandene Datei mit demselben Namen existiert, wird die Datei unverändert zurückgegeben.
        if (string.IsNullOrEmpty(vorhandeneDatei))
        {
            var panel = new Panel($"Die Datei [{Global.GetColor(Global.ColorPfadInDateien)}]{Path.GetFileName(AbsoluterPfad)}[/] wurde nicht gefunden. Ein Vergleich von neu & alt ist nicht möglich.")
                .Header("[bold aqua] Neue Datei [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Aqua);
            AnsiConsole.Write(panel);
            return this;
        }

        var table = TableErstellen(
                $"Veränderungen von alt ([{Global.GetColor(Global.ColorPfadInDateien)}]{vorhandeneDatei}[/]) nach neu ([{Global.GetColor(Global.ColorPfadInDateien)}]{AbsoluterPfad}[/]):",
                AnhandDieserSchlüsselAttributeWirdVerglichen);

        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($" {modusString} ...", ctx =>
        {
            var dateiendung = Path.GetExtension(AbsoluterPfad);
            var tableRows = new List<Text>();
            
            var vorhandeneRec = GetVorhandeneRec(quelldateien);
            if (vorhandeneRec == null || vorhandeneRec.Count == 0)
            {
                skipProcessing = true;
                return;
            }

            // Wenn in der vorhandenen eine Spalte namens Nachname existiert und die Dateiendung .dat ist,
            // dann muss in jeder Zeile der Nachname um #Klasse ergänzt werden.
            if (AnhandDieserSchlüsselAttributeWirdVerglichen.Contains("Nachname") && dateiendung == ".dat")
            {
                // In jeder Zeile den Nachnamen um #Klasse ergänzen
                foreach (var vRec in vorhandeneRec)
                {
                    var vorhandeneDict = (IDictionary<string, object>)vRec;
                    if (vorhandeneDict.ContainsKey("Nachname"))
                    {
                        var nachname = vorhandeneDict["Nachname"].ToString();
                        if (vorhandeneDict.ContainsKey("Klasse"))
                        {
                            nachname += "#" + vorhandeneDict["Klasse"];
                        }
                        vorhandeneDict["Nachname"] = nachname;
                    }
                }
            }

            // Die Tabelle soll begrenzt werden. Falls mehr als maxRows Zeilen gefunden werden, wird eine weitere Zeile mit "..." eingefügt.

            var maxRows = 300;
            var rows = 0;

            int ss = 0;

            // Für jede neue Zeile ...
            foreach (var neueRec in this)
            {
                ss++;

                if(ss == 194)
                {
                    var x = 1;
                }
                var neueDict = (IDictionary<string, object>)neueRec;

                if (neueDict.ContainsKey("Fach") && neueDict["Fach"].ToString().StartsWith("DPF"))
                {
                    var x = 1;
                }

                var anhandDieserSchlüsselAttributeWirdVerglichenString = "";

                foreach (var key in AnhandDieserSchlüsselAttributeWirdVerglichen)
                {
                    if (neueDict.TryGetValue(key, out var value) && value != null && !string.IsNullOrEmpty(value.ToString()))
                    {
                        if (anhandDieserSchlüsselAttributeWirdVerglichenString.Length > 0)
                            anhandDieserSchlüsselAttributeWirdVerglichenString += ", ";
                        anhandDieserSchlüsselAttributeWirdVerglichenString += value.ToString();
                    }
                }

                if (neueDict.ContainsKey("Fach") && neueDict["Fach"].ToString() == "GG G1")
                {
                    //if (neueDict.ContainsKey("KursBez") && !neueDict["KursBez"].ToString().StartsWith("ES"))
                    {
                        var x = 1;
                    }
                }

                var dieseAttributeWerdenBeimVergleichIgnoriert = DieseAttributeWerdenBeimVergleichIgnoriert;

                // ... wird geprüft, ob es eine vorhandene Zeile gibt, die auf die Schlüsselattribute matcht.
                var zeileMitIdentischenSchlüsselattributen = GetZeileMitIdentischenSchlüsselattributen(vorhandeneRec, neueDict, AnhandDieserSchlüsselAttributeWirdVerglichen);

                // Fall1: Wenn keine Zeile in den Vergleichsattributen auf die vorhandenen matcht, wird die Zeile neu angelegt.
                if (zeileMitIdentischenSchlüsselattributen == null)
                {
                    neueDatei.Add(neueRec);
                    if (rows > maxRows) break; // Wenn die maximale Anzahl an Zeilen erreicht ist, breche die Schleife ab.                                        
                    if (!string.IsNullOrEmpty(anhandDieserSchlüsselAttributeWirdVerglichenString))
                        table.AddRow(new Text(
                            $"{anhandDieserSchlüsselAttributeWirdVerglichenString}"
                            ), new Text($"  neue Zeile  "), new Text(""), new Text("")); rows++;
                    continue;
                } // und die Schleife übersprungen

                // Fall2: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht, werden abweichende Nicht-Schlüssel-Attributwerte gesucht, ...
                var nichtIdentischeSonstigeAttribute = GetNichtIdentischeSonstigeAttribute(zeileMitIdentischenSchlüsselattributen, neueDict);

                // Fall2a: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht abweichen, ...
                if (nichtIdentischeSonstigeAttribute.Count == 0) continue; // ... überspringe den Rest der Schleife

                // Fall2b: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht identisch sind, ...
                if (nichtIdentischeSonstigeAttribute.Count <= 0) continue;

                // Für die ersten 2 abweichenden Attribute dieser Zeile wird eine Zeile in der Tabelle erstellt.
                for (int i = 0; i < nichtIdentischeSonstigeAttribute.Count; i++)
                {
                    if (rows > maxRows)
                    {
                        if(modus == Global.Modus.Vergleichen)
                            break; // Wenn die maximale Anzahl an Zeilen erreicht ist, breche die Schleife ab.                    
                    }    
                    if (rows == maxRows)
                        table.AddRow(new Text("..."), new Text("..."), new Text("..."), new Text("..."));
                    if (rows < maxRows)
                        table.AddRow(
                            RenderZeile(
                                neueDict,
                                vorhandeneRec,
                                nichtIdentischeSonstigeAttribute[i],
                                zeileMitIdentischenSchlüsselattributen));
                    rows++;
                }
                neueDatei.Add(neueRec);
            }

            if (table.Rows.Count == 0)
                table.AddRow(new Text("keine Änderungen, nichts anzuzeigen"), new Text("..."), new Text("..."), new Text("..."));
        });

        // Im Vergleichsmodus wird die originale Datei nicht verändert.
        if (Global.Modus.Vergleichen == modus)
        {
            AnsiConsole.Write(table);
            if (table.Rows.Count > 3)
                Console.WriteLine("Anzahl Zeilen:" + table.Rows.Count);
            return this;
        }            

        // Im Updatemodus wird die originale Datei mit der neuen Datei überschrieben.
        Clear();
        AddRange(neueDatei);
        return this;
    }

    private string GetVorhandeneDatei(Dateien quelldateien)
    {
        var vorhandeneRec = new List<dynamic>();

        // .dat-Quelldateien haben denselben Namen wie die Zieldatei
        foreach (var vorhandeneDatei in quelldateien)
        {
            if (Path.GetFileName(vorhandeneDatei.AbsoluterPfad.ToLower()) == Path.GetFileName(AbsoluterPfad.ToLower()))
            {
                return vorhandeneDatei.AbsoluterPfad;
                break; // Schleife abbrechen, wenn die Datei gefunden wurde
            }
        }
        // Neue .csv-Dateien beginnen mit demselben Namen wie die Zieldatei, bis zum Unterstrich
        foreach (var vorhandeneDatei in quelldateien)
        {
            if (Path.GetFileNameWithoutExtension(vorhandeneDatei.AbsoluterPfad.ToLower()).Split('_')[0] == Path.GetFileNameWithoutExtension(AbsoluterPfad.ToLower()).Split('_')[0])
            {
                return vorhandeneDatei.AbsoluterPfad;
                break; // Schleife abbrechen, wenn die Datei gefunden wurde
            }
        }

        return string.Empty; // Rückgabe der gefundenen Datei oder leere Liste, wenn keine Datei gefunden wurde
    }

    internal Datei? Filtern(Dateien quelldateien)
    {
        var neueDatei = new Datei(AbsoluterPfad);
        bool skipProcessing = false;
        
        AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Filtern ...", ctx =>
        {
            var dateiendung = Path.GetExtension(AbsoluterPfad);
            
            var vorhandeneRec = GetVorhandeneRec(quelldateien);
            if (vorhandeneRec == null || vorhandeneRec.Count == 0)
            {
                skipProcessing = true;
                return;
            }

            // Wenn in der vorhandenen eine Spalte namens Nachname existiert und die Dateiendung .dat ist,
            // dann muss in jeder Zeile der Nachname um #Klasse ergänzt werden.
            if (AnhandDieserSchlüsselAttributeWirdVerglichen.Contains("Nachname") && dateiendung == ".dat")
            {
                // In jeder Zeile den Nachnamen um #Klasse ergänzen
                foreach (var vRec in vorhandeneRec)
                {
                    var vorhandeneDict = (IDictionary<string, object>)vRec;
                    if (vorhandeneDict.ContainsKey("Nachname"))
                    {
                        var nachname = vorhandeneDict["Nachname"].ToString();
                        if (vorhandeneDict.ContainsKey("Klasse"))
                        {
                            nachname += "#" + vorhandeneDict["Klasse"];
                        }
                        vorhandeneDict["Nachname"] = nachname;
                    }
                }
            }

            // Die Tabelle soll begrenzt werden. Falls mehr als maxRows Zeilen gefunden werden, wird eine weitere Zeile mit "..." eingefügt.

            var maxRows = 100;
            var rows = 0;

            // Für jede neue Zeile ...
            foreach (var neueRec in this)
            {
                var neueDict = (IDictionary<string, object>)neueRec;

                // ... wird geprüft, ob es eine vorhandene Zeile gibt, die auf die Schlüsselattribute matcht.
                var zeileMitIdentischenSchlüsselattributen = GetZeileMitIdentischenSchlüsselattributen(vorhandeneRec, neueDict, AnhandDieserSchlüsselAttributeWirdVerglichen);

                // Fall1: Wenn keine Zeile in den Vergleichsattributen auf die vorhandenen matcht, wird die Zeile neu angelegt.
                if (zeileMitIdentischenSchlüsselattributen == null) { neueDatei.Add(neueRec); continue; } // und die Schleife übersprungen

                // Fall2: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht, werden abweichende Nicht-Schlüssel-Attributwerte gesucht, ...
                var nichtIdentischeSonstigeAttribute = GetNichtIdentischeSonstigeAttribute(zeileMitIdentischenSchlüsselattributen, neueDict);

                // Fall2a: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht abweichen, ...
                if (nichtIdentischeSonstigeAttribute.Count == 0) continue; // ... überspringe den Rest der Schleife

                // Fall2b: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht identisch sind, ...
                if (nichtIdentischeSonstigeAttribute.Count <= 0) continue;

                neueDatei.Add(neueRec);
            }
        });

        if (skipProcessing)
            return this;

        // Entferne alle Zeilen aus this und ersetzte durch neueDatei.
        Clear();
        AddRange(neueDatei);
        return this;
    }

    internal List<dynamic> FilternFaecherGPU006()
    {
        return this;
    }

    internal void OrdnerOeffnen()
    {
        if (this.Count == 0)
        {
            return;
        }
    
        var ordnerPfad = Path.GetDirectoryName(AbsoluterPfad);
    
        try
        {
            if (OperatingSystem.IsWindows())
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = ordnerPfad,
                    UseShellExecute = true
                });
            }
            else if (OperatingSystem.IsLinux())
            {
                // Öffnet den Ordner mit dem Standard-Dateimanager
                Process.Start("xdg-open", ordnerPfad);
            }
            else if (OperatingSystem.IsMacOS())
            {
                // Öffnet den Ordner im Finder
                Process.Start("open", ordnerPfad);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Öffnen des Ordners: {ex.Message}");
        }
    }

    internal List<dynamic>? FilternStudentgroupStudents(Students? students, Klassen klassen)
    {
        // Entferne alle Zeilen, die in der Eigenschaft studentId nicht auf einen Schüler in der Liste students matchen.
        var liste = new List<dynamic>();
        foreach (var rec in this)
        {
            var dict = (IDictionary<string, object>)rec;
            if (dict.ContainsKey("studentId") && students != null && students.Any(s => s.Id.ToString() == dict["studentId"].ToString()))
            {
                liste.Add(rec);
            }
        }
        return liste;
    }

    internal void PutPage()
    {
        var configuration = Global.Konfig("WikiKleineAenderung", Global.Modus.Update, Konfiguration);
        DokuwikiZugriff = new DokuwikiZugriff(configuration);

        DokuwikiZugriff.Options = new XmlRpcStruct
        {
            { "sum", "Automatische Aktualisierung" },
            { "minor", Global.WikiSprechtagKleineAenderung } // Kein Minor-Edit
        };
        
        DokuwikiZugriff.Proxy.PutPage(Name, string.Join("\n", this), new XmlRpcStruct());
    }

    internal void ChatErzeugen(Menüeintrag m)
    {
        var table = new Table();
        table.AddColumn("Nr.");
        table.AddColumn("Gruppe");

        var zulässigeAuswahlOptionen = "";

        for (int i = 1; i < Count; i++)
        {
            string page = this[i].Page;
            table.AddRow(i.ToString(), page);
            zulässigeAuswahlOptionen += i + ", ";
        }

        table.AddRow((Count).ToString(), "(Eine einzelne) Klasse(n) wählen");
        zulässigeAuswahlOptionen += Count;

        AnsiConsole.Write(table);

        var configuration = Global.Konfig("TeamsChatAuswahl", Global.Modus.Update, Konfiguration);

        var nummer = int.Parse(configuration["TeamsChatAuswahl"]);

        if (nummer > 0 && nummer < Count)
        {
            var lehrers = new Lehrers();
            lehrers.GetTeamsUrl(this[nummer - 1].MitgliederMail.Split(';'), String.Join(';', m.IKlassen));            
        }

        if (nummer == Count)
        {
            m.FilterInteressierendeStudentsUndKlassen(configuration);
            var l = m.GetLehrerDerKlassen(configuration, Lehrers ?? []);
            var lehrers = new Lehrers();
            lehrers.GetTeamsUrl(this[nummer - 1].MitgliederMail.Split(';'), String.Join(';', m.IKlassen));
        }
    }

    internal void Auswählen(IConfiguration configuration, Menüeintrag m, Lehrers lehrers, Global.Modus modus)
    {
        // Der Modus AlleGruppen erlaubt die Auswahl aller Gruppen ohne Nachfrage.
        // Der Modus NurEineKlasse führt direkt zur Auswahl einer einzelnen Klasse. Der Modus wird verwendet bei der Zeugnisschreibung, um gezielt eine Klase auszuwählen.

        if (modus != Global.Modus.AlleGruppen)
        {
            this.Clear();
            //m.FilterInteressierendeStudentsUndKlassen(configuration);
            this.AddRange(KlassengruppenAuswählen(configuration, m, lehrers, modus, "Hallo LuL " + m.IKlassen.FirstOrDefault()));
            
            var panel = new Panel("")
                .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] Name des Chats [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Global.ColorHinweise);
            
            AnsiConsole.Write(panel);
            AnsiConsole.MarkupLine($"Zeugnisse {string.Join(", ", m.IKlassen)}");

            var panel2 = new Panel("")
                .Header($"[bold {Global.GetColor(Global.ColorHinweise)}] Nachricht [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Global.ColorHinweise);

            AnsiConsole.Write(panel2);
            
            AnsiConsole.MarkupLine($"Es gibt offene Fehlstunden:\nhttps://bk-borken.webuntis.com/open-periods\nOffene Fehlstunden müssen vor der Zeugniserstellung behandelt werden und werden nicht in das Zeugnis übernommen.");
            
            AnsiConsole.MarkupLine($"Die Zeugnisse sind vorbereitet:\nhttps://bkb.wiki/notenlisten:start\nhttps://bkb.wiki/konferenzen:teilkonferenzen:zeugniskonferenzen:start");

            return;            
        }

        var table = new Table();
        table.AddColumn("Nr.");
        table.AddColumn("Gruppe");

        var zulässigeAuswahlOptionen = "";

        for (int i = 1; i < Count; i++)
        {
            string page = this[i].Page;
            table.AddRow(i.ToString(), page);
            zulässigeAuswahlOptionen += i + ", ";
        }

        table.AddRow((Count).ToString(), "(Eine einzelne) Klasse(n) wählen");
        zulässigeAuswahlOptionen += Count;

        AnsiConsole.Write(table);

        configuration = Global.Konfig("TeamsChatAuswahl", Global.Modus.Update, configuration, "", -1, -1, "", "", null, zulässigeAuswahlOptionen);

        var nummer = int.Parse(configuration["TeamsChatAuswahl"]);

        // Entferne alle zeilen aus this, die nicht der ausgewählten Nummer entsprechen.
        // Wenn die Nummer Count ist, wird die Datei nicht verändert.
        if (nummer > 0 && nummer < Count)
        {
            for (int i = Count - 1; i >= 0; i--)
            {
                if (nummer != i)
                {
                    this.RemoveAt(i);
                }
            }

            this.UrlMitte = this[0].MitgliederMail;
            this.UrlRechts = "&message=" + Uri.EscapeDataString("Hallo ") + this[0].Page.Replace("kollegium:", "").Replace(":start", "").Replace(":fachschaften", "Fachschaft");
        }
        else if (nummer == Count)
        {
            this.Clear();
            m.FilterInteressierendeStudentsUndKlassen(configuration);
            this.AddRange(KlassengruppenAuswählen(configuration, m, lehrers, modus, "Hallo LuL " + m.IKlassen.FirstOrDefault()));            
        }
    }

    private Datei KlassengruppenAuswählen(IConfiguration configuration, Menüeintrag m, Lehrers lehrers, Global.Modus modus, string message)
    {        
        var klasse = m.IKlassen.FirstOrDefault();
        var lehrerDerKlasse = m.GetLehrerDerKlassen(configuration, lehrers ?? []);

        var gpu002 = m.Quelldateien.GetMatchingList(configuration, "gpu002", IStudents, m.Klassen);
        if (gpu002 == null || gpu002.Count == 0) return [];

        var verschiedeneLulKuerzel = gpu002
            .Where(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                var klassenString = dict["Field5"].ToString();
                var klassenListe = klassenString.Split('~'); // Zerlegt den String in eine Liste
                return m.IKlassen.Any(klasse => klassenListe.Contains(klasse)) &&
                    !string.IsNullOrEmpty(dict["Field6"].ToString());
            }).Select(rec =>
            {
                var dict = (IDictionary<string, object>)rec;
                return dict["Field6"].ToString();
            }).Distinct().ToList().OrderBy(x => x).ToList();

        var mitgliederMail = "";
        var mitgliederKuerzel = "";
        var mitglieder = "";
        foreach (var lehrer in lehrers.Where(l => verschiedeneLulKuerzel.Contains(l.Kürzel)))
        {
            if (!string.IsNullOrEmpty(mitgliederMail))
                mitgliederMail += ",";
            mitgliederMail += lehrer.Mail;
            if (!string.IsNullOrEmpty(mitgliederKuerzel))
                mitgliederKuerzel += ",";
            mitgliederKuerzel += lehrer.Kürzel;
            if (!string.IsNullOrEmpty(mitglieder))
                mitglieder += ", ";
            mitglieder += (lehrer.Titel != " " ? lehrer.Titel : "") + lehrer.Vorname + " " + lehrer.Nachname;
        }

        dynamic record = new ExpandoObject();
        record.Mitglieder = mitglieder;
        record.MitgliederMail = mitgliederMail;
        record.MitgliederKuerzel = mitgliederKuerzel;
        
        this.Add(record);
        this.UrlMitte = mitgliederMail;
        this.UrlRechts = "&message=" + Uri.EscapeDataString(message);            
        return this;
    }
}