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
    public string Dateiname { get; set; } = null!;
    public string ZipPfad { get; private set; }
    public string Fehlermeldung { get; set; }
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

    public Datei(string name, string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert)
    {
        Name = name;
        Dateiname = Path.GetFileName(name);
        AnhandDieserSchlüsselAttributeWirdVerglichen = anhandDieserAttributeWirdVerglichen;
        DieseAttributeWerdenBeimVergleichIgnoriert = dieseAttributeWerdenBeimVergleichIgnoriert;
        UnterordnerUndDateiname = name;
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

    public Datei(string absoluterPfad, string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise) : this(absoluterPfad)
    {
        AbsoluterPfad = absoluterPfad;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
    }

    public Datei(string absoluterPfad, string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert, string delimiter, char quote, Encoding encoding, bool shouldAllQuote, List<string> importhinweise) : this(absoluterPfad)
    {
        AbsoluterPfad = absoluterPfad;
        AnhandDieserSchlüsselAttributeWirdVerglichen = anhandDieserAttributeWirdVerglichen;
        DieseAttributeWerdenBeimVergleichIgnoriert = dieseAttributeWerdenBeimVergleichIgnoriert;
        Delimiter = delimiter;
        Quote = quote;
        Encoding = encoding;
        ShouldAllQuote = shouldAllQuote;
        Importhinweise = importhinweise;
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

    public List<dynamic> FilternAbsencePerLEssons()
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
            new Text(nichtIdentischesSonstigesAttribut.Replace("PUNKT", ".").Replace("LEERZEICHEN", " ").Replace("MINUS", "-").Replace("UNTERSTRICH", "_").Replace("SLASH", "/")),
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
        table.AddColumn(string.Join(", ", anhandDieserSchlüsselAttributeWirdVerglichen));
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
            if (Path.GetFileName(vorhandeneDatei.AbsoluterPfad) == Path.GetFileName(AbsoluterPfad))
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
        nichtIdentischesSonstigesAttribut = nichtIdentischesSonstigesAttribut.Replace(".","PUNKT").Replace(" ","LEERZEICHEN").Replace("-","MINUS").Replace("_","UNTERSTRICH").Replace("/","SLASH");

        return neueDict.TryGetValue(nichtIdentischesSonstigesAttribut, out var neuerWert)
            ? (neuerWert?.ToString()?.Length > 20
                ? neuerWert.ToString().Substring(0, 17) + "..."
                : neuerWert?.ToString() ?? string.Empty)
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
                if (linkeSeite.Length > 30)
                {
                    linkeSeite = linkeSeite.Substring(0, 27) + "...";
                    break;
                }
            }
        }

        return linkeSeite.TrimEnd(',').TrimEnd(' ').TrimEnd(',').TrimEnd(' ');
    }

    private List<string> GetNichtIdentischeSonstigeAttribute(IDictionary<string, object> vorhDict,
        IDictionary<string, object> neueDict)
    {
        List<string> nichtIdentischeSonstige = new List<string>();
        foreach (var key in vorhDict.Keys)
        {
            var k = key.Replace(".", "PUNKT").Replace(" ", "LEERZEICHEN").Replace("-", "MINUS").Replace("_", "UNTERSTRICH").Replace("/", "SCHRÄGSTRICH");
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
            AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start($"Fotos zippen ...", ctx =>
            {
                using (FileStream zipStream = File.Create(AbsoluterPfad))
                using (ZipOutputStream zip = new ZipOutputStream(zipStream))
                {
                    zip.SetLevel(kompressionsLevel); // Kompressionslevel (0-9, 9 = beste Kompression)

                    if (!string.IsNullOrEmpty(kennwort) && kennwort != " ")
                    {
                        zip.Password = kennwort; // Passwort setzen
                    }

                    var studentsMitNeuenFotos = students.Where(s => !string.IsNullOrEmpty(s.ZielFotoPfad) && File.Exists(s.ZielFotoPfad)).ToList();
                    if(studentsMitNeuenFotos.Count == 0)
                    {
                        throw new Exception("Es wurden keine Fotos gefunden, die gezippt werden könnten.");
                    }

                    foreach (var student in studentsMitNeuenFotos)
                    {
                        var tempDatei = Path.Combine(Path.GetTempPath(), student.MailSchulisch.Split('@')[0] + ".jpg");

                        // Verkleinere jedes zu zippende Bild auf 160x160px. Das Bild im Original soll aber unverändert bleiben.

                        using (var image = Image.Load(student.ZielFotoPfad))
                        {
                            image.Mutate(x => x.Resize(160, 160));
                            image.Mutate(x => x.Rotate(Convert.ToInt32(configuration["RotateFotos"])));
                            // Speichere die Datei im temporären Verzeichnis
                            image.Save(tempDatei);
                        }

                        byte[] buffer = new byte[4096];
                        string dateiName = Path.GetFileName(tempDatei);

                        // Füge die thumb-datei hinzu, wobei der Name ohne _thumb im Zipfile liegen soll.
                        // Zippe ohne Komprimierung
                        ZipEntry entry = new ZipEntry(dateiName)
                        {
                            DateTime = DateTime.Now,
                            CompressionMethod = CompressionMethod.Stored // Keine Komprimierung
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
                        // Lösche die temporäre Datei
                        File.Delete(tempDatei);
                    }

                    // Durchlaufe die Liste der absoluten Pfade und füge sie dem Zip-Archiv hinzu
                    try
                    {
                        zip.CloseEntry();                        
                    }
                    catch
                    {
                        throw new Exception("Fehler beim Erstellen des Zip-Archivs. Möglicherweise wurden keine Fotos gefunden.");
                    }
                    
                    zip.IsStreamOwner = true;
                }
            });
            Global.ZeileSchreiben("Fotos gezippt", AbsoluterPfad, ConsoleColor.Green, ConsoleColor.White);
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

    internal void Mailen(string subject, string body, IConfiguration configuration, string attachment = "")
    {
        if (this.Count == 0)
        {
            return;
        }

        configuration = Global.Konfig("SmtpUser", Global.Modus.Update, configuration);
        configuration = Global.Konfig("SmtpKennwort", Global.Modus.Update, configuration);
        configuration = Global.Konfig("SmtpPort", Global.Modus.Update, configuration);
        configuration = Global.Konfig("SmtpServer", Global.Modus.Update, configuration);
        //configuration = Global.Konfig("BccAdresse", Global.Modus.Update, configuration);
        configuration = Global.Konfig("NetmanMailReceiver", Global.Modus.Update, configuration);
        configuration = Global.Konfig("NetmanMailBccReceiver", Global.Modus.Update, configuration);

        var mail = new Mail();
        mail.Senden(subject, configuration, body, attachment, configuration["NetmanMailReceiver"], "", configuration["NetmanMailBccReceiver"]);
    }

    internal List<dynamic> FilterOpenPeriod()
    {
        return this;
    }

    internal void Verschieben(string v)
    {
        if (string.IsNullOrEmpty(AbsoluterPfad) || !File.Exists(AbsoluterPfad))
        {
            Console.WriteLine("Die Datei existiert nicht oder der Pfad ist ungültig.");
            return;
        }

        try
        {                 
            var zielPfad = Path.Combine(v, Path.GetFileName(AbsoluterPfad));      
            
            // Verschiebe die Datei
            File.Move(AbsoluterPfad, zielPfad);                        
            Global.ZeileSchreiben(zielPfad, "", ConsoleColor.Green, ConsoleColor.White);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Verschieben der Datei: {ex.Message}");
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
        var table = TableErstellen(
                $"Veränderungen von alt ([{Global.GetColor(Global.ColorPfadInDateien)}]{vorhandeneDatei}[/]) nach neu:",
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

            var maxRows = 100;
            var rows = 0;

            // Für jede neue Zeile ...
            foreach (var neueRec in this)
            {
                var neueDict = (IDictionary<string, object>)neueRec;
                var ersteSpalte = neueDict[AnhandDieserSchlüsselAttributeWirdVerglichen[0]].ToString();
                var zweiteSpalte = AnhandDieserSchlüsselAttributeWirdVerglichen.Length > 1 ? neueDict[AnhandDieserSchlüsselAttributeWirdVerglichen[1]].ToString() : "";
                var dritteSpalte = AnhandDieserSchlüsselAttributeWirdVerglichen.Length > 2 ? neueDict[AnhandDieserSchlüsselAttributeWirdVerglichen[2]].ToString() : "";
                var vierteSpalte = AnhandDieserSchlüsselAttributeWirdVerglichen.Length > 3 ? neueDict[AnhandDieserSchlüsselAttributeWirdVerglichen[3]].ToString() : "";
                var fünfteSpalte = AnhandDieserSchlüsselAttributeWirdVerglichen.Length > 4 ? neueDict[AnhandDieserSchlüsselAttributeWirdVerglichen[4]].ToString() : "";
                var anhandDieserSchlüsselAttributeWirdVerglichenString = ersteSpalte + (zweiteSpalte.Length > 0 ? ", " + zweiteSpalte : "") + (dritteSpalte.Length > 0 ? ", " + dritteSpalte : "") + (vierteSpalte.Length > 0 ? ", " + vierteSpalte : "") + (fünfteSpalte.Length > 0 ? ", " + fünfteSpalte : "");
                //anhandDieserSchlüsselAttributeWirdVerglichenString = anhandDieserSchlüsselAttributeWirdVerglichenString.Length > 60 ? anhandDieserSchlüsselAttributeWirdVerglichenString.Substring(0, 57) + "..." : anhandDieserSchlüsselAttributeWirdVerglichenString;

                if(neueDict.ContainsKey("Fach") && neueDict["Fach"].ToString() == "BWR")
                {
                    if (neueDict.ContainsKey("KursBez") && !neueDict["KursBez"].ToString().StartsWith("ES"))
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
                            ), new Text($"  neue Zeile  "), new Text(""), new Text(""));
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
                    if (rows > maxRows) break; // Wenn die maximale Anzahl an Zeilen erreicht ist, breche die Schleife ab.                    
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
                    neueDatei.Add(neueRec);
                }
            }

            if (table.Rows.Count == 0)
                table.AddRow(new Text("keine Änderungen, nichts anzuzeigen"), new Text("..."), new Text("..."), new Text("..."));
        });

        // Im Vergleichsmodus wird die originale Datei nicht verändert.
        if (Global.Modus.Vergleichen == modus)
        {
            AnsiConsole.Write(table);
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

        foreach (var vorhandeneDatei in quelldateien)
        {
            if (Path.GetFileName(vorhandeneDatei.AbsoluterPfad) == Path.GetFileName(AbsoluterPfad))
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
        Process.Start(new ProcessStartInfo
        {
            FileName = Path.GetDirectoryName(AbsoluterPfad),
            UseShellExecute = true
        });
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
}