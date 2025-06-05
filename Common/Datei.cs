using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Xml;

namespace Common;

#pragma warning disable CS8603 // Mögliche Null-Verweis-Rückgabe
#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise null-Objekts.
#pragma warning disable CS8604 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8600 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8618 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8619 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0219 // Möglicher Null-Verweis-Argument

public class Datei : List<dynamic>
{
    public string Endung { get; set; } = null!;
    public string Delimiter { get; set; } = null!;
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

    public Datei(string name, bool vorhanden)
    {
        UnterordnerUndDateiname = name;
        Vorhanden = vorhanden;
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
        if (name.Contains("onder"))
        {
            string a = "";
        }
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
            //if (dict["klassen"].ToString().Split('~').Any(klasse => IKlassen.Contains(klasse)))
            if (true)
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
            if (IStudents.Where(student =>
                    student.Nachname == dict["longName"].ToString() &&
                    student.Vorname == dict["foreName"].ToString() &&
                    student.Geburtsdatum == dict["birthDate"].ToString())
                .Any())
            {
                liste.Add(rec);
            }else{
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

    public void Erstellen(string delimiter, char quote, Encoding encoding, bool shouldAllQuote)
    {
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
                var config = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = delimiter,
                    Quote = quote,
                    ShouldQuote = args => shouldAllQuote
                };

                File.Delete(AbsoluterPfad);

                if (this != null && this.Any())
                {
                    using var writer = new StreamWriter(AbsoluterPfad, false, encoding);
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
                //Global.ZeileSchreiben(AbsoluterPfad, this.Count().ToString(), ConsoleColor.DarkMagenta, ConsoleColor.White);

                //var rechteSeite = "neu erstellt, bereit für den Import";

                Global.ZeileSchreiben(AbsoluterPfad, "", ConsoleColor.White, ConsoleColor.Blue);
            }
        }
    }

    public Datei? VergleichenUndFiltern(Dateien vorhandeneDateien, IConfiguration configuration, string[] anhandDieserAttributeWirdVerglichen, string[] dieseAttributeWerdenBeimVergleichIgnoriert)
    {
        var neueDatei = new Datei(AbsoluterPfad);

        List<dynamic> vorhandene = new List<dynamic>();

        // Die vorhandene hat denselben Namen wie die zieldatei
        // die vorhandene Datei wird im Downloadsordner gesucht
        foreach (var vorhandeneDatei in vorhandeneDateien)
        {
            if (Path.GetFileName(vorhandeneDatei.AbsoluterPfad) == Path.GetFileName(AbsoluterPfad))
            {
                vorhandene = vorhandeneDatei;
                break; // Schleife abbrechen, wenn die Datei gefunden wurde
            }
        }
         
        if (vorhandene == null || vorhandene.Count == 0)
        {
            return this;
        }

        // Für jede neue Zeile ...
        foreach (var neueRec in this)
        {
            var neueDict = (IDictionary<string, object>)neueRec;

            var zeileMitIdentischenVergleichsattributen =
                GetZeileMitIdentischenVergleichsattributen(vorhandene, neueDict, anhandDieserAttributeWirdVerglichen);

            // Fall1: Wenn keine Zeile in den Vergleichsattributen auf die vorhandenen matcht, wird die Zeile neu angelegt.
            if (zeileMitIdentischenVergleichsattributen == null)
            {
                neueDatei.Add(neueRec);
                Global.ZeileSchreiben(neueDict["InternBez"].ToString(),
                    "NEU",
                    ConsoleColor.Green, ConsoleColor.White);
                continue; // und die Schleife übersprungen
            }

            var nichtIdentischeSonstigeAttribute = GetNichtIdentischeSonstigeAttribute(
                zeileMitIdentischenVergleichsattributen, neueDict, dieseAttributeWerdenBeimVergleichIgnoriert);

            // Fall2: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht abweichen, ...
            if (nichtIdentischeSonstigeAttribute.Count == 0)
            {
                continue; // ... überspringe den Rest der Schleife
            }

            // Fall3: Wenn eine vorhandene Zeile auf die Vergleichsattribute matcht und die sonstigen Attribute nicht identisch sind, ...
            if (nichtIdentischeSonstigeAttribute.Count <= 0) continue;
            var vorhandeneZeileMitAbweichendenSonstigenAttr =
                GetVorhandeneZeile(neueDict, vorhandene, anhandDieserAttributeWirdVerglichen);
            var linkeSeite = GetLinkeSeite(neueDict, anhandDieserAttributeWirdVerglichen);

            //Console.ForegroundColor = ConsoleColor.DarkYellow;
            foreach (var nichtIdentischesSonstigesAttribut in nichtIdentischeSonstigeAttribute)
            {
                var alt = GetAlterWert(vorhandeneZeileMitAbweichendenSonstigenAttr,
                    nichtIdentischesSonstigesAttribut);
                var neu = GetNeuerWert(neueDict, nichtIdentischesSonstigesAttribut);
                var leh = Global.PrüfeAufNullOrEmpty(neueDict, "Fachlehrer");

                Global.ZeileSchreiben(
                    linkeSeite.ToString().TrimEnd(',') + leh,
                    nichtIdentischesSonstigesAttribut + ": " + alt + " \u2192 " + neu,
                    ConsoleColor.Blue, ConsoleColor.White);
            }


            neueDatei.Add(neueRec);
        }

        if (!neueDatei.Any())
        {
            Global.ZeileSchreiben(AbsoluterPfad, "identische Dateien", ConsoleColor.DarkYellow, ConsoleColor.White);
        }

        return neueDatei;
    }

    private string GetAlterWert(IDictionary<string, object> vorhandeneZeileMitAbweichendenSonstigenAttr,
        string nichtIdentischesSonstigesAttribut)
    {
        if (vorhandeneZeileMitAbweichendenSonstigenAttr == null ||
            string.IsNullOrEmpty(nichtIdentischesSonstigesAttribut))
        {
            return "   "; // Falls Eingaben ungültig sind
        }

        if (vorhandeneZeileMitAbweichendenSonstigenAttr.TryGetValue(nichtIdentischesSonstigesAttribut,
                out object value))
        {
            return value?.ToString(); // Wert als String zurückgeben
        }

        return "   "; // Falls der Key nicht existiert
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
        var k = nichtIdentischesSonstigesAttribut.Replace(".", "PUNKT").Replace(" ", "LEERZEICHEN")
            .Replace("-", "MINUS").Replace("_", "UNTERSTRICH");
        return Global.PrüfeAufNullOrEmpty(neueDict, k);
    }

    private string GetLinkeSeite(IDictionary<string, object> neueDict,
        string[] anhandDieserAttributeWirdVerglichen)
    {
        var linkeSeite = "";

        for (int i = 0; i < anhandDieserAttributeWirdVerglichen.Length; i++)
        {
            if (Global.PrüfeAufNullOrEmpty(neueDict, anhandDieserAttributeWirdVerglichen[i]).Length > 0)
            {
                linkeSeite += Global.PrüfeAufNullOrEmpty(neueDict, anhandDieserAttributeWirdVerglichen[i]) + ", ";
            }
        }

        return linkeSeite.TrimEnd(',').TrimEnd(' ').TrimEnd(',').TrimEnd(' ');
    }

    private List<string> GetNichtIdentischeSonstigeAttribute(IDictionary<string, object> vorhDict,
        IDictionary<string, object> neueDict, string[] dieseAttributeWerdenBeimVergleichIgnoriert)
    {
        List<string> nichtIdentischeSonstige = new List<string>();
        foreach (var key in vorhDict.Keys)
        {
            var k = key.Replace(".", "PUNKT").Replace(" ", "LEERZEICHEN").Replace("-", "MINUS")
                .Replace("_", "UNTERSTRICH").Replace("/", "SCHRÄGSTRICH");
            if (dieseAttributeWerdenBeimVergleichIgnoriert.Contains(k)) continue;
            if (!neueDict.TryGetValue(k, out var value)) continue;
            if (vorhDict[key].Equals(value)) continue;
            // Z.B. bei Fehlstunden bleibt die neue Zelle leer. In der alten steht 0
            if (vorhDict[key].ToString() == "0" && neueDict[k].ToString() == "") continue;
            nichtIdentischeSonstige.Add(key);
        }

        return nichtIdentischeSonstige;
    }

    /// <summary>
    /// Prüfe, ob die Vergleichsattribute identisch sind.
    /// </summary>
    /// <param name="vorhandene"></param>
    /// <param name="neueDict"></param>
    /// <param name="anhandDieserAttributeWirdVerglichen"></param>
    /// <returns></returns>
    public IDictionary<string, object> GetZeileMitIdentischenVergleichsattributen(List<dynamic> vorhandene,
        IDictionary<string, object> neueDict, string[] anhandDieserAttributeWirdVerglichen)
    {
        foreach (var vorhDict in vorhandene.Select(vorhRec => (IDictionary<string, object>)vorhRec))
        {
            var match = anhandDieserAttributeWirdVerglichen.All(key =>
                neueDict.ContainsKey(key) &&
                vorhDict.ContainsKey(key) &&
                neueDict[key].Equals(vorhDict[key])
            );
            if (match)
            {
                return vorhDict;
            }
        }
        return null;
    }

    internal void ZippeBilder(string pfadZuAtlantisFotos)
    {
        
/*        throw new NotImplementedException();
        using (var zipFile = new ZipFile())
        {
            Console.Write("Vorhandene Bilder ".PadRight(75, '.'));

            UTF8Encoding utf8NoBom = new UTF8Encoding(false);
            var filePath = "import.csv";

            File.WriteAllText(filePath, "\"id\",\"custom_id\",\"email\",\"path\"" + Environment.NewLine, utf8NoBom);

        List<string> bereitsVerarbeiteteBilder = new List<string>();
        bereitsVerarbeiteteBilder = (Properties.Settings.Default.bereitsVerarbeiteteBilder).Split(',').ToList();

        int bilderHinzugefügt = 0;
        int bilderBereitsVorhandenUndNichtErneutHinzugefügt = 0;

        foreach (var schueler in this)
        {
            var allFiles = Directory.GetFiles(pfadZuAtlantisFotos, "*" + schueler.Id + ".jpg", SearchOption.AllDirectories);

            foreach (var item in allFiles)
            {
                if (schueler.Id.ToString().Length >= 6)
                {
                    try
                    {
                        if ((from a in bereitsVerarbeiteteBilder where a == schueler.Id.ToString() select a).Count() == 0)
                        {
                            zip.AddFile(item).FileName = schueler.Id + ".jpg";
                            bilderHinzugefügt++;
                            bereitsVerarbeiteteBilder.Add(schueler.Id.ToString());
                            //Console.WriteLine("Bild hinzugefügt für " + schueler.Kurzname);

                            if (schueler.Mail != null && schueler.Mail != "")
                            {
                                File.AppendAllText(filePath, "\"\",\"\",\"" + schueler.Mail + "\",\"" + schueler.Id + ".jpg" + "\"" + Environment.NewLine, utf8NoBom);
                            }
                        }
                        else
                        {
                            bilderBereitsVorhandenUndNichtErneutHinzugefügt++;
                            // Console.WriteLine("bereits vorhanden");
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }
        string f = "";

        foreach (var item in bereitsVerarbeiteteBilder)
        {
            f = f + item + ",";
        }

        Properties.Settings.Default.bereitsVerarbeiteteBilder = f;
        Properties.Settings.Default.Save();

        Console.WriteLine((" " + bilderBereitsVorhandenUndNichtErneutHinzugefügt).PadLeft(30, '.'));

        Console.Write("Neu eingelesene und gezippte Bilder: ".PadRight(75, '.'));

        zip.Save("SchülerBilder-" + DateTime.Now.ToFileTime() + ".zip");

        zip.AddFile(filePath);

        string datei = "Geevoo-Import-" + DateTime.Now.ToFileTime() + ".zip";

        zip.Save(datei);

        Console.WriteLine((" " + bilderHinzugefügt).PadLeft(30, '.'));*/
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

    public string? GetAbsoluterPfad()
    {
        return AbsoluterPfad;
    }

    public void Zippen(string? absoluterPfad, IConfiguration configuration)
    {
        if (this.Count == 0)
        {
            return;
        }
        
        var zipPfad = absoluterPfad.Replace(".csv", ".zip");
        
        try
        {
            using (FileStream zipStream = File.Create(zipPfad))
            using (ZipOutputStream zip = new ZipOutputStream(zipStream))
            {
                zip.SetLevel(9); // Kompressionslevel (0-9, 9 = beste Kompression)

                if (!string.IsNullOrEmpty(Global.ZipKennwort))
                {
                    zip.Password = Global.ZipKennwort; // Passwort setzen
                }

                byte[] buffer = new byte[4096];
                string dateiName = Path.GetFileName(absoluterPfad);

                ZipEntry entry = new ZipEntry(dateiName)
                {
                    DateTime = DateTime.Now,
                    CompressionMethod = CompressionMethod.Deflated
                };

                zip.PutNextEntry(entry);

                using (FileStream dateiStream = File.OpenRead(absoluterPfad))
                {
                    int bytesRead;
                    while ((bytesRead = dateiStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        zip.Write(buffer, 0, bytesRead);
                    }
                }

                zip.CloseEntry();
                zip.IsStreamOwner = true;
            }

            Global.ZeileSchreiben(zipPfad,"erfolgreich erstellt",ConsoleColor.Green, ConsoleColor.White);            
            ZipPfad = zipPfad;
        }
        catch (Exception ex)
        {
            Console.WriteLine("Fehler beim Zippen: " + ex.Message);
        }
    }

    internal void Mailen(string subject, string absendername, string body, IConfiguration configuration)
    {
        if(this.Count == 0)
        {
            return;
        }

        configuration = Global.Konfig("SmtpUser", Global.Modus.Update, configuration, "Mail-Benutzer angeben");
        configuration = Global.Konfig("SmtpPassword", Global.Modus.Update, configuration, "Mail-Kennwort eingeben");
        configuration = Global.Konfig("SmtpPort", Global.Modus.Update, configuration, "SMTP-Port eingeben");
        configuration = Global.Konfig("SmtpServer", Global.Modus.Update, configuration, "SMTP-Server angeben");
        configuration = Global.Konfig("NetmanMailReceiver", Global.Modus.Update, configuration, "Wem soll die Netman-Mail geschickt werden?");
        configuration = Global.Konfig("NetmanMailBccReceiver", Global.Modus.Update, configuration, "Wem soll die Netman-Mail in BCC geschickt werden?","",Global.Datentyp.Mail,"");
        
        var mail = new Mail();
        mail.Senden(subject,configuration,body,ZipPfad, configuration["NetmanMailReceiver"], "", configuration["NetmanMailBccReceiver"]);
    }

    internal List<dynamic> FilternSchildSchuelerExport()
    {
            return this;
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
            Global.ZeileSchreiben(zielPfad, "verschoben", ConsoleColor.Green, ConsoleColor.White);
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

    internal void FehlermeldungRendern(string ersterSatz, IConfiguration configuration)
    {
        if (!string.IsNullOrEmpty(Fehlermeldung))
        {
            if (ersterSatz == "")
                ersterSatz = ersterSatz.TrimEnd('.') + ". " + "\n";

            if (IstOptional)
            {
                var panel2 = new Panel($"[bold blue]{ersterSatz}{Fehlermeldung}[/]\n[gray]{string.Join("\n", Hinweise)}[/]")
                .Header("[bold blue]  Optionale Datei  [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Blue);

                AnsiConsole.Write(panel2);
            }
            else
            {
                var panel2 = new Panel($"[bold red]{ersterSatz}{Fehlermeldung}[/]\n[gray]{string.Join("\n", Hinweise)}[/]")
                .Header("[bold red]  !?  [/]")
                .HeaderAlignment(Justify.Left)
                .SquareBorder()
                .Expand()
                .BorderColor(Color.Red);

                AnsiConsole.Write(panel2);
            }
        }
    }
}