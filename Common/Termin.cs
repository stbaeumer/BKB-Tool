using System.Text.RegularExpressions;

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

public class Termin
{
    public string Betreff { get; set; }
    public List<string> Seite { get; set; }
    public string Hinweise { get; set; }
    public DateTime Datum { get; set; }
    public string KategorienString { get; set; }
    public List<string> Kategorien { get; set; }
    public List<string> Ressourcen { get; set; }
    public string Verantwortlich { get; set; }
    public string Ort { get; set; }

    public string OptionaleTeilnehmer { get; set; }
    public string Nachricht { get; set; }
    public string FälligUm { get; set; }
    public string Beschriftung { get; set; }
    
    public DateTime Beginn { get; set; }
    public DateTime Ende { get; set; }
    public List<string> SJ { get; internal set; }
    public string BeginnString { get; set; }
    public string EndeString { get; set; }
    public string Serienmuster { get; set; }

    internal void ToWikiLink(string nachrichtEnthält)
    {            
        if (!string.IsNullOrEmpty(Nachricht))
        {
            Hinweise = Nachricht.Trim();

            // Regex-Muster für Hyperlinks
            string pattern = @"http[s]?://[^\s]+";

            // Regex-Objekt erstellen
            Regex regex = new Regex(pattern);

            // Hyperlinks im Text finden
            MatchCollection matches = regex.Matches(Nachricht);

            Seite = new List<string>();

            // Gefundene Hyperlinks ausgeben
            foreach (Match match in matches)
            {   
                if (match.Value.Contains(nachrichtEnthält))
                {
                    Seite.Add(match.Value.Replace(nachrichtEnthält, "").TrimEnd('>'));
                    Hinweise = Hinweise.Replace(match.ToString(), "").Trim();
                }
            }
        }
    }
}