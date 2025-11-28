using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using PdfSharp.Pdf;

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

public partial class PdfSeite
{
    public Student Student { get; set; }
    public string DateiName { get; set; }
    public string Inhalt { get; set; }
    public string Datum { get; set; }
    public int Seite { get; set; }
    public PdfSharp.Pdf.PdfPage Page { get; set; }
    public Lehrers MailReceiver { get; set; } = new Lehrers();

    public PdfSeite(int seite, string inhalt, string dateiName, PdfSharp.Pdf.PdfPage page)
    {
        Seite = seite;
        Page = page;
        Inhalt = inhalt;
        DateiName = dateiName;
    }

    public PdfSeite(int seite, string inhalt, string dateiName)
    {
        Seite = seite;
        Inhalt = inhalt;
        DateiName = dateiName;
    }

    public PdfSeite()
    {
    }

    public Student SchuelerZuSeiteZuordnen(Students students)
    {
        var studs = new List<Student>();

        foreach (var student in students.Where(student => Inhalt.Contains(student.Nachname, StringComparison.OrdinalIgnoreCase) &&
                                                          Inhalt.Contains(student.Vorname, StringComparison.OrdinalIgnoreCase) &&
                                                          Inhalt.Contains(student.Geburtsdatum, StringComparison.OrdinalIgnoreCase)))
        {
            if (!(studs.Where(st =>
                    st.Nachname == student.Nachname && st.Vorname == student.Vorname &&
                    st.Geburtsdatum == student.Geburtsdatum)).Any())
            {
                studs.Add(student);
            }

            // Suche nach einem weiteren Datum (neben dem Geburtsdatum). Das Datum wird der Eigenschaft Zeugnisdatum zugewiesen
            var dateMatches = Regex.Matches(Inhalt, @"\b\d{2}\.\d{2}\.\d{4}\b");
            foreach (Match dateMatch in dateMatches)
            {
                if (dateMatch.Value.Equals(student.Geburtsdatum)) continue;
                Datum = dateMatch.Value == null ? "" : dateMatch.Value;
                break;
            }

            dateMatches = Regex.Matches(Inhalt, @"Borken,\s(\d{2}\.\d{2}\.\d{4})");
            foreach (Match dateMatch in dateMatches)
            {
                if (dateMatch.Value.Equals(student.Geburtsdatum)) continue;
                try
                {
                    Zeugnisdatum = dateMatch.Groups[1].Value;
                }
                catch
                {
                    // ignored
                }

                break;
            }
        }

        if (studs.Count == 1) return studs[0];
        if (studs.Count <= 1) return null!;
        Console.WriteLine("Mehrere Studs gefunden");
        while (Console.KeyAvailable) Console.ReadKey(true);

        Console.ReadKey();
        return null!;
    }

    public List<string> SuchmusterAnwenden(List<string> muster)
    {
        var art = new List<string>();

        if (muster == null || muster.Count == 0) return art;

        // Sortiere die Muster absteigend nach Länge, damit längere (präzisere) Muster zuerst geprüft werden.
        var geordneteMuster = muster
            .Where(m => !string.IsNullOrEmpty(m))
            .OrderByDescending(m => m.Length)
            .ToList();

        foreach (var m in geordneteMuster)
        {
            var match = Regex.Match(Inhalt ?? string.Empty, m.Trim(), RegexOptions.IgnoreCase);
            if (!match.Success) continue;

            var value = match.Value ?? string.Empty;
            if (string.IsNullOrEmpty(value)) continue;

            // Wenn bereits ein vorhandener Eintrag den neuen Wert enthält (längerer Treffer), dann überspringen.
            if (art.Any(existing => existing.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0))
                continue;

            // Wenn der genaue Wert schon vorhanden ist, überspringen.
            if (art.Any(existing => string.Equals(existing, value, StringComparison.OrdinalIgnoreCase)))
                continue;

            art.Add(value);
        }

        return art;
    }

    public string Zeugnisdatum { get; set; }
    public PdfDocument? PdfDocument { get; set; }

    public Student SeiteZuStudentZuordnen(Students students)
    {
        var studs = new List<Student>();

        foreach (var student in students)
        {
            if (student.Vorname == "Mauritz")
            {
                var aa = "";
            }

            if (!Inhalt.Contains(student.Nachname!, StringComparison.OrdinalIgnoreCase) ||
                !Inhalt.Contains(student.Vorname!, StringComparison.OrdinalIgnoreCase) ||
                !Inhalt.Contains(student.Geburtsdatum!, StringComparison.OrdinalIgnoreCase)) continue;
            if (!(studs.Where(st =>
                    st.Nachname == student.Nachname && st.Vorname == student.Vorname &&
                    st.Geburtsdatum == student.Geburtsdatum)).Any())
            {
                studs.Add(student);
            }
        }

        if (studs.Count == 1) return studs[0];
        if (studs.Count <= 1) return null!;
        Console.WriteLine("Mehrere Studs gefunden");
        while (Console.KeyAvailable) Console.ReadKey(true);

        Console.ReadKey();
        return null!;
    }

    public IEnumerable<string> DatumFinden()
    {
        var datum = new List<string>();

        const string muster = @"\bBorken,?\s(0[1-9]|[12][0-9]|3[01])\.(0[1-9]|1[0-2])\.(\d{4})\b";

        // Wenn Zeugnis als Wort oder Namensbestandteil in der Seite vorkommt, dann gib das Wort zurück.
        var match = Regex.Match(Inhalt, muster, RegexOptions.IgnoreCase);

        if (!match.Success) return datum;
        var a = false ? "" : match.Value;

        if (datum.Contains(a)) return datum;
        const string pattern = @"\b\d{2}\.\d{2}\.\d{4}\b";

        var datumMatch = MyRegex().Match(a);
        datum.Add(datumMatch.Value);

        return datum;
    }

    [GeneratedRegex(@"\b\d{2}\.\d{2}\.\d{4}\b")]
    private static partial Regex MyRegex();

    internal void PdfDocumentEncrypt(string passwort)
    {        
            // Verschlüsselung des PDF-Dokuments
        if (PdfDocument != null)
        {
            var securitySettings = PdfDocument.SecuritySettings;
            securitySettings.UserPassword = passwort;
            securitySettings.OwnerPassword = passwort;
            securitySettings.PermitPrint = false; // Drucken deaktivieren
            securitySettings.PermitModifyDocument = false; // Bearbeiten deaktivieren
            securitySettings.PermitExtractContent = false; // Inhalt kopieren deaktivieren
            securitySettings.PermitAnnotations = false; // Anmerkungen deaktivieren
        }
    }

    internal void GetMailReceiver(Lehrers lehrers)
    {
        foreach (var leh in lehrers)               
        {
            if (!string.IsNullOrEmpty(leh.Mail))
            {
                if (IstValideEmail(leh.Mail))
                {
                    if (Inhalt.ToLower().Contains(leh.Mail.ToLower()))
                    {
                        MailReceiver.Add(leh);
                    }
                }
            }
        }
    }

    private bool IstValideEmail(string email)
    {
        var emailRegex = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
        return Regex.IsMatch(email, emailRegex);
    }

    internal void PdfDocumentCreate(string dateiPfad)
    {
        // Öffnen des ursprünglichen Dokuments im Import-Modus
        using (var originalPdf = PdfSharp.Pdf.IO.PdfReader.Open(dateiPfad, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
        {
            var pdfDocument = new PdfDocument();

            // Importieren der Seite aus dem ursprünglichen Dokument
            var importedPage = originalPdf.Pages[Seite - 1]; // Seite.PageNumber ist 1-basiert
            pdfDocument.AddPage(importedPage);

            // Speichere die extrahierte Seite in einer neuen PDF-Datei
            DateiName = Path.GetFileName(dateiPfad);
            //pdfDocument.Save(seitePdfName);
            PdfDocument = pdfDocument;
        }
    }

    internal void Mailen(IConfiguration configuration)
    {    
        foreach (var lehrer in MailReceiver)
        {        
            var receiverEmail = lehrer.Mail;
            var subject = $"{configuration["Betreff"]}";

            var body = configuration["Body"];
            body = body.Replace("#Lehrer#", $"{lehrer.Titel} {lehrer.Vorname} {lehrer.Nachname} ({lehrer.Kürzel})");
            body = body.Replace("\\n", Environment.NewLine);
            if (PdfDocument != null)
            {
                try
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        try
                        {
                            var passwort = configuration["PdfKennwort"];
                            if (!string.IsNullOrEmpty(passwort))
                            {
                                PdfDocumentEncrypt(passwort);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Fehler bei der Verschlüsselung: " + ex.Message);
                        }
                        try
                        {
                            // Speichern des PDF-Dokuments in den MemoryStream
                            PdfDocument.Save(memoryStream, false);
                            memoryStream.Position = 0; // Zurücksetzen des Streams auf den Anfang    
                        } catch (Exception ex)
                        {
                            Console.WriteLine("Fehler beim Speichern des PDF-Dokuments. Ist das PDF-Dokument noch geöffnet?: " + ex.Message);
                        }                        

                        // Erstellen und Senden der E-Mail
                        var mail = new Mail();
                        mail.Senden(configuration, subject, configuration["SmtpUser"], body, memoryStream, this.DateiName, receiverEmail);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Fehler beim Senden der E-Mail: " + ex.Message);
                }
            }
        }
    }
}