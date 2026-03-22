//using iText.Kernel.Pdf;
using Microsoft.Extensions.Configuration;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Spectre.Console;
using System.Globalization;
using System.Text.RegularExpressions;
using static Global;
#pragma warning disable CS8600 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8618 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8625 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0618 // Möglicher Null-Verweis-Argument

public class PdfDatei
{
    private PdfDocument _pdfDocument;

    // Wenn der Anwender 'w' wählt, wird die Nachfrage bis zum Neustart der App unterdrückt.
    private static bool _skipSaveConfirmation = false;


    public PdfPages Pages => _pdfDocument.Pages; // Zugriff auf die Seiten des Dokuments
    public string DateiName { get; set; }
    public Students Students { get; set; } = new Students();
    public PdfSeiten Seiten { get; set; } = new PdfSeiten();
    public string Art { get; set; }
    public string Datum { get; set; }
    public int AnzahlElementeInDieserDatei { get; set; }

    public PdfDatei(string dateiPfad)
    {
        _pdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(dateiPfad, PdfSharp.Pdf.IO.PdfDocumentOpenMode.ReadOnly);
        DateiName = dateiPfad;
        Seiten = new PdfSeiten();
        Students = new Students();
    }

    public Students PdfDateiVerarbeiten(Students students, IConfiguration configuration)
    {
        Students studentsMitSeiten = new Students();
        
        var laufendeSeitennummer = 1;

        var table = new Table();
        table.Expand();
        table.Border(TableBorder.Rounded);

        var title = Path.GetFileName("Datei: " + DateiName) + " Art:" + Art + " Seiten:" + Seiten.Count + " Elemente:" + AnzahlElementeInDieserDatei;

        if (!string.IsNullOrEmpty(title))
            table.Title = new TableTitle(title);

        var zuLöschendeSeiten = new List<int>();

        table.Expand();
        table.AddColumn("Nr.");
        table.AddColumn("Nachname");
        table.AddColumn("Vorname");
        table.AddColumn("GebDat");
        table.AddColumn("1.Seite");
        table.AddColumn("letzte");
        table.AddColumn("Zeugnisdat");
        table.AddColumn("Zielordner");
        table.AddColumn("Zieldatei");        
        table.AddColumn("gelöschte Seiten (Index)");


        // Durchlaufe alle Elemente (verschiedene )Schüler) in dieser Datei
        for (var i = 0; i < AnzahlElementeInDieserDatei; i++)
        {
            var pdfSeiten = new PdfSeiten();


            // Die Inhalte aller Seiten eines Elements zusammenfügen
            var inhalt = "";

            // Alle Seiten eines Elements durchlaufen
            var ersteSeiteDesElements = laufendeSeitennummer;
            var letzteSeiteDesElements = ersteSeiteDesElements + (Seiten.Count / AnzahlElementeInDieserDatei) - 1;

            for (laufendeSeitennummer = ersteSeiteDesElements; laufendeSeitennummer <= letzteSeiteDesElements; laufendeSeitennummer++)
            {
                inhalt += Seiten[laufendeSeitennummer - 1].Inhalt;
                pdfSeiten.Add(Seiten[laufendeSeitennummer - 1]);
            }

            // Den Student aus dem Inhalt aller Seiten eines Elements extrahieren
            
            Student student = students.Where(s => inhalt.Contains(s.Vorname + " " + s.Nachname) && inhalt.Contains(s.Geburtsdatum.ToString())).FirstOrDefault();

            if(student == null)
            {
                // Es kann sein, dass das Geburtsdatum d.m.yyyy geschrieben ist
                student = students.FirstOrDefault(s => 
                    inhalt.Contains($"{s.Vorname} {s.Nachname}") && 
                    inhalt.Contains(DateTime.Parse(s.Geburtsdatum).ToString("d.M.yyyy"))
                );
            }
            if(student == null)
            {
                // Es kann sein, dass das Geburtsdatum fehlt
                student = students.FirstOrDefault(s => 
                    inhalt.Contains($"{s.Vorname} {s.Nachname}")
                );
            }
            if(student == null)
            {
                // Es kann sein, dass nur Nachname und Geburtsdatum matchen, wegen z.B. Doppelvornamen oder fehlendem Vornamen
                student = students.FirstOrDefault(s => 
                    inhalt.Contains($"{s.Nachname}")
                    && 
                    inhalt.Contains(DateTime.Parse(s.Geburtsdatum).ToString("dd.MM.yyyy"))
                );
            }
            if(student == null)
            {
                // Es kann sein, dass nur Nachname und Geburtsdatum matchen, wegen z.B. Doppelvornamen oder fehlendem Vornamen
                student = students.FirstOrDefault(s => 
                    inhalt.Contains($"{s.Nachname}")
                    && 
                    inhalt.Contains(DateTime.Parse(s.Geburtsdatum).ToString("d.M.yyyy"))
                );
            }
            if(student == null)
            {
                // Es kann sein, dass nur Vorname und Geburtsdatum matchen, wegen z.B. Doppelvornamen oder fehlendem Vornamen
                student = students.FirstOrDefault(s => 
                    inhalt.Contains($"{s.Vorname}")
                    && 
                    inhalt.Contains(DateTime.Parse(s.Geburtsdatum).ToString("d.M.yyyy"))
                );
            }
            if(student == null)
            {
                // Es kann sein, dass nur Vorname und Geburtsdatum matchen, wegen z.B. Doppelvornamen oder fehlendem Vornamen
                student = students.FirstOrDefault(s => 
                    inhalt.Contains($"{s.Vorname}")
                    && 
                    inhalt.Contains(DateTime.Parse(s.Geburtsdatum).ToString("dd.MM.yyyy"))
                );
            }

            // Suche nach Datumsangaben im Text (dd.MM.yyyy, d.M.yyyy, dd.MM.yy, d.M.yy, mit . / -)
            var dateMatches = Regex.Matches(inhalt, @"\b\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}\b")
                                   .Cast<Match>()
                                   .Select(m => m.Value)
                                   .Distinct()
                                   .ToList();

            var formats = new[] { "dd.MM.yyyy", "d.M.yyyy", "dd.MM.yy", "d.M.yy" };
            var parsedDates = new List<DateTime>();

            foreach (var ds in dateMatches)
            {
                if (DateTime.TryParseExact(ds, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                {
                    parsedDates.Add(dt.Date);
                }
                else if (DateTime.TryParse(ds, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out dt))
                {
                    parsedDates.Add(dt.Date);
                }
            }

            // Geburtstagsdatum des Schülers versuchen zu parsen, damit es ausgeschlossen werden kann
            DateTime? geburtsDatumParsed = null;
            if (student != null && !string.IsNullOrEmpty(student.Geburtsdatum))
            {
                if (DateTime.TryParseExact(student.Geburtsdatum, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var bd))
                    geburtsDatumParsed = bd.Date;
                else if (DateTime.TryParse(student.Geburtsdatum, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out bd))
                    geburtsDatumParsed = bd.Date;
            }

            // Alle gefundenen Datumsangaben filtern: Geburt ausschließen, future-dates ignorieren
            var candidateDates = parsedDates
                .Where(d => geburtsDatumParsed == null || d.Date != geburtsDatumParsed.Value.Date)
                .Where(d => d <= DateTime.Now.AddDays(1)) // vermeide offensichtliche Zukunftsdaten
                .Distinct()
                .OrderByDescending(d => d)
                .ToList();

            // Wähle das plausibelste Datum: das jüngste Datum <= heute (oder null)
            DateTime? zeugnisDatum = candidateDates.FirstOrDefault();

            var gefundenesDatum = zeugnisDatum?.ToString("dd.MM.yyyy") ?? "";

            if (student == null)
            {
                table.AddRow(new Text((i + 1).ToString()), new Text(""), new Text(" "), new Text(" "), new Text($"{ersteSeiteDesElements}"), new Text($"{letzteSeiteDesElements}"), new Text($"{""}"), new Text($"{""}"), new Text($"{""}"), new Text($"{""}"));
            }
            else 
            {
                student.Zielordner = Path.Combine(configuration["PfadDokumentenverwaltung"], student.Nachname.Substring(0, 1), $"{student.Nachname}, {student.Vorname}, {student.Geburtsdatum.ToString()}");
                student.Zielordner = student.Zielordner.Replace(".", "_");
                var ordnerNeu = "vorhanden";

                if (!Directory.Exists(student.Zielordner))
                {
                    Directory.CreateDirectory(student.Zielordner);
                    ordnerNeu = "neu";
                }

                

                PdfDocument quelldatei = PdfReader.Open(DateiName, PdfDocumentOpenMode.Import);
                PdfDocument zieldatei = new PdfDocument();

                student.PdfSeiten = pdfSeiten;

                foreach (var pdfSeite in pdfSeiten)
                {
                    zieldatei.AddPage(quelldatei.Pages[pdfSeite.Seite - 1]);
                }

                CheckObDateiGeschlossen(DateiName);



                zieldatei.Save(student.Zielordner + "/" + $"{student.Nachname}_{student.Vorname}_{student.Geburtsdatum}_{Art}_{gefundenesDatum}.pdf");

                UserPrompts.ConfirmOrThrowSeiten($"Wurde {Path.GetFileName($"{student.Nachname}_{student.Vorname}_{student.Geburtsdatum}_{Art}_{gefundenesDatum}.pdf")} korrekt im Zielordner gespeichert?", student);

                // Löschen
                for (int j = student.PdfSeiten.Count; 0 < j; j--)
                {
                    var seite = student.PdfSeiten[j - 1];
                    zuLöschendeSeiten.Add(seite.Seite - 1);                
                }

                table.AddRow(new Text((i + 1).ToString()), new Text(student.Nachname), new Text(student.Vorname), new Text(student.Geburtsdatum), new Text($"{ersteSeiteDesElements}"), new Text($"{letzteSeiteDesElements}"), new Text(gefundenesDatum), new Text(ordnerNeu), new Text("erstellt"), new Text(string.Join(',', zuLöschendeSeiten.OrderBy(x=>x))));

                studentsMitSeiten.Add(student);
            }
        }
                
        CheckObDateiGeschlossen(DateiName);
        //OpenFolder(Path.GetDirectoryName(DateiName), true);

        using (var document = PdfReader.Open(DateiName, PdfDocumentOpenMode.Modify))
        {
            foreach (var z in zuLöschendeSeiten.OrderByDescending(x=>x))
            {                 
                document.Pages.RemoveAt(z);
            }

            if (document.Pages.Count == 0)
            {
                document.Close(); // optional, wird beim Dispose ausgeführt
                File.Delete(DateiName);
            }
            else
            {
                document.Save(DateiName);
            }
        } // document disposed hier
              
        
        AnsiConsole.Write(table);

        UserPrompts.ConfirmOrThrowDateien($"{Path.GetFileName(DateiName)} korrekt gelöscht bzw. Seiten entfernt?");

        return studentsMitSeiten;
    }

    public void SeitenAusQuelldateienLöschen()
    {
        var document = PdfSharp.Pdf.IO.PdfReader.Open(DateiName, PdfDocumentOpenMode.Modify);
        
        for (int i = Students.Count; 0 < i; i--)
        {
            for (int j = Students[i - 1].PdfSeiten.Count; 0 < j; j--)
            {
                var seite = Students[i - 1].PdfSeiten[j - 1];
                document.Pages.RemoveAt(seite.Seite - 1);
            }
        }

        if (document.Pages.Count == 0)
            File.Delete(DateiName);
        if (document.Pages.Count > 0)
            document.Save(DateiName);
    }

    public void Einlesen(string dateiPfad)
    {
        try
        {
            // Öffne die PDF-Datei
            using (var document = PdfSharp.Pdf.IO.PdfReader.Open(dateiPfad, PdfSharp.Pdf.IO.PdfDocumentOpenMode.ReadOnly))
            {
                // Iteriere durch alle Seiten der PDF-Datei
                for (int i = 0; i < document.PageCount; i++)
                {
                    var page = document.Pages[i];
                    var inhalt = ExtrahiereTextAusSeite(page, dateiPfad); // Implementieren Sie diese Methode, um den Text zu extrahieren
                    Seiten.Add(new PdfSeite(i + 1, inhalt, null, page));
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Einlesen der PDF-Datei: {ex.Message}");
        }
    }

    private string ExtrahiereTextAusSeite(PdfSharp.Pdf.PdfPage page, string dateiPfad)
    {
        // Da PdfSharp keine Textextraktion unterstützt, verwenden wir PdfPig
        using (var pdfDocument = UglyToad.PdfPig.PdfDocument.Open(dateiPfad))
        {
            // Bestimme die Seitennummer durch Iteration
            int pageIndex = 0;
            foreach (var pdfPage in page.Owner.Pages)
            {
                pageIndex++;
                if (pdfPage == page)
                {
                    break;
                }
            }

            // Hole die entsprechende Seite basierend auf der Seitennummer
            var pdfPigPage = pdfDocument.GetPage(pageIndex);

            // Extrahiere den gesamten Text der Seite
            return pdfPigPage.Text;
        }
    }

    public PdfDatei(IConfiguration configuration, string dateiPfad, Lehrers lehrers)
    {
        if (dateiPfad != null)
        {
            // Öffne die PDF-Datei
            using (var document = PdfReader.Open(dateiPfad, PdfDocumentOpenMode.ReadOnly))
            {
                // Iteriere durch alle Seiten der PDF-Datei
                for (int i = 0; i < document.PageCount; i++)
                {
                    var page = document.Pages[i];
                    var inhalt = ExtrahiereTextAusSeite(page, dateiPfad); // Implementieren Sie diese Methode, um den Text zu extrahieren
                    Seiten.Add(new PdfSeite(i + 1, inhalt, null, page));
                }
            }
        }
    }

    

    private PdfSharp.Pdf.PdfPage ConvertToPdfPage(PdfSeite seite)
    {
        var pdfPage = new PdfSharp.Pdf.PdfPage();
        // Kopieren Sie hier die relevanten Inhalte von `PdfSeite` nach `PdfPage`
        // Beispiel: pdfPage.Contents = seite.Contents;
        return pdfPage;
    }

    internal string GetArt(List<string> schlüsselwörter)
    {
        List<string> art = new List<string>();

        foreach (var pdfSeite in this.Seiten)
        {
            var aa = pdfSeite.SuchmusterAnwenden(schlüsselwörter);

            foreach (var a in aa)
            {
                if (!art.Contains(a.Trim()))
                {
                    art.Add(a.Trim());
                }
            }
        }

        if (art.Count == 1)
        {
            return art.OrderByDescending(x=>x).FirstOrDefault();
        }

        if (art.Count == 0)
        {
            Console.WriteLine("Art nicht erkannt: " + this.DateiName);
            //while (Console.KeyAvailable) Console.ReadKey(true);

            //Console.ReadKey();
            return "";
        }
        return art[0];        
    }

    internal int GetAnzahlElementeProDatei(IConfiguration configuration)
    {
        int anzahlElemete = 0;

        foreach (var pdfSeite in this.Seiten)
        {
            // Wenn das Sort Schulnummer und diie Schulnummer selbst gefunden wurde, dann zählen.
            if (pdfSeite.Inhalt.Contains("Schulnummer") && pdfSeite.Inhalt.Contains(configuration["Schulnummer"]))
            {
                anzahlElemete++;
            }
        }

        // Wenn die Schulnummer nicht gefunden wurde, suche auch nochmal nach dem Wort Rechtsmittelbelehrung
        if (anzahlElemete == 0)
        {
            foreach (var pdfSeite in this.Seiten)
            {
                // Wenn das Sort Schulnummer und diie Schulnummer selbst gefunden wurde, dann zählen.
                if (pdfSeite.Inhalt.Contains("Rechtsmittelbelehrung") || pdfSeite.Inhalt.Contains("Rechtsbehelfsbelehrung")|| pdfSeite.Inhalt.Contains("APO-BK"))
                {
                    anzahlElemete++;
                }
            }
        }

        // E muss immer mindestens 1 Element geben
        //return Math.Max(1, anzahlElemete);
        return anzahlElemete;
    }
}