using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Microsoft.Extensions.Configuration;
#pragma warning disable CS8600 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8618 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8625 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0618 // Möglicher Null-Verweis-Argument

public class PdfDatei
{
    private PdfDocument _pdfDocument;

    public PdfPages Pages => _pdfDocument.Pages; // Zugriff auf die Seiten des Dokuments
    public string DateiName { get; set; }
    public Students Students { get; set; } = new Students();
    public PdfSeiten Seiten { get; set; } = new PdfSeiten();
    public string Art { get; set; }
    public string Datum { get; set; } 

    public PdfDatei(string dateiPfad)
    {
        _pdfDocument = PdfSharp.Pdf.IO.PdfReader.Open(dateiPfad, PdfSharp.Pdf.IO.PdfDocumentOpenMode.ReadOnly);
        DateiName = dateiPfad;
        Seiten = new PdfSeiten();
        Students = new Students();
    }

    public Students GetStudentsMitSeiten(Students students)
    {
        Students studentsMitSeiten = new Students();

        foreach (var pdfSeite in Seiten)
        {
            // Der passende Student zu dieser Seite wird ermittelt.
            Student student = pdfSeite.SeiteZuStudentZuordnen(students);

            if (student != null)
            {
                Student st = studentsMitSeiten.Where(s => s.Vorname == student.Vorname && s.Nachname.ToLower() == student.Nachname.ToLower() && s.Geburtsdatum == student.Geburtsdatum).FirstOrDefault();

                // Wenn es den Student noch nicht in der Liste gibt
                if (st == null)
                {
                    student.PdfSeiten = new PdfSeiten();
                    student.PdfSeiten.Add(pdfSeite);
                    studentsMitSeiten.Add(student);
                } 
            }
            // Wenn auf geraden Seiten kein Student gefunden wurde, ...
            if (pdfSeite.Seite % 2 == 0 && student == null)
            {
                // ... aber auf der Seite zuvor
                if (studentsMitSeiten.Last().PdfSeiten.Last().Seite == pdfSeite.Seite - 1)
                {
                    // und die Seite eine Rückseite ist, erkennbar an der Schulnummer
                    if (pdfSeite.Inhalt.Contains("177659"))
                    {
                        studentsMitSeiten.Last().PdfSeiten.Add(pdfSeite);
                    }    
                }
            }
        }
        return studentsMitSeiten;
    }

    public void SeitenAusQuelldateienLöschen()
    {
        List<int> del = new List<int>();
        foreach (var student in this.Students)
        {
            foreach (var pdfSeite in student.PdfSeiten)
            {
                del.Add(pdfSeite.Seite);
            }
        }

        PdfDocument document = PdfReader.Open(DateiName, PdfDocumentOpenMode.Modify);
        
        foreach (var pdfSeite in Seiten.OrderByDescending(x => x.Seite).Where(x => x.Student != null))
        {
            document.Pages.RemoveAt(pdfSeite.Seite - 1);
        }

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
        Global.DisplayHeader(configuration);

        if (dateiPfad != null)
        {
            var pdfDatei = new PdfDatei(dateiPfad);
            
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
    }

    

    private PdfSharp.Pdf.PdfPage ConvertToPdfPage(PdfSeite seite)
    {
        var pdfPage = new PdfSharp.Pdf.PdfPage();
        // Kopieren Sie hier die relevanten Inhalte von `PdfSeite` nach `PdfPage`
        // Beispiel: pdfPage.Contents = seite.Contents;
        return pdfPage;
    }
}