//using System.Text;
using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.Extensions.Configuration;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

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

public class PdfSeiten : List<PdfSeite>
{
    public string QuellDateiName { get; set; }
    
    public List<PdfSeite> Seiten { get; private set; } = new List<PdfSeite>();
    public int AnzahlElementeInDieserDatei { get; private set; }

    public void ZwischenseitenZuordnen()
    {
        if (this.Count(x => x.Student != null) > 0)
        {
            var seitenProSchüler = this.Count / this.Count(x => x.Student != null);

            if (this.Count(x => x.Student != null) > 0)
            {
                for (int z = 1; z < this.Count; z = z + seitenProSchüler)
                {
                    int ersteSeite = z;
                    int letzteSeite = z + seitenProSchüler - 1;
                    var schueler = this.Where(x => x.Seite >= ersteSeite && x.Seite <= letzteSeite)
                        .Select(x => x.Student).FirstOrDefault();

                    if (schueler != null)
                    {
                        for (int ii = z; ii < letzteSeite; ii++)
                        {
                            this[ii].Student = schueler;
                        }
                    }
                }
            }
        }
    }

    public void Einlesen(string dateiName)
    {
        List<string> lehrer = new List<string>();

        QuellDateiName = dateiName;

        using (var pdfDocument = PdfDocument.Open(dateiName))
        {
            int seitenNummer = 1;
            foreach (Page page in pdfDocument.GetPages())
            {
                var pdfSeite = new PdfSeite();

                foreach (var word in page.GetWords())
                {
                    pdfSeite.Inhalt += word + " ";
                }

                pdfSeite.DateiName = dateiName;
                pdfSeite.Seite = seitenNummer;
                seitenNummer++;

                this.Add(pdfSeite);
            }
        }
    }


    public string GetDatum()
    {
        List<string> datum = new List<string>();

        foreach (var pdfSeite in this)
        {
            var aa = pdfSeite.DatumFinden();

            foreach (var a in aa)
            {
                datum.Add(a);
            }
        }

        return datum.GroupBy(s => s) // Gruppiere die Strings
            .OrderByDescending(g => g.Count()) // Sortiere nach Häufigkeit
            .FirstOrDefault()?.Key; // Nimm die häufigste Gruppe und gib den Schlüssel zurück;
    }


    internal void ZähleOffeneKlassenbuchEinträge(object lehrer)
    {
        throw new NotImplementedException();
    }
}