using Microsoft.Extensions.Configuration;
//using PdfSharp.Pdf.Security;
using iText.Kernel.Pdf;
using iText.Layout;
//using PdfSharp.Pdf;
using System.Text;
using Spectre.Console; // Add this directive for PdfDocument

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

public class PdfDateien : List<PdfDatei>
{
    public List<string> Muster { get; set; }
    public string? InputFolder { get; private set; }
    public string? OutputFolder { get; private set; }

    public PdfDateien()
    {
    }


    public void KennwortSetzen(IConfiguration configuration)
    {
        try
        {
            var fileGroupPdf =
                (from f in Directory.GetFiles(configuration["PfadDownloads"], "*.pdf")
                    where !f.Contains("-Kennwort")
                    select f).ToList();

            Global.ZeileSchreiben("Dateien bereit für die Verschlüsselung:", fileGroupPdf.Count == 0 ? "keine gefunden" : fileGroupPdf.Count.ToString(), ConsoleColor.White, ConsoleColor.DarkBlue);
            
            foreach (var file in fileGroupPdf)
            {
                //Console.WriteLine(" "  + file);
                var path = new TextPath(file);
                AnsiConsole.Write(path);
                Console.WriteLine();
            }
            if (fileGroupPdf.Count == 0)
            {
                return;
            }

            Console.WriteLine("");
            Console.WriteLine("   1: Schips");
            Console.WriteLine("   2: Notenlisten");
            Console.WriteLine("   3: andere PDFs");

            configuration = Global.Konfig("SchipsOderZeugnisseOderAnderePdfs",Global.Modus.Update, configuration, "Was wollen Sie verschlüsseln?");

            var passwort = "";
            var url = "";
            List<string> regex = new List<string>(); 

            if (configuration["SchipsOderZeugnisseOderAnderePdfs"] == "1")
            {   
                configuration = Global.Konfig("SchipsUrl", Global.Modus.Update, configuration, "Schips-Url angeben");
                configuration = Global.Konfig("SchipsPasswort", Global.Modus.Update, configuration, "Schips-Kennwort festlegen");    
                passwort = configuration["SchipsPasswort"];
                url = Global.SchipsUrl;
                regex.Add("schips");                
            }
            else if (configuration["SchipsOderZeugnisseOderAnderePdfs"] == "2")
            {
                configuration = Global.Konfig("ZeugnisUrl", Global.Modus.Update,configuration, "Zeugnis-Url angeben");
                configuration = Global.Konfig("ZeugnisPasswort", Global.Modus.Update, configuration, "Zeugnis-Kennwort festlegen");    
                passwort = configuration["ZeugnisPasswort"];
                url = configuration["ZeugnisUrl"];
            }
            else if (configuration["SchipsOderZeugnisseOderAnderePdfs"] == "3")
            {
                passwort = configuration["PdfKennwort"];
            }
                
            foreach (string fileName in fileGroupPdf)
            {
                // Nurt wenn der Dateiname den Regex groß oder klein enthält.
                // Falls kein Regex, dann alle Dateien
                if (regex.Count == 0 || regex.Any(r => fileName.IndexOf(r, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    var neueDatei = fileName.Replace(Path.GetFileNameWithoutExtension(fileName),
            Path.GetFileNameWithoutExtension(fileName) + "-Kennwort");

        // Writer mit Verschlüsselung konfigurieren
        WriterProperties writerProperties = new WriterProperties()
            .SetStandardEncryption(
                Encoding.UTF8.GetBytes(passwort),         // Benutzerpasswort
                Encoding.UTF8.GetBytes(passwort),        // Besitzerpasswort
                EncryptionConstants.ALLOW_PRINTING |          // Rechte: Nur Drucken erlaubt
                EncryptionConstants.ALLOW_COPY,               // Kopieren erlaubt
                EncryptionConstants.STANDARD_ENCRYPTION_40);      // Verschlüsselung: AES-256

                // PDF-Dokument schreiben
                PdfWriter writer = new PdfWriter(neueDatei, writerProperties);
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fileName), writer);
                Document document = new Document(pdfDoc);
                document.Close();
                pdfDoc.Close();

                    Global.ZeileSchreiben(neueDatei, "Kopie mit Kennwort erstellt", ConsoleColor.Blue,ConsoleColor.Black);
                }                
            }
            Global.OpenWebseite(url);

        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
    }
}