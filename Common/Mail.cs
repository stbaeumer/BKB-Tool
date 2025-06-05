using System.Net;
using System.Net.Mail;
using MailKit.Net.Smtp;
using Microsoft.Extensions.Configuration;
using MimeKit;
using MimeKit.Utils;

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


public class Mail
{
    public Mail() {}

    public Mail(string mailsCsv, string campusfestjpg, string betreffMassenmail, IConfiguration configuration, int anzahl)
    {        
        var empfänger = GetNächsteMailadressen(mailsCsv, anzahl);
        empfänger.Add("stefan.baeumer@berufskolleg-borken.de");
        SendenMitEingebettetemBild(betreffMassenmail, campusfestjpg, empfänger, configuration);
        LöscheEmpfängerAusListe(empfänger, mailsCsv);
    }

    private void LöscheEmpfängerAusListe(List<string> empfänger, string mailsCsv)
    {
        try
        {
            var zeilen = File.ReadAllLines(mailsCsv).ToList();

            foreach (var empfängerEmail in empfänger)
            {
                zeilen.RemoveAll(zeile => zeile.Contains(empfängerEmail));
            }

            File.WriteAllLines(mailsCsv, zeilen);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Löschen der Empfänger aus der Liste: {ex.Message}");
        }
    }

    private List<string> GetNächsteMailadressen(string mailsCsv, int anzahl)
    {
        var gültigeMailadressen = new List<string>();

        try
        {
            // Lese alle Zeilen aus der CSV-Datei
            var zeilen = File.ReadAllLines(mailsCsv);

            foreach (var zeile in zeilen)
            {
                // Extrahiere die E-Mail-Adresse (z. B. wenn die CSV mehrere Spalten hat)
                var spalten = zeile.Split(';'); // Annahme: Spalten sind durch Semikolon getrennt
                if (spalten.Length > 0)
                {
                    var email = spalten[0].Trim(); // Erste Spalte enthält die E-Mail-Adresse

                    // Überprüfe, ob die E-Mail-Adresse syntaktisch gültig ist
                    if (IstMailadresseGültig(email))
                    {
                        gültigeMailadressen.Add(email);

                        // Beende die Schleife, wenn anzahl Adressen gesammelt wurden
                        if (gültigeMailadressen.Count == anzahl)
                        {
                            break;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Lesen der Datei {mailsCsv}: {ex.Message}");
        }

    return gültigeMailadressen;
}

private bool IstMailadresseGültig(string email)
{
    try
    {
        var addr = new System.Net.Mail.MailAddress(email);
        return addr.Address == email;
    }
    catch
    {
        return false;
    }
}

    public string BetreffMassenmail { get; }
    public string BodyMassenmail { get; }
    public IConfiguration Configuration { get; }

    public void Senden(string subject, IConfiguration configuration, string body, string attachment, string receiverEmail, string cc = "", string bcc = "") 
    {
        try{
            string smtpServer = configuration["SmtpServer"];
            int smtpPort = Convert.ToInt32(configuration["SmtpPort"]);
            string senderEmail = configuration["SmtpUser"];

            if(configuration["SmtpPassword"] == null || configuration["SmtpPassword"].Length <= 3)
            {
                Console.WriteLine("Bitte geben Sie das Passwort von " + configuration["SmtpUser"] +" für den E-Mail-Versand ein:");
                configuration["SmtpPassword"] = Console.ReadLine();
            }
            
            string senderPassword = configuration["SmtpPassword"];

            var email = new MimeMessage();
            email.From.Add(new MailboxAddress(configuration["SmtpUser"], senderEmail));
            email.To.Add(new MailboxAddress("Empfänger", receiverEmail));
            //email.To.Add(new MailboxAddress("Empfänger", "stefan.baeumer@berufskolleg-borken.de"));
            
            email.Subject = subject;

            if(!string.IsNullOrEmpty(cc))
            {
                email.Cc.Add(new MailboxAddress("Empfänger", cc));
            }
            
            if(!string.IsNullOrEmpty(bcc))
            {
                email.Bcc.Add(new MailboxAddress("Empfänger", bcc));
            }
            

            // 1️⃣ Erstelle den Haupttext der E-Mail
            var textPart = new TextPart("plain") { Text = body };

            // 2️⃣ Falls eine Datei angegeben wurde, erstelle den Anhang
            var multipart = new Multipart("mixed");
            multipart.Add(textPart); // Erst den Text hinzufügen

            if (!string.IsNullOrEmpty(attachment) && System.IO.File.Exists(attachment))
            {
                var attachmentPart = new MimePart()
                {
                    Content = new MimeContent(System.IO.File.OpenRead(attachment)),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = System.IO.Path.GetFileName(attachment)
                };

                multipart.Add(attachmentPart);
            }

            // 3️⃣ Setze den E-Mail-Body auf multipart (Text + Anhang)
            email.Body = multipart;

            using (var smtpClient = new MailKit.Net.Smtp.SmtpClient())
            {
                smtpClient.ServerCertificateValidationCallback = (s, c, h, e) => true; // SSL-Zertifikatsvalidierung deaktivieren
                smtpClient.Connect(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
                smtpClient.Authenticate(senderEmail, senderPassword);
                smtpClient.Send(email);
                smtpClient.Disconnect(true);
            }

            Global.ZeileSchreiben(receiverEmail, "gesendet", ConsoleColor.Green, ConsoleColor.White);        
        }
        catch(Exception ex){
            Global.ZeileSchreiben(receiverEmail, "Versand gescheitert.", ConsoleColor.Red,ConsoleColor.Gray);        
            Console.WriteLine("Fehler beim Versand der E-Mail an " + receiverEmail + ": " + ex.Message);
        }
    }

    public void SendenMitEingebettetemBild(string subject, string bild, List<string> bcc, IConfiguration configuration)
    {
        try
        {            
            string smtpServer = configuration["SmtpServerMassenmail"] ;
            int smtpPort = Convert.ToInt32(configuration["SmtpPort"]);
            string smtpUser = configuration["SmtpUserMassenmail"];
            string smtpPassword = configuration["SmtpPasswordMassenmail"];            

            var email = new MimeMessage();
            email.From.Add(new MailboxAddress(smtpUser, smtpUser));
            email.To.Add(new MailboxAddress("Empfänger", smtpUser));
            email.Subject = subject;

            foreach (var bccEmail in bcc)
            {
                // Wenn die Mail NON-Ascii enthält, dann wird sie nicht gesendet.
                if(!(bccEmail.Contains("ä") || bccEmail.Contains("ö") || bccEmail.Contains("ü") || bccEmail.Contains("ß")))
                {                
                    email.Bcc.Add(new MailboxAddress("Empfänger", bccEmail));
                }
            }

            // Erstelle den HTML-Body mit eingebettetem Bild
            var builder = new BodyBuilder();

            if (!string.IsNullOrEmpty(bild) && File.Exists(bild))
            {
                // Bild als eingebettete Ressource hinzufügen
                var image = builder.LinkedResources.Add(bild);
                image.ContentId = MimeUtils.GenerateMessageId();

                // HTML-Inhalt mit dem eingebetteten Bild
                builder.HtmlBody = $"<html><body><img src=\"cid:{image.ContentId}\" alt=\"Flyer Campusfest Berufskolleg Borken\" /></body></html>";
            }
            else
            {
                builder.HtmlBody = "<html><body><p>Bild konnte nicht geladen werden.</p></body></html>";
                Console.WriteLine($"Bild {bild} konnte nicht gefunden werden.");
                Console.ReadKey();
            }

            email.Body = builder.ToMessageBody();

            // Sende die E-Mail
            using (var smtpClient = new MailKit.Net.Smtp.SmtpClient())
            {
                smtpClient.ServerCertificateValidationCallback = (s, c, h, e) => true; // SSL-Zertifikatsvalidierung deaktivieren
                smtpClient.Connect(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
                smtpClient.Authenticate(smtpUser, smtpPassword);
                smtpClient.Send(email);
                smtpClient.Disconnect(true);
            }

            Console.WriteLine($"E-Mail an {bcc[0]} und {bcc.Count - 1} weitere BCC-Emfängern gesendet.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler beim Senden der E-Mail: {ex.Message}");
            Console.ReadKey();
        }
    }

    public void Senden(string subject, string sender, string body, Stream attachmentStream, string attachmentName, string receiver)
    {
        var mailMessage = new MailMessage(sender, receiver, subject, body);
        mailMessage.Attachments.Add(new Attachment(attachmentStream, attachmentName));

        // Füge den Absender als CC hinzu
        mailMessage.CC.Add(sender);

        if(Global.SmtpPassword == null || Global.SmtpPassword.Length <= 3)
        {
            Console.WriteLine("Bitte geben Sie das Passwort von " + Global.SmtpUser +" für den E-Mail-Versand ein:");
            Global.SmtpPassword = Console.ReadLine();
        }

        using (var smtpClient = new System.Net.Mail.SmtpClient(Global.SmtpServer, Convert.ToInt32(Global.SmtpPort)))
        {
            smtpClient.Credentials = new NetworkCredential(Global.SmtpUser, Global.SmtpPassword);
            smtpClient.EnableSsl = true;
            smtpClient.Send(mailMessage);
        }
    }
}