using System.Net;
using System.Net.Mail;
using MailKit.Net.Smtp;
using Microsoft.Extensions.Configuration;
using MimeKit;
using MimeKit.Utils;
using Spectre.Console;

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

 public void Senden(IConfiguration configuration, string subject, string body, List<string> to, List<string> cc, List<string> bcc, List<string> attachment) 
 {
  try
  {
    var panel = new Panel($"[green]An:  {string.Join(", ", to)}[/] \nCC: {string.Join(", ", cc)} \nBCC: {string.Join(", ", bcc)}\n\nBetreff: {subject}\n\n{body}")
    .Header("[bold red]  Mail jetzt wie angezeigt senden? [/]")
    .HeaderAlignment(Justify.Left)
    .SquareBorder()
    .Expand()
    .BorderColor(Color.Red);

    AnsiConsole.Write(panel);

    // Bestätige mit ENTER, Anykey für Abbruch
    Console.WriteLine("Mail wie angezeigt senden mit ENTER, Anykey für Abbruch ...");
    var keyInfo = Console.ReadKey();
    if (keyInfo.Key != ConsoleKey.Enter)
    {
     throw new Exception("Sie haben abgebrochen.");
    }

   AnsiConsole.Status().Spinner(Spinner.Known.Dots).Start("Mails senden ...", ctx =>
   {
    string smtpServer = configuration["SmtpServer"];
    int smtpPort = Convert.ToInt32(configuration["SmtpPort"]);
    string senderEmail = configuration["SmtpUser"];

    if (configuration["SmtpKennwort"] == null || configuration["SmtpKennwort"].Length <= 3)
    {
     Console.WriteLine("Bitte geben Sie das Passwort von " + configuration["SmtpUser"] + " für den E-Mail-Versand ein:");
     configuration["SmtpKennwort"] = Console.ReadLine();
    }

    string senderPassword = configuration["SmtpKennwort"];

    var email = new MimeMessage();
    email.From.Add(new MailboxAddress(configuration["SmtpUser"], senderEmail));

    email.Subject = subject;

    foreach(var m in to)
     email.To.Add(new MailboxAddress("Empfänger", m.Trim()));    
    
    foreach(var m in cc)
     email.Cc.Add(new MailboxAddress("Empfänger", m.Trim()));
    
    foreach(var m in bcc)
     email.Bcc.Add(new MailboxAddress("Empfänger", m.Trim()));
        
    var textPart = new TextPart("plain") { Text = body };

    // Falls eine Datei angegeben wurde, erstelle den Anhang
    var multipart = new Multipart("mixed");
    multipart.Add(textPart); // Erst den Text hinzufügen

    foreach(var attachment in attachment)
    {
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
    }
    
    // Setze den E-Mail-Body auf multipart (Text + Anhang)
    email.Body = multipart;

    using (var smtpClient = new MailKit.Net.Smtp.SmtpClient())
    {
     smtpClient.ServerCertificateValidationCallback = (s, c, h, e) => true; // SSL-Zertifikatsvalidierung deaktivieren
     smtpClient.Connect(smtpServer, smtpPort, MailKit.Security.SecureSocketOptions.StartTls);
     smtpClient.Authenticate(senderEmail, senderPassword);
     //smtpClient.Send(email);
     smtpClient.Disconnect(true);
    }
   });
  }
  catch(Exception ex)
  {
   throw;
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
                while (Console.KeyAvailable) Console.ReadKey(true);

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
            while (Console.KeyAvailable) Console.ReadKey(true);

            Console.ReadKey();
        }
    }

    public void Senden(IConfiguration configuration, string subject, string sender, string body, Stream attachmentStream, string attachmentName, string receiver)
    {
        try
        { 
            var mailMessage = new MailMessage(sender, receiver, subject, body);
            mailMessage.Attachments.Add(new Attachment(attachmentStream, attachmentName));

            if(!string.IsNullOrEmpty(configuration["BCCAdresse"]) && configuration["BCCAdresse"].Contains("@"))
            {
                mailMessage.Bcc.Add(configuration["BCCAdresse"]);
            }
            
            if(configuration["SmtpKennwort"]  == null || configuration["SmtpKennwort"].Length <= 3)
            {
                Console.WriteLine($"Bitte geben Sie das Passwort von {configuration["SmtpUser"]} für den E-Mail-Versand ein:");
                Global.SmtpKennwort = Console.ReadLine();
            }

            using (var smtpClient = new System.Net.Mail.SmtpClient(configuration["SmtpServer"], Convert.ToInt32(configuration["SmtpPort"])))
            {
                smtpClient.Credentials = new NetworkCredential(sender, configuration["SmtpKennwort"]);
                smtpClient.EnableSsl = true;
                smtpClient.Send(mailMessage);
                smtpClient.Dispose();
                mailMessage.Dispose();                
                AnsiConsole.Write(new Panel($"[bold green]  Mail gesendet  [/]: [green]{receiver}[/]")
                    .SquareBorder()
                    .Expand()
                    .BorderColor(Color.Green));
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Fehler beim Senden der E-Mail an {receiver}: {ex.Message}");
        }
    }
}