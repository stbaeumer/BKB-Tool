using System.Dynamic;
using Microsoft.Extensions.Configuration;

namespace Common;

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

public class Gruppe
{
    public dynamic Record { get; set; }
    public string Name { get; set; }

    public Gruppe(string kurzname)
    {
        Name = kurzname;
    }

    public Gruppe()
    {
    }

    public Gruppe Get(List<dynamic> gpu020,
        Lehrers lehrers, string wikiLink, List<string> beteiligteKlassen, List<int> jahrgänge)
    {
        var gruppe = new Gruppe(wikiLink);
        var aktSj = Convert.ToInt32(Global.AktSj[0]);
        Lehrers = new Lehrers();
        
        var jahre = jahrgänge.Select(jahrgang => (aktSj - 2000 - jahrgang + 1).ToString()).ToList();

        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = gpu020.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            var dictKlassen = dict["Field5"].ToString().Split('~');

            // Prüfe, ob eine der Klassen in subject mit einem der Präfixe in beteiligteKlassen beginnt.
            return dictKlassen.Any(klasse => beteiligteKlassen.Any(prefix => klasse.Contains(prefix)));
        }).Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            var dictKlassen = dict["Field5"].ToString().Split('~');

            // Prüfe, ob eine der Klassen die Jahreszahl zum Jahrgang enthält.
            return !string.IsNullOrEmpty(dict["Field6"].ToString()) &&
                   dictKlassen.Any(subject => jahre.Any(jahr => subject.Contains(jahr)));
        }).Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Field6"].ToString();
        }).Distinct().OrderBy(x => x);

        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
            if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

            if (!lehrerKürzel.Any(x => x.Contains(leh.Kürzel)))
                lehrerKürzel.Add(leh.Kürzel);

            if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
                lehrerMail.Add(leh.Mail);

            if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
            {
                lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
            }
        }

        record.Namen = string.Join(", ", lehrerName);
        record.Mail = string.Join("; ", lehrerMail);
        record.Kürzel = string.Join(", ", lehrerKürzel);
        record.Art = ":schulgemeinschaft:gruppen";
        gruppe.Record = record;

        //Global.ZeileSchreiben("Gruppe: " + wikiLink, lehrerName.Count().ToString(), ConsoleColor.Black, ConsoleColor.White);

        return gruppe;
    }

    public Lehrers Lehrers { get; set; }

    public Gruppe GetFachschaft(List<dynamic> gpu002,
        Lehrers lehrers,
        string wikiLink, List<string> faecher)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = gpu002.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            var dictSubject = dict["Field7"].ToString().Split('~');

            // Prüfe, ob eine der Klassen in subject mit einem der Präfixe in beteiligteKlassen beginnt.
            return dictSubject.Any(fach => faecher.Any(x => fach == x));
        }).Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["Field6"].ToString();
        }).Distinct().OrderBy(x => x);

        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
            if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen
            
            if (!lehrerKürzel.Any(x => x.Contains(leh.Kürzel)))
            {
                lehrerKürzel.Add(leh.Kürzel);
            }

            if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
            {
                lehrerMail.Add(leh.Mail);
            }

            if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
            {
                lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
            }
        }

        record.Namen = string.Join(", ", lehrerName.OrderBy(name => name));
        record.Mail = string.Join("; ", lehrerMail.OrderBy(name => name));
        record.Kürzel = string.Join(", ", lehrerKürzel.OrderBy(name => name));
        record.Art = ":schulgemeinschaft:gruppen";
        gruppe.Record = record;

        //Global.ZeileSchreiben("Gruppe: " + wikiLink, lehrerName.Count().ToString(), ConsoleColor.Black, ConsoleColor.White);

        return gruppe;
    }

    public Gruppe GetKollegium(List<dynamic> gpu002, Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = gpu002
        .Select(rec => ((IDictionary<string, object>)rec)["Field6"]?.ToString())
        .Where(field6 => !string.IsNullOrEmpty(field6))
        .Where(field6 => field6 != "?")
        .Distinct()
        .OrderBy(field6 => field6)
        .ToList();

        try
        {
            foreach (var member in members)
            {
                var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);

                if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

                if (!lehrerKürzel.Any(x => x.Contains(leh.Kürzel)))
                {
                    lehrerKürzel.Add(leh.Kürzel);
                }

                if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
                {
                    lehrerMail.Add(leh.Mail);
                }

                if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
                {
                    lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
                }
            }
        }
        catch
        {
            throw new Exception("Fehler beim Verarbeiten der Kollegium-Gruppe");
        }

        record.Namen = string.Join(", ", lehrerName.OrderBy(name => name));
        record.Mail = string.Join("; ", lehrerMail.OrderBy(name => name));
        record.Kürzel = string.Join(", ", lehrerKürzel.OrderBy(name => name));
        record.Art = ":schulgemeinschaft:gruppen";
        gruppe.Record = record;
        return gruppe;
    }

    public Gruppe GetLehrerinnen(Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

 var members = anrechnungen
            .Where(rec => rec.Lehrer != null)
            .Select(rec => { return rec.Lehrer.Kürzel; })
            .Distinct()
            .OrderBy(x => x);


        foreach (var member in members)
        {
            if (lehrers.Any(l => l.Kürzel == member && l.Geschlecht.ToLower() == "w"))
            {
                var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);

                if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

                if (!lehrerKürzel.Any(x => x.Contains(leh.Kürzel)))
                {
                    lehrerKürzel.Add(leh.Kürzel);
                }

                if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
                {
                    lehrerMail.Add(leh.Mail);
                }

                if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
                {
                    lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
                }
            }
        }

        record.Namen = string.Join(", ", lehrerName.OrderBy(name => name));
        record.Mail = string.Join("; ", lehrerMail.OrderBy(name => name));
        record.Kürzel = string.Join(", ", lehrerKürzel.OrderBy(name => name));
        record.Art = "schulgemeinschaft:gruppen";
        gruppe.Record = record;
        return gruppe;
    }

    public Gruppe GetRefs(Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = lehrers.Where(rec => { return rec.Beschäftigungsart.StartsWith("ST"); }).Select(rec => { return rec.Kürzel; })
            .Distinct();


        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);

            if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

            if (!lehrerKürzel.Any(x => x.Contains(leh.Kürzel)))
            {
                lehrerKürzel.Add(leh.Kürzel);
            }

            if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
            {
                lehrerMail.Add(leh.Mail);
            }

            if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
            {
                lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
            }
        }

        record.Namen = string.Join(',', lehrerName.OrderBy(name => name));
        record.Mail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.Kürzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
        record.Art = "schulgemeinschaft:gruppen";
        gruppe.Record = record;
        return gruppe;
    }

    public Gruppe GetKlassenleitungen(List<dynamic> gpu003,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = gpu003
            .Where(rec => rec != null)
            .Select(rec => ((IDictionary<string, object>)rec)["Field30"]?.ToString())
            .Where(field30 => !string.IsNullOrEmpty(field30))
            .Select(field30 => field30.Split(',')[0]) // Nur den Teil vor dem Komma nehmen
            .Distinct()
            .OrderBy(field30 => field30)
            .ToList();

        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);

            if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

            if (!lehrerKürzel.Any(x => x == leh.Kürzel)) // Exakte Übereinstimmung prüfen
            {
                lehrerKürzel.Add(leh.Kürzel);
            }

            if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
            {
                lehrerMail.Add(leh.Mail);
            }

            if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
            {
                lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
            }
        }

        record.Namen = string.Join(", ", lehrerName.OrderBy(name => name));
        record.Mail = string.Join("; ", lehrerMail.OrderBy(name => name));
        record.Kürzel = string.Join(", ", lehrerKürzel.OrderBy(name => name));
        record.Art = ":schulgemeinschaft:gruppen";
        gruppe.Record = record;
        return gruppe;
    }


    public Gruppe GetBildungsgangleitungen(Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;
        
        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();
        
        var members = anrechnungen.Where(rec => { return rec.Text.Contains("Bildungsgangleitung"); })
            .Select(rec => { return rec.LehrerKuerzel; }).Distinct().OrderBy(x => x);

        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);

            if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

            if (!lehrerKürzel.Any(x => x.Contains(leh.Kürzel)))
            {
                lehrerKürzel.Add(leh.Kürzel);
            }

            if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
            {
                lehrerMail.Add(leh.Mail);
            }

            if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
            {
                lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
            }
        }

        record.Namen = string.Join(", ", lehrerName.OrderBy(name => name));
        record.Mail = string.Join("; ", lehrerMail.OrderBy(name => name));
        record.Kürzel = string.Join(", ", lehrerKürzel.OrderBy(name => name));
        record.Art = ":schulgemeinschaft:gruppen";
        gruppe.Record = record;
        return gruppe;
    }

    public Gruppe GetByWikilink(Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;
        record.Link = wikiLink;
    
        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = anrechnungen.Where(rec => { return rec.Beschr.Contains(wikiLink); })
            .Select(rec => { return rec.Lehrer.Kürzel; }).Distinct().OrderBy(x => x);

        if (members == null || !members.Any())
        {
            members = anrechnungen.Where(rec => { return rec.Beschr.Contains(wikiLink.Replace("schulgemeinschaft:", "kollegium:")); })
            .Select(rec => { return rec.Lehrer.Kürzel; }).Distinct().OrderBy(x => x);
        }

        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);

            if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

            if (!lehrerKürzel.Any(x => x.Contains(leh.Kürzel)))
            {
                lehrerKürzel.Add(leh.Kürzel);
            }

            if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
            {
                lehrerMail.Add(leh.Mail);
            }

            //if (!lehrerName.Any(x => x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname)))
            if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower())))
            {
                record.TitelVornameNachname = (String.IsNullOrEmpty(leh.Titel) ? $"{leh.Vorname} {leh.Nachname}" : $"{leh.Titel} {leh.Vorname} {leh.Nachname}");                
                //lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
                lehrerName.Add(":schulgemeinschaft:" + leh.Kürzel.ToLower());
            }
        }

        record.Namen = string.Join(", ", lehrerName.OrderBy(name => name));
        record.Mail = string.Join("; ", lehrerMail.OrderBy(name => name));
        record.Kürzel = string.Join(", ", lehrerKürzel.OrderBy(name => name));
        record.Art = "schulgemeinschaft:gruppen";
        gruppe.Record = record;
        return gruppe;
    }

 
  public Gruppen GetKlassen(List<dynamic> gpu002, List<dynamic> gpu003,
        Lehrers lehrers,
        Students students,
        Anrechnungen anrechnungen,
        string wikiLink)
    {
        var gruppen = new Gruppen();

        var alleVerschiedenenKlassen = gpu002
    .Cast<IDictionary<string, object>>()
    .Select(rec => rec.ContainsKey("Field5") ? rec["Field5"]?.ToString() : null)
    .Where(klasse => !string.IsNullOrEmpty(klasse) && klasse != "?")
    .Distinct()
    .OrderBy(klasse => klasse)
    .ToList();

        foreach(var k in alleVerschiedenenKlassen)
        {
            var lehrerInDerKlasse = gpu002
                .Cast<IDictionary<string, object>>()
                .Where(rec => rec.ContainsKey("Field5") && rec["Field5"]?.ToString() == k)
                .Select(rec => rec.ContainsKey("Field6") ? rec["Field6"]?.ToString() : null)
                .Where(lehrer => !string.IsNullOrEmpty(lehrer) && lehrer != "?")
                .Distinct()
                .OrderBy(lehrer => lehrer)
                .ToList();


            var gruppe = new Gruppe(wikiLink);
            dynamic record = new ExpandoObject();
            record.Page = wikiLink + ":" + k;
            record.Link = wikiLink + ":" + k;

            var lehrerKürzel = new List<string>();
            var lehrerMail = new List<string>();
            var lehrerName = new List<string>();

            foreach (var member in lehrerInDerKlasse)
            {
                var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);

                if (leh == null) continue; // Wenn kein Lehrer gefunden, nächsten Eintrag ansehen

                if (!lehrerKürzel.Any(x => x == leh.Kürzel)) // Exakte Übereinstimmung prüfen
                {
                    lehrerKürzel.Add(leh.Kürzel);
                }

                if (!lehrerMail.Any(x => x.Contains(leh.Mail)))
                {
                    lehrerMail.Add(leh.Mail);
                }

                if (!lehrerName.Any(x => x.Contains(":schulgemeinschaft:" + leh.Kürzel.ToLower()))) // Prüfen, ob der Name bereits in der Liste ist
                {
                    lehrerName.Add((":schulgemeinschaft:" + leh.Kürzel.ToLower()));
                }
            }

            record.Klasse = k;
            record.Namen = string.Join(", ", lehrerName.OrderBy(name => name));
            record.Mail = string.Join("; ", lehrerMail.OrderBy(name => name));
            record.Kürzel = string.Join(", ", lehrerKürzel.OrderBy(name => name));
            record.Art = ":klassen:start";

            var schuelerDerKlasse = "";

            var sus = students            
            .Where(s => !string.IsNullOrEmpty(s.Klasse)) // Optional: Leere/Null-Werte herausfiltern
            .Where(s => s.Klasse == k)
            .Distinct()
            .OrderBy(s => s.Nachname)
            .ToList();

            foreach(var s in sus)
            {
                schuelerDerKlasse += s.Vorname + " " + s.Nachname.Split('#')[0] + ", ";
            }

            record.KlasseSus = schuelerDerKlasse.TrimEnd(',').TrimEnd(' ').TrimEnd(',').TrimEnd(' ');

            var klassenLeitung = gpu003
    .Cast<IDictionary<string, object>>()
    .Where(rec => rec.ContainsKey("Field1") && rec["Field1"]?.ToString() == k)
    .Select(rec => rec.ContainsKey("Field30") ? rec["Field30"]?.ToString() : null)
    .Where(field30 => !string.IsNullOrEmpty(field30))
    .Select(field30 => field30.Split(',')[0].Trim().ToLower()) // Ersten Lehrer ermitteln & für Wiki-Namespace säubern
    .FirstOrDefault();

record.VorsitzLeitung = !string.IsNullOrEmpty(klassenLeitung) 
    ? "schulgemeinschaft:" + klassenLeitung 
    : string.Empty;

            
            string klassenName = k.Trim().ToLower(); // z.B. "bs26a", "hbg24b", "bt23a"
            var match = System.Text.RegularExpressions.Regex.Match(klassenName, @"^([a-z]+)(\d{2})([a-z0-9]*)$");

            if (match.Success)
            {
                string bildungsgang = match.Groups[1].Value; // z. B. "bs", "hbg"
                int einschulungsJahr = int.Parse(match.Groups[2].Value); // z. B. 26, 25, 24, 23

                // Bestimmung des aktuellen Schuljahr-Startjahres (ab August beginnt das neue Schuljahr)
                DateTime heute = DateTime.Now;
                int aktuellesSchuljahrStart = (heute.Month >= 8) ? (heute.Year % 100) : ((heute.Year - 1) % 100);

                // Jahrgang berechnen (1 bis 4)
                int jahrgangsStufe = (aktuellesSchuljahrStart - einschulungsJahr) + 1;

                if (jahrgangsStufe >= 1 && jahrgangsStufe <= 4)
                {
                    string targetNamespace = $"djp:{bildungsgang}:jg{jahrgangsStufe}:start";

                    switch (jahrgangsStufe)
                    {
                        case 1: record.DJP1 = targetNamespace; break;
                        case 2: record.DJP2 = targetNamespace; break;
                        case 3: record.DJP3 = targetNamespace; break;
                        case 4: record.DJP4 = targetNamespace; break;
                    }
     
                    string targetEnd = $"{bildungsgang}:start";   // z. B. "bs:start"

                    // 2. Anrechnung suchen
                    var passendeAnrechnung = anrechnungen
                        .Where(x => !string.IsNullOrEmpty(x.Beschr))
                        .FirstOrDefault(x => x.Beschr.StartsWith("bildungsgaenge:", StringComparison.OrdinalIgnoreCase) 
                                        && x.Beschr.EndsWith(targetEnd, StringComparison.OrdinalIgnoreCase));

                    if (passendeAnrechnung != null)
                    {
                        record.BGSeite = passendeAnrechnung.Beschr;
                    }
                }
            }





            gruppe.Record = record;
            gruppen.Add(gruppe);
        }
        
        return gruppen;
    }
 }
