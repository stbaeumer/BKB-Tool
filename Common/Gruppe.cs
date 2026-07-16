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

            if (!lehrerName.Any(x =>
                    x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname)))
                lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
        }

        record.Mitglieder = string.Join(',', lehrerName);
        record.MitgliederMail = string.Join(',', lehrerMail);
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel);
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

            if (!lehrerName.Any(x => x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " +
                                                leh.Nachname)))
            {
                lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " +
                               leh.Nachname);
            }
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
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

                if (!lehrerName.Any(x =>x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname)))
                {
                    lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
                }
            }
        }
        catch
        {
            throw new Exception("Fehler beim Verarbeiten der Kollegium-Gruppe");
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
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

                if (!lehrerName.Any(x =>
                        x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " +
                                   leh.Nachname)))
                {
                    lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
                }
            }
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
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

            if (!lehrerName.Any(x => x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname)))
            {
                lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
            }
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
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

            if (!lehrerName.Any(x =>
                    x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " +
                               leh.Nachname)))
            {
                lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
            }
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
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
            .Select(rec => { return rec.Lehrer.Kürzel; }).Distinct().OrderBy(x => x);


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

            if (!lehrerName.Any(x =>
                    x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " +
                               leh.Nachname)))
            {
                lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
            }
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
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

            if (!lehrerName.Any(x =>
                    x.Contains((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " +
                               leh.Nachname)))
            {
                lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " + leh.Nachname);
            }
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
        gruppe.Record = record;
        return gruppe;
    }
}