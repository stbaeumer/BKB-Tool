using System.Dynamic;

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

    public Gruppe Get(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen,
        Lehrers lehrers, string wikiLink, List<string> beteiligteKlassen, List<int> jahrgänge)
    {
        var gruppe = new Gruppe(wikiLink);
        var aktSj = Convert.ToInt32(Global.AktSj[0]);
        Lehrers = new Lehrers();
        
        var jahre = jahrgänge.Select(jahrgang => (aktSj - 2000 - jahrgang + 1).ToString()).ToList();

        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = exportlessons.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            var dictKlassen = dict["klassen"].ToString().Split('~');

            // Prüfe, ob eine der Klassen in subject mit einem der Präfixe in beteiligteKlassen beginnt.
            return dictKlassen.Any(klasse => beteiligteKlassen.Any(prefix => klasse.Contains(prefix)));
        }).Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            var dictKlassen = dict["klassen"].ToString().Split('~');

            // Prüfe, ob eine der Klassen die Jahreszahl zum Jahrgang enthält.
            return !string.IsNullOrEmpty(dict["teacher"].ToString()) &&
                   dictKlassen.Any(subject => jahre.Any(jahr => subject.Contains(jahr)));
        }).Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["teacher"].ToString();
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

    public Gruppe GetFachschaft(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink, List<string> faecher)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = exportlessons.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            var dictSubject = dict["subject"].ToString().Split('~');

            // Prüfe, ob eine der Klassen in subject mit einem der Präfixe in beteiligteKlassen beginnt.
            return dictSubject.Any(fach => faecher.Any(x => fach == x));
        }).Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["teacher"].ToString();
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

    public Gruppe GetKollegium(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen, Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = exportlessons.Where(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return !string.IsNullOrEmpty(dict["teacher"].ToString());
        }).Select(rec =>
        {
            var dict = (IDictionary<string, object>)rec;
            return dict["teacher"].ToString();
        }).Distinct().OrderBy(x => x);


        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
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
                lehrerName.Add((leh.Titel == "" ? "" : leh.Titel + " ") + leh.Vorname + " " +
                               leh.Nachname);
            }
        }

        record.Mitglieder = string.Join(',', lehrerName.OrderBy(name => name));
        record.MitgliederMail = string.Join(',', lehrerMail.OrderBy(name => name));
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel.OrderBy(name => name));
        gruppe.Record = record;
        return gruppe;
    }

    public Gruppe GetLehrerinnen(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = anrechnungen.Select(rec => { return rec.Lehrer.Kürzel; }).Distinct().OrderBy(x => x);


        foreach (var member in members)
        {
            if (lehrers.Any(l => l.Kürzel == member && l.Geschlecht.ToLower() == "w"))
            {
                var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
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

    public Gruppe GetRefs(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen, Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = lehrers.Where(rec => { return rec.Kürzel.StartsWith("Y"); }).Select(rec => { return rec.Kürzel; })
            .Distinct();


        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
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

    public Gruppe GetKlassenleitungen(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = klassen
            .SelectMany(rec => rec.Klassenleitungen) // Alle Lehrer aus allen Klassenleitungen extrahieren
            .Select(lehrer => lehrer.Kürzel) // Nur das Kürzel jedes Lehrers auswählen
            .Distinct(); // Doppelte Kürzel entfernen

        foreach (var member in members)
        {
if(member == "GV"){
    string a="";
}

            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
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


    public Gruppe GetBildungsgangleitungen(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = anrechnungen.Where(rec => { return rec.Text.Contains("Bildungsgangleitung"); })
            .Select(rec => { return rec.Lehrer.Kürzel; }).Distinct().OrderBy(x => x);


        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
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

    public Gruppe GetByWikilink(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen,
        Lehrers lehrers,
        string wikiLink)
    {
        var gruppe = new Gruppe(wikiLink);
        dynamic record = new ExpandoObject();
        record.Page = wikiLink;

        var lehrerKürzel = new List<string>();
        var lehrerMail = new List<string>();
        var lehrerName = new List<string>();

        var members = anrechnungen.Where(rec => { return rec.Beschr.Contains(wikiLink); })
            .Select(rec => { return rec.Lehrer.Kürzel; }).Distinct().OrderBy(x => x);

        foreach (var member in members)
        {
            var leh = lehrers.FirstOrDefault(l => l.Kürzel == member);
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