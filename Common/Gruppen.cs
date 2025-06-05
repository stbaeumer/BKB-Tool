using System.Dynamic;
using System.Text.RegularExpressions;
using Spectre.Console;

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

public class Gruppen : List<Gruppe>
{
    public Gruppen()
    {
    }

    public Gruppen GetBildungsgaenge(List<dynamic> exportLessons, Klassen klassen, Anrechnungen anrechnungs,
        Lehrers lehrers)
    {        
        var gruppen = new Gruppen();

        var records = new List<dynamic>();
        var bildungsgaengeWikiLinks = (from a in anrechnungs
            where a.Text.Contains("Bildungsgangleitung")
            where a.Beschr.StartsWith("bildungsgaenge:")
            select a.Beschr).Distinct().ToList().OrderBy(x => x);

        foreach (var b in bildungsgaengeWikiLinks)
        {
            dynamic record = new ExpandoObject();
            var kurzname = GetKurzname(b);
            var wikiLink = b;
            var schulform = GetSchulform(b);

            var members = GetMembers(exportLessons, lehrers, new List<int>() { 1, 2, 3, 4 }, kurzname);
            record.Page = wikiLink;
            var enumerable = members.ToList();
            record.Mitglieder = string.Join(',',
                enumerable.Select(x => (x.Titel != " " ? x.Titel : "") + x.Vorname + " " + x.Nachname));
            record.MitgliederMail = string.Join(',', enumerable.Select(x => x.Mail));
            record.MitgliederKuerzel = string.Join(',', enumerable.Select(x => x.Kürzel));
            var gruppe = new Gruppe(kurzname);
            records.Add(record);
            gruppe.Record = record;
            gruppen.Add(gruppe);
        }

        var panel = new Panel("Die Bildungsgänge werden aus den Anrechnungen ermittelt. Zwei Dinge sind wichtig:\n1. Das Wort 'Bildungsgangleitung' muss im [bold blue]Text[/] in Untis enthalten sein\n2. einen Link 'bildungsgaenge: ...' in der [bold blue]Beschr[/].")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Red);

                AnsiConsole.Write(panel);

        //Global.ZeileSchreiben("Bildungsgang-Gruppen", records.Count().ToString(), ConsoleColor.Black,ConsoleColor.White);

        return gruppen;
    }

    public Gruppen GetSchulformen(List<dynamic> exportLessons, Klassen klassen, Anrechnungen anrechnungs,
        Lehrers lehrers)
    {
        var gruppen = new Gruppen();
        
        var panel = new Panel("Die Schulformen werden aus den Anrechnungen ermittelt. Der Name der Schulformen wird aus der Untis-Beschr. ausgelesen")
                        .HeaderAlignment(Justify.Left)
                        .SquareBorder()
                        .Expand()
                        .BorderColor(Color.Red);

                AnsiConsole.Write(panel);


        var schulformen = (from a in anrechnungs
            where a.Text.Contains("Bildungsgangleitung")
            where a.Beschr.StartsWith("bildungsgaenge:")
            select GetSchulform(a.Beschr)).Distinct().ToList().OrderBy(x => x);

        // Für jede Schulform
        foreach (var schulform in schulformen)
        {
            var lehrerKürzel = new List<string>();
            var lehrerMail = new List<string>();
            var lehrerName = new List<string>();

            // und jeden enthaltenen Bildungsgang 
            foreach (var kurzname in (from a in anrechnungs
                         where a.Text.Contains("Bildungsgangleitung")
                         where GetSchulform(a.Beschr) == schulform
                         select GetKurzname(a.Beschr)).Distinct().ToList())
            {
                var members = GetMembers(exportLessons, lehrers, new List<int>() { 1, 2, 3, 4 }, kurzname);

                foreach (var member in members)
                {
                    if (!lehrerKürzel.Any(x => x.Contains(member.Kürzel)))
                    {
                        lehrerKürzel.Add(member.Kürzel);
                    }

                    if (!lehrerMail.Any(x => x.Contains(member.Mail)))
                    {
                        lehrerMail.Add(member.Mail);
                    }

                    if (!lehrerName.Any(x =>
                            x.Contains((member.Titel == "" ? "" : member.Titel + " ") + member.Vorname + " " +
                                       member.Nachname)))
                    {
                        lehrerName.Add((member.Titel == "" ? "" : member.Titel + " ") + member.Vorname + " " +
                                       member.Nachname);
                    }
                }
            }

            dynamic record = new ExpandoObject();
            record.Page = "bildungsgaenge:" + schulform + ":start";
            record.Mitglieder = string.Join(',', lehrerName);
            record.MitgliederMail = string.Join(',', lehrerMail);
            record.MitgliederKuerzel = string.Join(',', lehrerKürzel);
            Gruppe gruppe = new Gruppe(schulform);
            gruppe.Record = record;
            gruppen.Add(gruppe);
        }

        //Global.ZeileSchreiben("Schulform-Gruppen", gruppen.Count().ToString(), ConsoleColor.Black, ConsoleColor.White);

        return gruppen;
    }


    public IEnumerable<Lehrer> GetMembers(List<dynamic> exportLessons, Lehrers lehrers, List<int> jahrgänge,
        string kurzname)
    {
        var members = new List<Lehrer>();

        foreach (var recExp in exportLessons)
        {
            var dict = (IDictionary<string, object>)recExp;

            var klassenkürzels = new List<string>();
            klassenkürzels.AddRange(dict["klassen"].ToString().Split('~'));

            foreach (var klassenkürzel in klassenkürzels)
            {
                if ((klassenkürzel.ToLower().StartsWith(kurzname.ToLower()) && klassenkürzel.Any(c => char.IsDigit(c))))
                {
                    if (JahrgangPasst(klassenkürzel, jahrgänge))
                    {
                        var leh = (from l in lehrers where l.Kürzel == dict["teacher"].ToString() select l)
                            .FirstOrDefault();

                        if (leh != null && !(from l in members where l.Mail == leh.Mail select l).Any())
                        {
                            if (!(from m in members where m.Mail == leh.Mail select m).Any())
                            {
                                members.Add(leh);
                            }
                        }
                    }
                }
            }
        }

        return members;
    }

    private bool JahrgangPasst(string klassenkürzel, List<int> jahrgänge)
    {
        foreach (var jahrgang in jahrgänge)
        {
            if (ExtractNumber(klassenkürzel) == (Convert.ToInt32(Global.AktSj[0]) - 2000 + 1 - jahrgang) ||
                klassenkürzel.Contains("FM2"))
            {
                return true;
            }
        }

        return false;
    }

    private int ExtractNumber(string input)
    {
        // Der reguläre Ausdruck, der eine oder mehrere Ziffern findet
        Match match = Regex.Match(input, @"\d+");

        // Wenn ein Match gefunden wird, gib die übereinstimmende Zahl zurück
        if (match.Success)
        {
            return Convert.ToInt32(match.Value);
        }

        // Rückgabe eines leeren Strings, wenn keine Zahl gefunden wird
        return -99;
    }

    private static string GetSchulform(string input)
    {
        var firstColonIndex = input.IndexOf(':');
        if (firstColonIndex < 0)
        {
            return string.Empty; // Rückgabe eines leeren Strings, wenn kein Doppelpunkt vorhanden ist
        }

        var secondColonIndex = input.IndexOf(':', firstColonIndex + 1);
        return secondColonIndex < 0
            ? string.Empty
            : // Rückgabe eines leeren Strings, wenn nur ein Doppelpunkt vorhanden ist
            input.Substring(firstColonIndex + 1, secondColonIndex - firstColonIndex - 1);
    }

    public string GetKurzname(string input)
    {
        int lastColonIndex = input.LastIndexOf(':');
        if (lastColonIndex < 0)
        {
            return string.Empty; // Rückgabe eines leeren Strings, wenn kein Doppelpunkt vorhanden ist
        }

        int secondLastColonIndex = input.LastIndexOf(':', lastColonIndex - 1);
        if (secondLastColonIndex < 0)
        {
            return string.Empty; // Rückgabe eines leeren Strings, wenn nur ein Doppelpunkt vorhanden ist
        }

        return input.Substring(secondLastColonIndex + 1, lastColonIndex - secondLastColonIndex - 1).ToUpper();
    }

    public IEnumerable<dynamic> Get(List<dynamic> exportlessons, Klassen klassen, Anrechnungen anrechnungen,
        Lehrers lehrers, string wikiLink, List<string> beteiligteKlassen, List<int> jahrgänge)
    {
        List<string> jahre = new List<string>();
        var aktSj = Convert.ToInt32(Global.AktSj[0]);
        foreach (var jahrgang in jahrgänge)
        {
            jahre.Add((aktSj - 2000 - jahrgang + 1).ToString());
        }

        var records = new List<dynamic>();
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

        record.Mitglieder = string.Join(',', lehrerName);
        record.MitgliederMail = string.Join(',', lehrerMail);
        record.MitgliederKuerzel = string.Join(',', lehrerKürzel);
        records.Add(record);
        return records;
    }
}