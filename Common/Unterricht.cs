using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

public class Unterricht
{
    public string Kursleiter { get; set; }
    public string? Schülergruppe { get; set; }
    public List<int> UnterrichtsIds { get; set; }
    public int KursleiterWochenstunden { get; set; }
    public string? Fach { get; set; }
    public List<string> Klassen { get; internal set; }
    public List<string> Lehrkraefte { get; internal set; }
    public List<int> LehrkraefteWochenstunden { get; internal set; }
    public string KursBez { get; internal set; }
    public string Kursart { get; internal set; }
    public int Wochenstunden { get; internal set; }
    public List<string> Jahrgaenge { get; internal set; }
    public Students Students { get; set; }    
    public Unterricht(Global.Zweck zweck, Menüeintrag m, IConfiguration configuration, string? unterrichtsId, string fach, string? schuelergruppe, string? klasse, string? lehrer, int wochentundenLehrkraft, List<dynamic> studentgroupStudents)
    {
        Fach = Bereinigen(fach);
        Klassen = new List<string>() { klasse };
        Kursart = "PUK";
        Wochenstunden = wochentundenLehrkraft;
        Kursleiter = lehrer;
        KursleiterWochenstunden = wochentundenLehrkraft;
        UnterrichtsIds = new List<int> { int.Parse(unterrichtsId) };
        Students = m.IStudents.Filter(configuration, zweck, klasse, schuelergruppe, studentgroupStudents);
        Global.ZeileSchreiben($"{klasse} {fach} {lehrer}", $"Schüler*innen: {Students.Count}");                    
    }

    public Unterricht(Global.Zweck zweck, Menüeintrag m, string? unterrichtsId, string fach, string? schuelergruppe, string? klasse, string? lehrer, int wochentundenLehrkraft, List<dynamic> kurseDat, IConfiguration configuration, List<dynamic> studentgroupStudents) 
        : this(zweck, m, unterrichtsId, fach, schuelergruppe, klasse, lehrer, wochentundenLehrkraft, studentgroupStudents)
    {
        // Klassen nicht leer -> Alle bekommen den Kurs zugewiesen
        //                       KursBez darf leer bleiben
        // Klassen       leer -> Zuweisung aus StudentgroupStudents

        KursBez = $"{lehrer}-{unterrichtsId}";
        Fach = Bereinigen(fach);
        Klassen = new List<string>() { klasse };
        // Die Kursart wird aus der Kurse.dat ermittelt, wenn sie dort einmal gesetzt ist. Sie steckt in "Kursart" des Dictionaries.
        Kursart = GetKursart(configuration, kurseDat, fach, lehrer, unterrichtsId);
        Wochenstunden = wochentundenLehrkraft;
        Kursleiter = lehrer;
        KursleiterWochenstunden = wochentundenLehrkraft;
        Lehrkraefte = new List<string>();
        LehrkraefteWochenstunden = new List<int>();
        Schülergruppe = schuelergruppe;
        UnterrichtsIds = new List<int> { int.Parse(unterrichtsId) };
        Students = m.IStudents.Filter(configuration, zweck, klasse, schuelergruppe, studentgroupStudents);
        Global.ZeileSchreiben($"{KursBez}", $"");
    }

    private string GetKursart(IConfiguration configuration, List<dynamic> kurseDat, string fach, string? kursleiter, string? unterrichtsId)
    {
        List<string> kursarten = new List<string>() { "GK", "LK", "AB", "ZK", "VTF", "PJK" };
        List<string> unserekursarten = configuration["Kursarten"]?.Split(',').Select(s => s).ToList();

        var dictKurs = kurseDat
            .Select(k => k as IDictionary<string, object>)
            .FirstOrDefault(k => k != null && k.ContainsKey("Fach") &&
                k["Fach"]?.ToString() == fach &&
                k["Kursleiter"]?.ToString() == kursleiter &&
                k["KursBez"].ToString().Contains(unterrichtsId) &&
                k.ContainsKey("Kursart") && k["Kursart"] != null
                );
        var stringKurs = dictKurs?["Kursart"]?.ToString() ?? string.Empty;

        // Wenn die Kursart in der Kurse.dat bereits vorher gesetzt wurde, wird sie übernommen.        
        if (stringKurs != "")
            return stringKurs;

        if (unserekursarten == null || unserekursarten.Count != 6)
            return ""; // Fehlerfall, leere Rückgabe

        try
        {
            for (int i = 0; i < kursarten.Count; i++)
            {
                // Kursart in der Konfiguration ist in der Kurse.dat nicht gesetzt, aber in der Konfiguration vorhanden
                if (fach.Contains(unserekursarten[i]))
                {
                    return kursarten[i];
                }
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Fehler beim Zuordnen der Kursarten[/]");
            return ""; // Fehlerfall, leere Rückgabe
        }        
        return ""; // Wenn keine Kursart gefunden wurde, wird eine leere Zeichenkette zurückgegeben
    }

    private string Bereinigen(string? fach)
    {
        string fachBereinigt = fach;
        if (!string.IsNullOrEmpty(fach) && fach.Length > 1 && char.IsDigit(fach[^1]) && fach.Count(c => c == ' ') == 0)
        {
            // Wenn das Fach nur aus Buchstaben und einer Ziffer am Ende besteht (z.B. "M1"), entferne die Ziffer
            fachBereinigt = fach.Substring(0, fach.Length - 1);
        }
        return fachBereinigt;
    }

    internal void Updaten(Global.Zweck zweck, IConfiguration configuration, string fach, string lehrer, string unterrichtsId, string schuelergruppe, string klasse, int wochentundenLehrkraft, List<dynamic> studentgroupStudents)
    {
        // Wenn der Kurs bereits existiert und in einer zweiten Zeile eine weitere Klasse hinzugefügt wird
        if (KursBez.StartsWith(lehrer) && KursBez.Contains(unterrichtsId) && Schülergruppe == schuelergruppe && Bereinigen(Fach) == Bereinigen(fach))
        {
            if (!Klassen.Contains(klasse))
            {
                Klassen.Add(klasse);
                // Die SuS der weiteren Klasse werden dem Kurs hinzugefügt
                Students.AddRange(m.IStudents.Filter(m, zweck, klasse, schuelergruppe, studentgroupStudents));
            }

            // Die Anzahl der Wochenstunden wird nicht erhöht.
            return;
        }

        // Wenn UnterrichtsId nicht in der Liste der UnterrichtsIds des Kurses enthalten ist, wird sie hinzugefügt
        if (!UnterrichtsIds.Contains(int.Parse(unterrichtsId)))
        {
            UnterrichtsIds.Add(int.Parse(unterrichtsId));
            // Eine weitere UterrichtsId verändert die Kursbezeichnung                            
            KursBez = $"{Kursleiter}-{string.Join('-', UnterrichtsIds)}";
            // Die Wochenstunden des Kurses erhöhen sich nur, wenn die UnterrichtsId neu ist
            Wochenstunden += wochentundenLehrkraft;
        }

        // Wenn lehrer == Kursleiter, dann wird die Wochenstundenzahl des Kursleiters aktualisiert
        if (lehrer == Kursleiter)
        {
            KursleiterWochenstunden += wochentundenLehrkraft;
        }
        else
        {
            // Wenn lehrer != Kursleiter, dann wird der Lehrer als Zusatzlehrer hinzugefügt
            if (!Lehrkraefte.Contains(lehrer))
            {
                Lehrkraefte.Add(lehrer);
                LehrkraefteWochenstunden.Add(wochentundenLehrkraft);
            }
            else
            {
                // Wenn der Lehrer bereits als Zusatzlehrer existiert, wird die Wochenstundenzahl aktualisiert
                int index = Lehrkraefte.IndexOf(lehrer);
                if (index >= 0)
                {
                    LehrkraefteWochenstunden[index] += wochentundenLehrkraft;
                }
            }
        }
    }
}