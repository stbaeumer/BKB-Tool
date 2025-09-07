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

using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

public class Relationsgruppen : List<Relationsgruppe>
{
    private IConfiguration configuration;
    private Students students;

    public Relationsgruppen(Klassen klassen, Students students, IConfiguration configuration)
    {
        // Dokumentation siehe schips.nrw.de/
        // Relationen gemäß §93 SchulG

        this.Add(new Relationsgruppe("BK BS Fachklasse EQ TZ",
            ["A01"], ["01", "02", "03"], Enumerable.Range(10000, 79999 - 10000).ToList(), 41.64));
        this.Add(new Relationsgruppe("BK BS Fachklasse EQ TZ (hj. endend)", ["A01"],
            ["04"], null, 83.28));
        this.Add(new Relationsgruppe("BK BS Ausbildungsvorbereitung VZ", ["A12"],
            ["01"], null, 16.18));
        this.Add(new Relationsgruppe("BK BS Ausbildungsvorbereitung TZ", ["A13"],
            ["01"], null, 41.64));
        this.Add(new Relationsgruppe("BK BF ber. Kennt. (Vor. HSA) 1-jährig", ["B06"],
            ["01"], null, 16.18));
        this.Add(new Relationsgruppe("BK BF ber. Kennt. (Vor. HSA 10) 1-jährig",
            ["B02", "B07"], ["01"], null, 16.18));
        this.Add(new Relationsgruppe("BK BF ber. Kennt. und FHS 2-jährig", ["C03", "C13"],
            ["01", "02"], null, 16.18));
        this.Add(new Relationsgruppe("BK BF BA-LRecht und FOS 2-jährig", ["B01", "B04", "B08"],
            ["01", "02"], null, 14.34));
        this.Add(new Relationsgruppe("BK BF ber. Kennt. und AHR 3-jährig", ["D02"],
            ["01", "02", "03"], null, 14.34));
        this.Add(new Relationsgruppe("BK FO 12B 1-jährig VZ", ["C08"],
            ["01"], null, 14.34));
        this.Add(new Relationsgruppe("BK FO Klasse 11 1-jährig TZ", ["C05"],
            ["01"], null, 41.64));
        this.Add(new Relationsgruppe("BK FO Klasse 12 1-jährig VZ", ["C06"],
            ["01"], null, 14.34));

        Students studentsFiltered = new Students();

        studentsFiltered.AddRange(from s in students where s.Status == "2" || s.Status == "6" select s);

        var table = new Table();
        table.Expand();
        table.Border(TableBorder.Rounded);
        table.Title = new TableTitle($"Fehler in Relationsgruppe");
        table.Expand();
        table.AddColumn("Name");
        table.AddColumn("Gl.");
        table.AddColumn("Jahrgang");
        table.AddColumn("Fachklasse");
        table.AddColumn("Fehler");

        table = studentsFiltered.GetRelationsgruppe(this, table);

        if (table.Rows.Count > 0)
        {
            AnsiConsole.Write(table);
        }

        List<string?> datei =
        [
            "^  Relationsgruppe  ^  ".PadRight(43) + "1.Jg  ^".PadRight(6) + "2.Jg  ^".PadRight(6) +
            "3.Jg  ^".PadRight(6) + "4.Jg  ^".PadRight(5) + "Summe  ^  Relation  ^  StellenBS  ^  StellenVZ  ^"
        ];

        var table1 = new Table();
        table1.Expand();
        table1.Border(TableBorder.Rounded);
        table1.Title = new TableTitle($"Relationsgruppen");
        table1.Expand();
        table1.AddColumn("Relationsgruppe");
        table1.AddColumn("1.Jg");
        table1.AddColumn("2.Jg");
        table1.AddColumn("3.Jg");
        table1.AddColumn("4.Jg");
        table1.AddColumn("Summe");
        table1.AddColumn("Relation");
        table1.AddColumn("StellenBS");
        table1.AddColumn("StellenVZ");
        // alle Spaten rechtsbündig
        table1.Columns[0].Alignment = Justify.Left;
        table1.Columns[1].Alignment = Justify.Right;
        table1.Columns[2].Alignment = Justify.Right;
        table1.Columns[3].Alignment = Justify.Right;
        table1.Columns[4].Alignment = Justify.Right;
        table1.Columns[5].Alignment = Justify.Right;
        table1.Columns[6].Alignment = Justify.Right;
        table1.Columns[7].Alignment = Justify.Right;
        table1.Columns[8].Alignment = Justify.Right;

        int summe = 0;
        double stellenBs = 0;
        double stellenVz = 0;
        List<Student> schuelerSchnellmeldung = [];

        foreach (var relationsgruppe in this)
        {
            List<string> zeile = new List<string> { "|" + relationsgruppe.BeschreibungSchulministerium + ":" };

            foreach (var jg in new List<string?>() { "01", "02", "03", "04" })
            {
                string z = "";
                if (relationsgruppe.Jahrgänge.Contains(jg))
                {
                    int x = (from s in studentsFiltered
                             where s.Jahrgang.EndsWith(jg)
                             where s.Relationsgruppe == relationsgruppe.BeschreibungSchulministerium
                             select s).Count();

                    schuelerSchnellmeldung.AddRange((from s in studentsFiltered where s.Jahrgang.EndsWith(jg) select s));
                    summe += x;
                    z = x.ToString();
                }

                zeile.Add(z);
            }

            int t = (from s in studentsFiltered
                     where s.Relationsgruppe == relationsgruppe.BeschreibungSchulministerium
                     select s).Count();

            zeile.Add($"[{Global.GetColor(Global.ColorHinweise)}]{t.ToString()}[/]");

            // Relation:

            zeile.Add(relationsgruppe.Relation.ToString());

            // Stellen:

            if (relationsgruppe.BeschreibungSchulministerium.Contains("TZ"))
            {
                zeile.Add((t / relationsgruppe.Relation).ToString("0.0000").PadLeft(10));
                zeile.Add("");
                stellenBs = stellenBs + (t / relationsgruppe.Relation);
            }
            else
            {
                zeile.Add("");
                zeile.Add((t / relationsgruppe.Relation).ToString("0.0000") + "|");
                stellenVz = stellenVz + (t / relationsgruppe.Relation);
            }

            datei.Add(string.Join("|", zeile));
            zeile = zeile.Select(z => z.Replace("|", "")).ToList();
            table1.AddRow(zeile.ToArray());
        }

        table1.AddRow([
            "Summen:",
            studentsFiltered.Where(s=>s.Jahrgang.EndsWith("01")).Count().ToString(),
            studentsFiltered.Where(s=>s.Jahrgang.EndsWith("02")).Count().ToString(),
            studentsFiltered.Where(s=>s.Jahrgang.EndsWith("03")).Count().ToString(),
            studentsFiltered.Where(s=>s.Jahrgang.EndsWith("04")).Count().ToString(),
            $"[{Global.GetColor(Global.ColorHinweise)}]{summe.ToString()}[/]",
            studentsFiltered.Count().ToString(),
            stellenBs.ToString("0.0000"),
            stellenVz.ToString("0.0000")
        ]);

        AnsiConsole.Write(table1);

        datei.Add("|Summe:".PadRight(67) + summe + "|||||||    " + stellenBs.ToString("0.0000").PadLeft(10) + "|" +
                    stellenVz.ToString("0.0000").PadLeft(10) + "|");
        datei.Add("");

        var panel = new Panel(
            new Markup((9 + 50 * 0.7 + ((stellenBs + stellenVz) - 50) * 0.3).ToString("0.00") +
                    "   (= 9 + 50 * 0,7 + (" + stellenBs.ToString("0.0000") + "+" + stellenVz.ToString("0.0000") +
                    " - 50) * 0,3) Verordnung zur Ausführung des § 93 Abs. 2 Schulgesetz (VO zu § 93 Abs. 2 SchulG) vom 18.03.2005"))
        {
            Header = new PanelHeader("  Leitungszeit  ", Justify.Left),
            Border = BoxBorder.Rounded,
            //Padding = new Padding(1, 1),
            Expand = true
        };
        AnsiConsole.Write(panel);

        datei.Add("Leitungszeit: " + (9 + 50 * 0.7 + ((stellenBs + stellenVz) - 50) * 0.3).ToString("0.00") +
                    "   (= 9 + 50 * 0,7 + (" + stellenBs.ToString("0.0000") + "+" + stellenVz.ToString("0.0000") +
                    " - 50) * 0,3)  Verordnung zur Ausführung des § 93 Abs. 2 Schulgesetz (VO zu § 93 Abs. 2 SchulG) vom 18.03.2005");

        panel = new Panel(
            new Markup((stellenBs * 0.5 + stellenVz * 1.2).ToString("00.00") + "   (= " +
                    stellenBs.ToString("0.0000") + " * 0,5 + " + stellenVz.ToString("0.0000") + " * 1,2)"))
        {
            Header = new PanelHeader("  Anrechnungen  ", Justify.Left),
            Border = BoxBorder.Rounded,
            //Padding = new Padding(1, 1),
            Expand = true
        };
        AnsiConsole.Write(panel);
        datei.Add("");
        datei.Add("Anrechnungen: " + (stellenBs * 0.5 + stellenVz * 1.2).ToString("00.00") + "   (= " +
                    stellenBs.ToString("0.0000") + " * 0,5 + " + stellenVz.ToString("0.0000") + " * 1,2)");
        datei.Add("");

        foreach (var student in studentsFiltered)
        {
            if (!(from s in schuelerSchnellmeldung where s.Id == student.Id select s).Any())
            {
                datei.Add("Der Schüler " + student.Nachname + ", " + student.Vorname + " " + student.Klasse +
                            " ist nicht in der Schnellmeldung erfasst. Prüfen!");
            }
        }

        datei.Add("   Schüler in SchILD insgesamt: " + studentsFiltered.Count());
        datei.Add("");

        datei.Add("Details");
        datei.Add("=======");
        datei.Add("");


        table1 = new Table();
        table1.Expand();
        table1.Border(TableBorder.Rounded);
        table1.Title = new TableTitle($"Relationsgruppen");
        table1.Expand();

        foreach (var relationsgruppe in this)
        {
            table1.AddColumn(relationsgruppe.BeschreibungSchulministerium);
        }

        int summeUeberAlle = 0;
        foreach (var relationsgruppe in this)
        {
            int sum = 0;
            datei.Add(relationsgruppe.BeschreibungSchulministerium);
            //Console.WriteLine(relationsgruppe.BeschreibungSchulministerium + ": " + studentsFiltered.Where(s => s.Relationsgruppe == relationsgruppe.BeschreibungSchulministerium).Count() + ":");
            int i = 1;
            foreach (var klasse in (from k in studentsFiltered
                                    where k.Relationsgruppe == relationsgruppe.BeschreibungSchulministerium
                                    select k.Klasse).Distinct().OrderBy(s => s).ToList())
            {
                int d = (from s in studentsFiltered
                         where s.Klasse == klasse
                         where s.Relationsgruppe == relationsgruppe.BeschreibungSchulministerium
                         select s).Count();
                sum += d;
                summeUeberAlle += d;
                datei.Add(i.ToString().PadLeft(2) + ". " + klasse.PadRight(12) + d.ToString().PadLeft(2));
                //Console.WriteLine(" " + i.ToString().PadLeft(2) + " " + klasse.PadRight(18) + ": " + studentsFiltered.Where(s => (s.Status == "2" || s.Status == "6") && s.Klasse == klasse).Count().ToString().PadLeft(4));
                i++;
            }

            datei.Add("-".PadRight(56, '-'));
            datei.Add("".PadRight(45) + "Summe: " + sum.ToString().PadLeft(4));
            datei.Add("");
        }

        //datei.Add("Gesamtsumme: " + summeUeberAlle.ToString().PadLeft(4));
        //datei.Add("");

        datei.Add("-".PadRight(56, '-'));
        datei.Add("".PadRight(39) + "Gesamtsumme: " + summeUeberAlle.ToString().PadLeft(4));

        datei.Add(Environment.UserName + " | " + DateTime.Now);

        string pfad = Path.Combine(configuration["PfadDownloads"], DateTime.Now.ToString("yyyyMMdd") + ".txt");

        using (StreamWriter outputFile = new StreamWriter(pfad))
        {
            foreach (string? line in datei)
                outputFile.WriteLine(line);
        }

        Global.EditorOeffnen(pfad);

        // Öffne Browser: https://www.schulministerium.nrw.de/BiPo/ppschips/pages/index.jsf
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = "https://www.schulministerium.nrw.de/BiPo/ppschips/pages/index.jsf",
            UseShellExecute = true
        });
    }
}