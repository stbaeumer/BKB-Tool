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

    public class Relationsgruppen : List<Relationsgruppe>
    {
        public Relationsgruppen(Klassen klassen, Students students)
        {
            Console.Write("Relationsgruppe ..");
            // Dokumentation siehe schips.nrw.de/

            // Relationen gemäß §93 SchulG

            this.Add(new Relationsgruppe("BK BS Fachklasse EQ TZ", ["A01"],
                ["01", "02", "03"], Enumerable.Range(10000, 79999 - 10000).ToList(), 41.64));
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
            Console.WriteLine(".".PadRight(30, '.') + " " + (this.Count - 1));

            klassen.Relationsgruppen(this);
            klassen.KlasseOhneRelationsgruppe(students);
            Console.WriteLine("");

            List<string?> datei =
            [
                "^  Relationsgruppe  ^  ".PadRight(43) + "1.Jg  ^".PadRight(6) + "2.Jg  ^".PadRight(6) +
                "3.Jg  ^".PadRight(6) + "4.Jg  ^".PadRight(5) + "Summe  ^  Relation  ^  StellenBS  ^  StellenVZ  ^"
            ];
            Console.WriteLine("Relationsgruppe".PadRight(43) + "1.Jg".PadRight(6) + "2.Jg".PadRight(6) +
                              "3.Jg".PadRight(6) + "4.Jg".PadRight(6) + "Summe Relation StellenBS StellenVZ");

            int summe = 0;
            double stellenBs = 0;
            double stellenVz = 0;
            List<Student> schuelerSchnellmeldung = [];

            foreach (var relationsgruppe in this)
            {
                string? zeile = "|" + (relationsgruppe.BeschreibungSchulministerium + ":").PadRight(41) + "  |  ";

                var kl = (from k in klassen
                    where k.Relationsgruppe == relationsgruppe.BeschreibungSchulministerium
                    select k).ToList();

                foreach (var jg in new List<string?>() { "01", "02", "03", "04" })
                {
                    string z = "";
                    if (relationsgruppe.Jahrgänge.Contains(jg))
                    {
                        int x = (from s in students
                            where (from k in kl
                                where k.Name == s.Klasse
                                where s.Jahrgang == jg
                                select k).Any()
                            select s).Count();

                        schuelerSchnellmeldung.AddRange((from s in students
                            where (from k in kl where k.Name == s.Klasse where s.Jahrgang == jg select k).Any()
                            select s));
                        summe += x;
                        z = x.ToString();
                    }

                    zeile = zeile + z.PadLeft(6) + "|";
                }

                int t = (from s in students
                    where (from k in kl where k.Name == s.Klasse select k).Any()
                    select s).Count();

                zeile = zeile + t.ToString().PadLeft(6) + "|";

                // Relation:

                zeile = zeile + relationsgruppe.Relation.ToString().PadLeft(9) + "|";

                // Stellen:

                if (relationsgruppe.BeschreibungSchulministerium.StartsWith("BK BS"))
                {
                    zeile = zeile + (t / relationsgruppe.Relation).ToString("0.0000").PadLeft(10) + "|             |";
                    stellenBs = stellenBs + (t / relationsgruppe.Relation);
                }
                else
                {
                    zeile = zeile + "          |" + (t / relationsgruppe.Relation).ToString("0.0000").PadLeft(10) + "|";
                    stellenVz = stellenVz + (t / relationsgruppe.Relation);
                }

                datei.Add(zeile);
                Console.WriteLine(zeile);
            }

            datei.Add("|Summe:".PadRight(67) + summe + "|||||||    " + stellenBs.ToString("0.0000").PadLeft(10) + "|" +
                      stellenVz.ToString("0.0000").PadLeft(10) + "|");
            datei.Add("");
            datei.Add("   Leitungszeit: " + (9 + 50 * 0.7 + ((stellenBs + stellenVz) - 50) * 0.3).ToString("0.00") +
                      "   (= 9 + 50 * 0,7 + (" + stellenBs.ToString("0.0000") + "+" + stellenVz.ToString("0.0000") +
                      " - 50) * 0,3)  Verordnung zur Ausführung des § 93 Abs. 2 Schulgesetz (VO zu § 93 Abs. 2 SchulG) vom 18.03.2005");
            datei.Add("   Anrechnungen: " + (stellenBs * 0.5 + stellenVz * 1.2).ToString("00.00") + "   (= " +
                      stellenBs.ToString("0.0000") + " * 0,5 + " + stellenVz.ToString("0.0000") + " * 1,2)");
            datei.Add("");
            Console.WriteLine("Summe: " + summe);

            foreach (var student in students)
            {
                if (!(from s in schuelerSchnellmeldung where s.IdSchild == student.IdSchild select s).Any())
                {
                    datei.Add("Der Schüler " + student.Nachname + ", " + student.Vorname + " " + student.Klasse +
                              " ist nicht in der Schnellmeldung erfasst. Prüfen!");
                }
            }

            datei.Add("   Schüler in SchILD insgesamt: " + students.Count());
            Console.WriteLine("Schüler in SchILD insgesamt: " + students.Count());
            datei.Add("");

            datei.Add("Details");
            datei.Add("=======");
            datei.Add("");

            foreach (var relationsgruppe in this)
            {
                int sum = 0;
                datei.Add(relationsgruppe.BeschreibungSchulministerium);
                int i = 1;
                foreach (var klasse in (from k in klassen
                             where k.Relationsgruppe == relationsgruppe.BeschreibungSchulministerium
                             select k).ToList())
                {
                    int d = (from s in students where s.Klasse == klasse.Name select s).Count();
                    sum += d;
                    datei.Add(i.ToString().PadLeft(2) + ". " + klasse.Name.PadRight(20) + d.ToString().PadLeft(2));
                    i++;
                }

                datei.Add("--------------------------");
                datei.Add("".PadRight(15) + "Summe: " + sum.ToString().PadLeft(4));
                datei.Add("");
            }

            datei.Add(Environment.UserName + " | " + DateTime.Now);

            string pfad = DateTime.Now.ToString("yyyyMMdd") + ".txt";

            using (StreamWriter outputFile = new StreamWriter(pfad))
            {
                foreach (string? line in datei)
                    outputFile.WriteLine(line);
            }

            Global.EditorOeffnen(pfad);

            System.Diagnostics.Process.Start("https://www.schips.nrw.de/");
        }
    }