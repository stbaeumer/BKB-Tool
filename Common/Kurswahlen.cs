internal class Kurswahlen : List<Kurswahl>
{
    /*
    private readonly string datei;
    private readonly List<Student> schuelers;
    private readonly Unterrichts unterrichts;
    private readonly Lehrers lehrers;
    private readonly Klasses klasses;
    private readonly Unterrichts untisUnterrichts;
    private readonly Gruppen untisGruppen;
    private readonly string hzJz;

    public Kurswahlen(
        string datei,
        List<Student> schuelers,
        Unterrichts unterrichts,
        Lehrers lehrers,
        Klasses klasses,
        Unterrichts untisUnterrichts,
        Gruppen untisGruppen,
        string hzJz
        )
    {
        Global.OrdnerAnlegen(datei);
                    
        File.WriteAllText(Global.TempPfad + datei, "====== Klausurbelegungspläne ======" + Environment.NewLine);

        File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

        File.AppendAllText(Global.TempPfad + datei, "  Bitte diese Seite nicht manuell ändern." + Environment.NewLine);

        File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

        File.AppendAllText(Global.TempPfad + datei, "Klausurbelegungspläne des[[berufliches_gymnasium:start | Beruflichen Gymnasiums]]. Siehe auch:" + Environment.NewLine);

        File.AppendAllText(Global.TempPfad + datei, "  * [[oeffentlich:klausurplanung|Klausurplanung]]" + Environment.NewLine);
        File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

        File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

        // 11er

        var verschiedene11erKlassen = (from s in schuelers
                                        where s.Klasse.NameUntis.StartsWith("G")
                                        where s.Klasse.NameUntis.Contains(Global.AktSj[0].Substring(2,2))
                                        select s.Klasse).Distinct().ToList();

        Global.WriteLine("Kurswahlen in den Klassen der Jahrgangsstufe 11", verschiedene11erKlassen.Count);

        File.AppendAllText(Global.TempPfad + datei, "===== Jahrgang 11 (Belegung aus Webuntis) =====" + Environment.NewLine);
        File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

        foreach (var klasse in verschiedene11erKlassen)
        {
            File.AppendAllText(Global.TempPfad + datei, "====" + klasse.NameUntis + "====" + Environment.NewLine);
            File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

            Schuelers sus = new Schuelers();
            sus.AddRange(from s in schuelers where s.Klasse == klasse select s);
            sus.GetWebuntisUnterrichte(untisUnterrichts, untisGruppen, klasse.NameUntis, hzJz);

            var verschiedene11erFächer = (from s in sus
                                            from u in s.UnterrichteAusWebuntis
                                            select new { u.Fach, u.Lehrkraft }).Distinct().ToList();

            var kopfzeile1 = "^  Nr.  ^  Name  ^  ";
            var kopfzeile2 = "^  :::  ^  :::   ^  ";

            foreach (var fach in verschiedene11erFächer)
            {
                kopfzeile1 += fach.Fach + "  ^  ";
                kopfzeile2 += fach.Lehrkraft + "  ^  ";
            }

            File.AppendAllText(Global.TempPfad + datei, "Jahrgang: 11  |  " + (hzJz == "HZ" ? "1." : "2.") + " Halbjahr  |  Fehler gefunden? ((Fehler müssen in Webuntis korrigiert werden. Anschließend wird die Tabelle automatisch täglich aktualisiert.))" + Environment.NewLine);
            File.AppendAllText(Global.TempPfad + datei, kopfzeile1 + Environment.NewLine);
            File.AppendAllText(Global.TempPfad + datei, kopfzeile2 + Environment.NewLine);

            var z = 1;

            Global.WriteLine(" SuS mit Kurswahl in " + klasse.NameUntis, sus.Count);

            foreach (var s in sus)
            {
                var zeileSuS = "|  " + z + ".|" + s.Nachname + " " + s.Vorname + "  |  ";
                z++;
                foreach (var fach in verschiedene11erFächer)
                {
                    if ((from ss in s.UnterrichteAusWebuntis where ss.Fach == fach.Fach select ss).Any())
                    {
                        zeileSuS += "  X  |";
                    }
                    else
                    {
                        zeileSuS += "     | ";
                    }
                }
                File.AppendAllText(Global.TempPfad + datei, zeileSuS + Environment.NewLine);
            }
            File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);
        }


        var abfrage = "";

        foreach (var schuelerId in (from s in schuelers select s.Id))
        {
            abfrage += "(DBA.schueler.pu_id = " + schuelerId + @") OR ";
        }

        try
        {
            abfrage = abfrage.Substring(0, abfrage.Length - 4);
        }
        catch (Exception ex)
        {

        }


        try
        {
            using (OdbcConnection connection = new OdbcConnection(Global.ConnectionStringAtlantis))
            {
                DataSet dataSet = new DataSet();

                OdbcDataAdapter schuelerAdapter = new OdbcDataAdapter(@"
SELECT DBA.noten_einzel.noe_id AS LeistungId,
DBA.noten_einzel.fa_id,
DBA.noten_einzel.kurztext AS Fach,
DBA.noten_einzel.zeugnistext AS Zeugnistext,
DBA.noten_einzel.s_note AS Note,
DBA.noten_einzel.punkte AS Punkte,
DBA.noten_einzel.punkte_12_1 AS Punkte_12_1,
DBA.noten_einzel.punkte_12_2 AS Punkte_12_2,
DBA.noten_einzel.punkte_13_1 AS Punkte_13_1,
DBA.noten_einzel.punkte_13_2 AS Punkte_13_2,
DBA.noten_einzel.s_eingebracht_12_1 AS Eingebracht_12_1,
DBA.noten_einzel.s_eingebracht_12_2 AS Eingebracht_12_2,
DBA.noten_einzel.s_eingebracht_13_1 AS Eingebracht_13_1,
DBA.noten_einzel.s_eingebracht_13_2 AS Eingebracht_13_2,
DBA.noten_einzel.s_abiturfach AS Abiturfach,
DBA.noten_einzel.s_tendenz AS Tendenz,
DBA.noten_einzel.s_einheit AS Einheit,
DBA.noten_einzel.ls_id_1 AS LehrkraftAtlantisId,
DBA.noten_einzel.position_1 AS Reihenfolge,
DBA.schueler.name_1 AS Nachname,
DBA.schueler.name_2 AS Vorname,
DBA.schueler.dat_geburt,
DBA.schueler.pu_id AS SchlüsselExtern,
DBA.schue_sj.s_religions_unterricht AS Religion,
DBA.schue_sj.dat_austritt AS ausgetreten,
DBA.schue_sj.dat_rel_abmeld AS DatumReligionAbmeldung,
DBA.schue_sj.vorgang_akt_satz_jn AS SchuelerAktivInDieserKlasse,
DBA.schue_sj.vorgang_schuljahr AS Schuljahr,
(substr(schue_sj.s_berufs_nr,4,5)) AS Fachklasse,
DBA.klasse.s_klasse_art AS Anlage,
DBA.klasse.jahrgang AS Jahrgang,
DBA.schue_sj.s_gliederungsplan_kl AS Gliederung,
DBA.noten_kopf.s_typ_nok AS HzJz,
DBA.noten_kopf.unterzeichner_rechts AS Klassenleiter,
DBA.noten_kopf.unterzeichner_rechts_text AS Klassenleitername,
DBA.noten_kopf.nok_id AS NOK_ID,
s_art_fach,
DBA.noten_kopf.s_art_nok AS Zeugnisart,
DBA.noten_kopf.bemerkung_block_1 AS Bemerkung1,
DBA.noten_kopf.bemerkung_block_2 AS Bemerkung2,
DBA.noten_kopf.bemerkung_block_3 AS Bemerkung3,
DBA.noten_kopf.dat_notenkonferenz AS Konferenzdatum,
DBA.klasse.klasse AS Klasse
FROM(((DBA.noten_kopf JOIN DBA.schue_sj ON DBA.noten_kopf.pj_id = DBA.schue_sj.pj_id) JOIN DBA.klasse ON DBA.schue_sj.kl_id = DBA.klasse.kl_id) JOIN DBA.noten_einzel ON DBA.noten_kopf.nok_id = DBA.noten_einzel.nok_id ) JOIN DBA.schueler ON DBA.noten_einzel.pu_id = DBA.schueler.pu_id
WHERE schue_sj.s_typ_vorgang = 'A' AND (s_typ_nok = 'JZ' OR s_typ_nok = 'HZ' OR s_typ_nok = 'GO') AND
(  
" + abfrage + @"
)
ORDER BY DBA.klasse.s_klasse_art DESC, DBA.noten_kopf.dat_notenkonferenz DESC, DBA.klasse.klasse ASC, DBA.noten_kopf.nok_id, DBA.noten_einzel.position_1; ", connection);

                connection.Open();
                schuelerAdapter.Fill(dataSet, "DBA.leistungsdaten");

                string bereich = "";

                foreach (DataRow theRow in dataSet.Tables["DBA.leistungsdaten"].Rows)
                {
                    if (theRow["s_art_fach"].ToString() == "U")
                    {
                        bereich = theRow["Zeugnistext"].ToString();
                    }
                    else
                    {
                        DateTime austrittsdatum = theRow["ausgetreten"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["ausgetreten"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);

                        Kurswahl kurswahl = new Kurswahl();

                        try
                        {
                            // Wenn der Schüler nicht in diesem Schuljahr ausgetreten ist ...

                            if (!(austrittsdatum > new DateTime(DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.Year - 1, 8, 1) && austrittsdatum < DateTime.Now))
                            {
                                kurswahl.LeistungId = Convert.ToInt32(theRow["LeistungId"]);
                                kurswahl.NokId = Convert.ToInt32(theRow["NOK_ID"]);
                                kurswahl.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"]);
                                kurswahl.Schuljahr = theRow["Schuljahr"].ToString();
                                kurswahl.Gliederung = theRow["Gliederung"].ToString();
                                kurswahl.Abifach = theRow["Abiturfach"].ToString();
                                kurswahl.HatBemerkung = (theRow["Bemerkung1"].ToString() + theRow["Bemerkung2"].ToString() + theRow["Bemerkung3"].ToString()).Contains("Fehlzeiten") ? true : false;
                                kurswahl.Jahrgang = Convert.ToInt32(theRow["Jahrgang"].ToString().Substring(3, 1));
                                kurswahl.Name = theRow["Nachname"] + " " + theRow["Vorname"];
                                kurswahl.Nachname = theRow["Nachname"].ToString();
                                kurswahl.Vorname = theRow["Vorname"].ToString();

                                kurswahl.KlassenleiterName = theRow["Klassenleitername"].ToString();
                                kurswahl.Klassenleiter = theRow["Klassenleiter"].ToString();

                                if ((theRow["LehrkraftAtlantisId"]).ToString() != "")
                                {
                                    kurswahl.LehrkraftAtlantisId = Convert.ToInt32(theRow["LehrkraftAtlantisId"]);
                                }
                                kurswahl.Bereich = bereich;
                                try
                                {
                                    kurswahl.Reihenfolge = Convert.ToInt32(theRow["Reihenfolge"].ToString());
                                }
                                catch (Exception)
                                {
                                    throw;
                                }

                                kurswahl.Geburtsdatum = theRow["dat_geburt"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["dat_geburt"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                kurswahl.Volljährig = kurswahl.Geburtsdatum.AddYears(18) > DateTime.Now ? false : true;
                                kurswahl.Klasse = theRow["Klasse"].ToString();
                                kurswahl.Fach = theRow["Fach"] == null ? "" : theRow["Fach"].ToString();
                                kurswahl.Gesamtnote = theRow["Note"].ToString() == "" ? null : theRow["Note"].ToString() == "Attest" ? "A" : theRow["Note"].ToString();
                                kurswahl.Gesamtpunkte_12_1 = theRow["Punkte_12_1"].ToString() == "" ? null : (theRow["Punkte_12_1"].ToString()).Split(',')[0];
                                kurswahl.Gesamtpunkte_12_2 = theRow["Punkte_12_2"].ToString() == "" ? null : (theRow["Punkte_12_2"].ToString()).Split(',')[0];
                                kurswahl.Gesamtpunkte_13_1 = theRow["Punkte_13_1"].ToString() == "" ? null : (theRow["Punkte_13_1"].ToString()).Split(',')[0];
                                kurswahl.Gesamtpunkte_13_2 = theRow["Punkte_13_2"].ToString() == "" ? null : (theRow["Punkte_13_2"].ToString()).Split(',')[0];
                                kurswahl.Gesamtpunkte = theRow["Punkte"].ToString() == "" ? null : (theRow["Punkte"].ToString()).Split(',')[0];
                                kurswahl.Eingebracht_12_1 = theRow["Eingebracht_12_1"].ToString() == "" ? null : theRow["Eingebracht_12_1"].ToString();
                                kurswahl.Eingebracht_12_2 = theRow["Eingebracht_12_2"].ToString() == "" ? null : theRow["Eingebracht_12_2"].ToString();
                                kurswahl.Eingebracht_13_1 = theRow["Eingebracht_13_1"].ToString() == "" ? null : theRow["Eingebracht_13_1"].ToString();
                                kurswahl.Eingebracht_13_2 = theRow["Eingebracht_13_2"].ToString() == "" ? null : theRow["Eingebracht_13_2"].ToString();
                                kurswahl.Tendenz = theRow["Tendenz"].ToString() == "" ? null : theRow["Tendenz"].ToString();
                                kurswahl.EinheitNP = theRow["Einheit"].ToString() == "" ? "N" : theRow["Einheit"].ToString();
                                kurswahl.SchlüsselExtern = Convert.ToInt32(theRow["SchlüsselExtern"].ToString());
                                kurswahl.HzJz = theRow["HzJz"].ToString();
                                kurswahl.Anlage = theRow["Anlage"].ToString();
                                kurswahl.Zeugnisart = theRow["Zeugnisart"].ToString();
                                kurswahl.Zeugnistext = theRow["Zeugnistext"].ToString();
                                kurswahl.Konferenzdatum = theRow["Konferenzdatum"].ToString().Length < 3 ? new DateTime() : (DateTime.ParseExact(theRow["Konferenzdatum"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)).AddHours(15);
                                kurswahl.DatumReligionAbmeldung = theRow["DatumReligionAbmeldung"].ToString().Length < 3 ? new DateTime() : DateTime.ParseExact(theRow["DatumReligionAbmeldung"].ToString(), "dd.MM.yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                                kurswahl.SchuelerAktivInDieserKlasse = theRow["SchuelerAktivInDieserKlasse"].ToString() == "J";
                                kurswahl.Beschreibung = "";

                                if (kurswahl.HzJz == "GO" && kurswahl.Klasse.StartsWith("G") && !kurswahl.Klasse.Contains(Global.AktSj[0].Substring(2,2)))
                                {
                                    kurswahl.Lehrkraft = (from u in unterrichts
                                                            where u.KlasseKürzel == kurswahl.Klasse
                                                            where u.FachKürzel == kurswahl.Fach.Replace("  ", " ")
                                                            select u.LehrerKürzel).FirstOrDefault();

                                    if (kurswahl.Lehrkraft == null)
                                    {
                                        var fach = (from u in unterrichts
                                                    where u.KlasseKürzel == kurswahl.Klasse
                                                    where u.FachKürzel == kurswahl.Fach.Replace("  ", " ")
                                                    select u.FachKürzel).FirstOrDefault();
                                        kurswahl.Lehrkraft = "";

                                        Console.WriteLine(kurswahl.Klasse + ": Atlantis: " + kurswahl.Fach + " <-> Untis: ---");
                                    }

                                    this.Add(kurswahl);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            string a = "";
                        }
                    }
                }
                connection.Close();
            }
        }
        catch (Exception ex)
        {
            throw ex;
        }

        // 12er und 13er
        

        File.AppendAllText(Global.TempPfad + datei, "===== Jahrgang 12 & 13 (Belegung aus Atlantis) =====" + Environment.NewLine);
        File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

        var verschiedeneKlassen = (from t in this.OrderBy(x => x.Jahrgang).ThenBy(x => x.Klasse)
                                    where t.Konferenzdatum > new DateTime(Convert.ToInt32(Global.AktSj[0]), 10, 01)
                                    where t.Konferenzdatum < new DateTime(Convert.ToInt32(Global.AktSj[1]), 08, 01)
                                    select t.Klasse).Distinct().ToList();

        Global.WriteLine("Kurswahlen in den Klassen der Jahrgangsstufe 12 & 13", verschiedeneKlassen.Count);

        var bbereiche = (from k in klasses where verschiedeneKlassen.Contains(k.NameUntis) select new { V = k.NameUntis.Substring(0, 2), k.BildungsgangLangname, k.WikiLink }).Distinct().ToList();

        foreach (var bereich in bbereiche)
        {
            foreach (var klasse in verschiedeneKlassen.Where(x => x.StartsWith(bereich.V)))
            {
                File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);
                File.AppendAllText(Global.TempPfad + datei, "====" + klasse + "====" + Environment.NewLine);
                File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

                var jahrgang = (from t in this.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname) where t.Klasse == klasse select t.Jahrgang).FirstOrDefault();

                var gliederung = (from t in this.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname) where t.Klasse == klasse select t.Gliederung).FirstOrDefault();

                var konferenzdatum = (from t in this.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname) where t.Klasse == klasse select t.Konferenzdatum.ToShortDateString()).FirstOrDefault();

                File.AppendAllText(Global.TempPfad + datei, "[[" + bereich.WikiLink + " | " + klasse.Substring(0, 2) + " ]]  |  Jahrgang:" + (10 + jahrgang) + "  |  " + (hzJz == "HZ" ? "1." : "2.") + " Halbjahr  |  Konferenzdatum: [[:konferenzen:zeugniskonferenzen|" + konferenzdatum + "]]  |  Gliederung: " + gliederung + "  |  Fehler gefunden? ((Fehler müssen in Atlantis korrigiert werden. Anschließend wird die Tabelle automatisch täglich aktualisiert.))" + Environment.NewLine);

                File.AppendAllText(Global.TempPfad + datei, Environment.NewLine);

                var schuelerDieserKlasse = (from t in this.OrderBy(x => x.Nachname).ThenBy(x => x.Vorname) where t.Klasse == klasse select t.SchlüsselExtern).Distinct().ToList();

                var bereiche = (from t in this.OrderBy(x => x.Reihenfolge) where t.Klasse == klasse select t.Bereich).Distinct().ToList();

                var klassenleitung = (from t in this.OrderBy(x => x.Reihenfolge) where t.Klasse == klasse select t.Klassenleiter + ", " + t.KlassenleiterName).Distinct().ToList();

                var alleFächer = (from t in this.OrderBy(x => x.Reihenfolge) where t.Klasse == klasse select new { t.Fach, t.Bereich, t.Lehrkraft }).Distinct().ToList();

                var kopfzeile1 = "^                                 ^^" + alleFächer[1].Bereich;
                var kopfzeile2 = "^:::                              ^^";
                var kopfzeile3 = "^:::                              ^^";

                for (int i = 0; i < alleFächer.Count; i++)
                {
                    if (i > 0 && alleFächer[i - 1].Bereich != alleFächer[i].Bereich)
                    {
                        kopfzeile1 += alleFächer[i].Bereich + "  ^";
                    }
                    else
                    {
                        kopfzeile1 += "^";
                    }

                    kopfzeile2 += alleFächer[i].Fach.PadRight(5) + "^";
                    kopfzeile3 += "  " + alleFächer[i].Lehrkraft.PadRight(5) + "^";
                }

                // Wahlklausur anhängen
                kopfzeile1 += "  Wahlklausuren  ^^^^";
                kopfzeile2 += " ^^^^";
                kopfzeile3 += "12.1^12.2^13.1^13.2^";

                File.AppendAllText(Global.TempPfad + datei, kopfzeile1 + Environment.NewLine);
                File.AppendAllText(Global.TempPfad + datei, kopfzeile2 + Environment.NewLine);
                File.AppendAllText(Global.TempPfad + datei, kopfzeile3 + Environment.NewLine);

                int y = 1;

                foreach (var id in schuelerDieserKlasse)
                {
                    var schueler = (from s in schuelers where s.Id == id select s).FirstOrDefault();

                    var wahlklausur = RemoveLineEndings(schueler.Wahlklausur12_1) + "  |  " + RemoveLineEndings(schueler.Wahlklausur12_2) + "  |  " + RemoveLineEndings(schueler.Wahlklausur13_1) + "  |  " + RemoveLineEndings(schueler.Wahlklausur13_2) + "  |";

                    var fächerDesSchülers = (from t in this where t.SchlüsselExtern == id select t).ToList();

                    var zeile = "|" + y.ToString().PadLeft(3) + ".|" + (from t in this where t.SchlüsselExtern == id select t.Nachname + ", " + t.Vorname).FirstOrDefault().PadRight(27) + "  |";

                    y++;

                    foreach (var fach in alleFächer)
                    {
                        var belegung = (from f in fächerDesSchülers
                                        where f.Fach == fach.Fach
                                        select f).FirstOrDefault();

                        if (belegung != null)
                        {
                            if (jahrgang == 1)
                            {
                                zeile += "  X  |";

                            }
                            else if (belegung.IstAbifach1bis3()) // Abifächer (außer 4.) sind immer P
                            {
                                zeile += "  P**(" + belegung.Abifach + ")**  |";
                            }
                            else
                            {
                                // 12

                                if (jahrgang == 2)
                                {
                                    if (IstFremdsprache(fach.Fach)) // bis einschl 13.1. immer P
                                    {
                                        zeile += "  P  |";
                                    }
                                    else if (IstDeutschOderMathe(fach.Fach)) // bis einschl 12.2. immer P
                                    {
                                        zeile += "  P" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                    else if (IstWahlklausur(fach.Fach, schueler.Wahlklausur12_1)) // W
                                    {
                                        zeile += "  W" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                    else
                                    {
                                        zeile += "  X" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                }

                                // 13.1

                                if (jahrgang == 3 && hzJz == "HZ")
                                {
                                    if (IstFremdsprache(fach.Fach)) // bis einschl 13.1. immer P
                                    {
                                        zeile += "  P" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                    else if (IstWahlklausur(fach.Fach, schueler.Wahlklausur13_1)) // W
                                    {
                                        zeile += "  W" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                    else
                                    {
                                        zeile += "  X" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                }

                                // 13.2

                                if (jahrgang == 3 && hzJz == "JZ")
                                {
                                    if (IstWahlklausur(fach.Fach, schueler.Wahlklausur13_1)) // W
                                    {
                                        zeile += "  W" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                    else
                                    {
                                        zeile += "  X" + (belegung.Abifach == "4" ? "**(4)**" : "") + "  |";
                                    }
                                }
                            }
                        }
                        else
                        {
                            zeile += "     |";
                        }
                    }
                    File.AppendAllText(Global.TempPfad + datei, zeile.TrimEnd(' ') + wahlklausur + Environment.NewLine);
                }

                File.AppendAllText(Global.TempPfad + datei, "X: Belegung (ohne Klausur); P: Belegung (mit Pflichtklausur); W: Belegung (mit Wahlklausur); 1,2,3,4: Abiturfächer" + Environment.NewLine);
                File.AppendAllText(Global.TempPfad + datei, "" + Environment.NewLine);
            }
        }
        File.AppendAllText(Global.TempPfad + datei, "" + Environment.NewLine);
        File.AppendAllText(Global.TempPfad + datei, "Seite erstellt mit [[github>stbaeumer/Push2Dokuwiki|Push2Dokuwiki]]." + Environment.NewLine);

        File.AppendAllText(Global.TempPfad + datei, "" + Environment.NewLine);

        Global.Dateischreiben(datei);
    }

    public Kurswahlen()
    {
    }

    private bool IstWahlklausur(string fach, string wahlklausur)
    {
        if (
            fach.Contains(" G") &&
            wahlklausur != "" &&
            (from w in RemoveLineEndings(wahlklausur).Split(',').Select(p => p.Trim()).ToList() where fach.StartsWith(w) select w).Any())
        {
            return true;
        }
        return false;
    }

    private bool IstDeutschOderMathe(string fach)
    {
        if ((fach.StartsWith("D ") || fach.StartsWith("M ")) && fach.Contains(" G"))
        {
            return true;
        }
        return false;
    }

    private bool IstFremdsprache(string fach)
    {
        if ((fach.StartsWith("E ") || fach.StartsWith("N ") || fach.StartsWith("S ")) && fach.Contains(" G"))
        {
            return true;
        }
        return false;
    }

    private string RemoveLineEndings(string value)
    {
        if (String.IsNullOrEmpty(value))
        {
            return value;
        }
        string lineSeparator = ((char)0x2028).ToString();
        string paragraphSeparator = ((char)0x2029).ToString();

        return value.Replace("\r\n", string.Empty)
                    .Replace("\n", string.Empty)
                    .Replace("\r", string.Empty)
                    .Replace(lineSeparator, string.Empty)
                    .Replace(paragraphSeparator, string.Empty);
    }*/
}