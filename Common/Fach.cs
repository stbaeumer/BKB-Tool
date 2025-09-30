internal class Kurswahl
{
    public DateTime Geburtsdatum { get; internal set; }
    public bool Volljährig { get; internal set; }
    public string HzJz { get; internal set; }
    public string Anlage { get; internal set; }
    public string Zeugnisart { get; internal set; }
    public string Zeugnistext { get; internal set; }
    public DateTime Konferenzdatum { get; internal set; }
    public DateTime DatumReligionAbmeldung { get; internal set; }
    public bool SchuelerAktivInDieserKlasse { get; internal set; }
    public object Abschlussklasse { get; internal set; }
    public string Beschreibung { get; internal set; }
    public object ReligionAbgewählt { get; internal set; }
    public string Klasse { get; internal set; }
    public string Gesamtnote { get; internal set; }
    public string Fach { get; internal set; }
    public string Gesamtpunkte_12_1 { get; internal set; }
    public int LeistungId { get; internal set; }
    public int NokId { get; internal set; }
    public int SchlüsselExtern { get; internal set; }
    public string Schuljahr { get; internal set; }
    public string Gliederung { get; internal set; }
    public bool HatBemerkung { get; internal set; }
    public int Jahrgang { get; internal set; }
    public string Name { get; internal set; }
    public string Nachname { get; internal set; }
    public string Vorname { get; internal set; }
    public int LehrkraftAtlantisId { get; internal set; }
    public string Bereich { get; internal set; }
    public int Reihenfolge { get; internal set; }
    public string Gesamtpunkte_12_2 { get; internal set; }
    public string Gesamtpunkte_13_1 { get; internal set; }
    public string Gesamtpunkte_13_2 { get; internal set; }
    public string Gesamtpunkte { get; internal set; }
    public string Tendenz { get; internal set; }
    public string EinheitNP { get; internal set; }
    public string Eingebracht_12_1 { get; internal set; }
    public string Eingebracht_12_2 { get; internal set; }
    public string Eingebracht_13_1 { get; internal set; }
    public string Eingebracht_13_2 { get; internal set; }
    public string Klassenleiter { get; internal set; }
    public string KlassenleiterName { get; internal set; }
    public string Abifach { get; internal set; }
    public string Lehrkraft { get; internal set; }

    internal bool IstAbifach1bis3()
    {
        if (Abifach != null && Abifach != "" && Convert.ToInt32(Abifach) < 4)
        {
            return true;
        }
        return false;
    }
}