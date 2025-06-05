public class Klasse
{
    public int IdUntis { get; set; }
    public string? Name { get; set; }
    public List<Lehrer>? Klassenleitungen { get; set; }
    public bool IstVollzeit { get; internal set; }
    public string? BildungsgangLangname { get; internal set; }
    public string? BildungsgangGekürzt { get; internal set; }
    public string? WikiLink { get; internal set; }
    public string? Stufe { get; internal set; }
    public string? Raum { get; internal set; }
    public string? Jahrgang { get; set; }
    public string? Gliederung { get; set; }
    public string? Fachklassenschlüssel { get; set; }
    public string? Relationsgruppe { get; set; }
}