namespace Common;

public class Anrechnung
{
    public double Wert { get; set; }
    public string? Text { get; set; }
    public int Grund { get; set; }
    public DateTime Von { get; set; }
    public DateTime Bis { get; set; }
    public string? Beschr { get; set; }
    public int TeacherIdUntis { get; internal set; }
    public Lehrer? Lehrer { get; internal set; }
    public string? Rolle { get; internal set; }
    public string? Amt { get; internal set; }
    public string? Hinweis { get; internal set; }
    public List<string>? Kategorien { get; internal set; }
    public string? Name { get; set; }
}