public class Raum
{
    public int IdUntis { get; internal set; }
    public string? Raumnummer { get; internal set; }
    public int Anzahl { get; internal set; }
    
    public Raum()
    {
    }

    public Raum(string? raum)
    {
        Raumnummer = raum;
        Anzahl = 1;
    }
}