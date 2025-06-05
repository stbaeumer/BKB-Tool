public class Relationsgruppe
{
    public string? BeschreibungSchulministerium { get; private set; }

    public List<string?> Jahrgänge { get; private set; }
    public List<string?> Fachklassenschlüssel { get; private set; }
    public List<string?> Gliederungen { get; private set; }
    public double Relation { get; private set; }

    public Relationsgruppe(string? beschreibungSchulministerium, List<string?> gliederungen, List<string?> jahrgänge,
        List<int> fachklassenschlüssel, double relation)
    {
        this.Relation = relation;
        this.Fachklassenschlüssel = new List<string?>();
        this.BeschreibungSchulministerium = beschreibungSchulministerium;
        this.Gliederungen = gliederungen;
        this.Jahrgänge = jahrgänge;
        try
        {
            if (fachklassenschlüssel == null) return;
            foreach (var item in fachklassenschlüssel)
            {
                this.Fachklassenschlüssel.Add(item.ToString());
            }
        }
        catch
        {
            // ignored
        }
    }
}