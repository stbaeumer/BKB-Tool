using System.Diagnostics;
using Common;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

#pragma warning disable CS0252 // Unbeabsichtigter Verweisvergleich. Wandeln Sie die linke Seite in den Typ "string" um, um einen Wertvergleich durchzuführen.
#pragma warning disable CA2200
#pragma warning disable CS8618
#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise nullverweisenden Objekts.
#pragma warning disable CS8604 // Möglicherweise wird ein nullverweisendes Argument an eine nicht-nullbare Parameterreferenz übergeben.
#pragma warning disable CS8073
#pragma warning disable CS0472
#pragma warning disable CS8625
#pragma warning disable CS8600
#pragma warning disable CS0162

public class Menue : List<Menüeintrag>
{
    public int AusgewaehlterMenueEintrag { get; set; }
    private Menüeintrag AusgewaehlterEintrag { get; set; }
    public Dateien Quelldateien { get; set; }
    public Klassen Klassen { get; set; }
    public Lehrers Lehrers { get; set; }

    public Menue() { }

    /// <summary>
    /// Menü aufbauen und Dateien Filtern
    /// </summary>
    public Menue(Dateien quelldateien, Klassen klassen, Lehrers lehrers, Students students, List<Menüeintrag> menueEintraege)
    {
        Quelldateien = quelldateien;
        Klassen = klassen;
        Lehrers = lehrers;

        if (!string.IsNullOrEmpty(Global.AktuellerPfad))
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = Global.AktuellerPfad,
                UseShellExecute = true,
                Verb = "open"
            });
            Global.AktuellerPfad = "";
        }

        //Global.WeiterMitAnyKey();
        //Global.DisplayHeader(Global.Header.H1, Global.H1!, Global.Protokollieren.Nein);


        Global.Protokoll = new List<string>();
        AddRange(menueEintraege);
        //AddRange(menueEintraege.Where(x => !x.MenüeintragAusblenden));
    }

    public void AuswahlGridRendern()
    {
        var grid = new Grid();
        grid.Expand();

        grid.AddColumn();
        grid.AddColumn();
        grid.AddColumn();

        
        for (int i = 0; i < this.Count(); i++)
        {
            var links = this[i].Titel.Split(':')[0].Trim();
            var rechts = this[i].Titel.Split(':')[1].Trim();
            if (this[i].Titel.Split(':').Count() > 2)
            {
                rechts = string.Join(":", this[i].Titel.Split(':').Skip(1).ToArray()).Trim();
            }

            grid.AddRow(new Text[]{
                new Text((i + 1).ToString(), new Style(Color.White, Color.Black)).RightJustified(),
                new Text(links, new Style(Color.SpringGreen2, Color.Black)).LeftJustified(),
                new Text(rechts, new Style(Color.White , Color.Black)).LeftJustified()
            });
        }
         
        AnsiConsole.Write(grid);
    }

    private string GetZulässigeAuswahl(List<string> weitere)
    {
        string zulässigeWerte = "";
        for (int i = 0; i < this.Count(); i++)
        {
            zulässigeWerte += ((i + 1).ToString()) + ",";
        }
        if (weitere != null && weitere.Count > 0)
        {
            zulässigeWerte += string.Join(",", weitere);
        }
        return zulässigeWerte.TrimEnd(','); // Entfernt das letzte Komma
    }

    public IConfiguration GetAusgewaehlterMenueintrag(IConfiguration configuration, List<string> weitere)
    {
        var zulässigeAuswahlOptionen = GetZulässigeAuswahl(weitere);

        configuration = Global.Konfig("Auswahl", Global.Modus.Update, configuration, "", -1, -1, "", null, null, zulässigeAuswahlOptionen);

        switch (configuration["Auswahl"])
        {
            case "h":
                Global.OpenWebseite(configuration["OnlineHilfeURL"]);
                configuration["Auswahl"] = "-1";
                return configuration;
            case "e":
                Global.DisplayHeader(configuration);
                configuration = Global.EinstellungenDurchlaufen(configuration, Global.Modus.Update);
                configuration["Auswahl"] = "-1";
                return configuration;
            default:
                if (int.TryParse(configuration["Auswahl"], out int auswahl) && auswahl > 0 && auswahl <= this.Count())
                {
                    configuration["Auswahl"] = (auswahl - 1).ToString();
                    return configuration;
                }
                else
                {
                    AnsiConsole.MarkupLine("[red]Ungültige Auswahl. Bitte erneut versuchen.[/]");
                    return configuration;
                }
        }
    }
}