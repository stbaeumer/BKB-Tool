using System.Globalization;
using Microsoft.Extensions.Configuration;

public class Lehrer
{
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

    public Lehrer()
    {
    }

    public int IdUntis { get; internal set; }
    public string Kürzel { get; internal set; }
    public string? Mail { get; internal set; }
    public string Geschlecht { get; internal set; }
    public double Deputat { get; internal set; }
    public string? Nachname { get; internal set; }
    public string? Vorname { get; internal set; }
    public string? Titel { get; internal set; }
    public string? Raum { get; internal set; }
    public DateTime Geburtsdatum { get; internal set; }
    public double AusgeschütteteAltersermäßigung { get; internal set; }
    public int ProzentStelleInSchild { get; internal set; }
    public int AlterAmErstenSchultagDesJahres { get; internal set; }
    public string? Flags { get; internal set; }
    public string Beschreibung { get; internal set; }
    public string? Text2 { get; internal set; }
    public double DeputatLautSchild { get; internal set; }
    public int AlterAmErstenSchultagDesKommendenJahres { get; internal set; }
    public double AltersermäßigungSoll { get; internal set; }
    public double DeputatLautUntis { get; set; }
    public double AltersermäßigungSollKommendes { get; internal set; }

    internal int GetAlterAmErstenSchultagDesSchuljahres(int jahr)
    {
        int years = jahr - Geburtsdatum.Year;
        DateTime birthday = Geburtsdatum.AddYears(years);
        if (new DateTime(jahr, 7, 31).CompareTo(birthday) < 0)
        {
            years--;
        }

        return years;
    }

    internal int GetProzentStelle(IConfiguration configuration)
    {
        var volleStelle = float.TryParse(configuration["VolleStelle"]?.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out float volleStelleValue) ? volleStelleValue : 25.5f;
        
        return Convert.ToInt32(Math.Floor(100 / volleStelle * this.DeputatLautSchild));
    }

    public DateTime GetGeburtsdatum(string geburtsdatum)
    {        
        DateTime geburtsdatumDateTime;
        
        if (!DateTime.TryParseExact(geburtsdatum, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out geburtsdatumDateTime))
        {
            // Fehlerbehandlung, falls das Datum nicht konvertiert werden kann
            return DateTime.MinValue;
        }

        // Das Geburtsdatum wird auf den 31.07. des Jahres gesetzt, um das Alter am ersten Schultag zu berechnen.
        return geburtsdatumDateTime;
    }

    internal double CheckAltersermäßigung(int alterAmErstenSchultagDesJahres, double ausgeschütteteAltersermäßigung)
    {
        switch (alterAmErstenSchultagDesJahres)
        {
            case >= 60:
                {
                    switch (ProzentStelleInSchild)
                    {
                        case >= 96 when (ausgeschütteteAltersermäßigung != 3 || ausgeschütteteAltersermäßigung < 0):
                            {
                                return 3;
                            }
                        case >= 75 and < 96 when (ausgeschütteteAltersermäßigung != 2 || ausgeschütteteAltersermäßigung < 0):
                            {
                                return 2;
                            }
                        case >= 50 and < 75 when (ausgeschütteteAltersermäßigung != 1.5 || ausgeschütteteAltersermäßigung < 0):
                            {
                                return 1.5;
                            }
                    }

                    break;
                }
            case >= 55 and < 60:
                {
                    switch (ProzentStelleInSchild)
                    {
                        case 100 when (ausgeschütteteAltersermäßigung != 1 || ausgeschütteteAltersermäßigung < 0):
                            {
                                return 1;
                            }
                        case >= 50 and < 100 when (ausgeschütteteAltersermäßigung != 0.5 || ausgeschütteteAltersermäßigung < 0):
                            {
                                return 0.5;
                            }
                    }
                    break;
                }
        }
        return 0;
    }

    internal double CheckAltersermäßigungSoll(int alterAmErstenSchultagDesJahres)
    {
        switch (alterAmErstenSchultagDesJahres)
        {
            case >= 60:
            {
                switch (ProzentStelleInSchild)
                {
                    case >= 96:
                    {
                        return 3;
                    }
                    case >= 75 and < 96:
                    {                            
                        return 2;
                    }
                    case >= 50 and < 75:
                    {                        
                        return 1.5;
                    }
                }

                break;
            }
            case >= 55 and < 60:
            {
                switch (ProzentStelleInSchild)
                {
                    case 100:
                    {                        
                        return 1;
                    }
                    case >= 50 and < 100:
                    {                        
                        return 0.5;
                    }
                }
                break;
            }
        }
        return 0;
    }

    internal double GetDeputatSchild(string? deputatString)
    {
        double deputatDoubleLautSchild = 0.0;
        double.TryParse(deputatString?.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out deputatDoubleLautSchild);
        return deputatDoubleLautSchild;
    }

    internal int GetAlterAmErstenSchultagDesSchuljahres(int akt, object geburtsdatumDateTime)
    {
        throw new NotImplementedException();
    }

    internal double GetDeputatLautUntis(List<dynamic> gpu004)
    {
        var deputatLautUntisString = gpu004
        .Where(rec => ((IDictionary<string, object>)rec)["Field1"].ToString() == Kürzel)
                    .Select(rec => ((IDictionary<string, object>)rec)["Field15"].ToString())
                    .FirstOrDefault();

        if (deputatLautUntisString != null)
        {
            double deputatLautUntisDouble = 0.0;
            double.TryParse(deputatLautUntisString.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out deputatLautUntisDouble);
            return deputatLautUntisDouble;
        }
        else
        {
            return 0;
        }
    }

    public double GetAnrechnungswertGPU020Soll(List<dynamic> gpu020, string bisherigerAnrechnungsgrund)
    {
        if (gpu020 == null || gpu020.Count == 0) return 0;

        var wertSoll = 0.0;

        var anrechnungenDesLehrers = gpu020.Where(rec =>
        {
            return ((IDictionary<string, object>)rec)["Field4"].ToString() == Kürzel &&
                    ((IDictionary<string, object>)rec)["Field5"].ToString() == bisherigerAnrechnungsgrund;
        });

        foreach (var anrechnungsgrund in anrechnungenDesLehrers)
        {
            var wert = Convert.ToDouble(((IDictionary<string, object>)anrechnungsgrund)["Field13"].ToString()) / 1000;

            var von = new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1);
            var bis = new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31);

            // Field7 prüfen: Nur überschreiben, wenn nicht leer
            var field7 = ((IDictionary<string, object>)anrechnungsgrund)["Field7"]?.ToString();
            if (!string.IsNullOrWhiteSpace(field7))
            {
                // yyyymmdd nach DateTime parsen
                if (DateTime.TryParseExact(field7, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedVon))
                {
                    von = parsedVon;
                }
            }

            // Field8 prüfen: Nur überschreiben, wenn nicht leer
            var field8 = ((IDictionary<string, object>)anrechnungsgrund)["Field8"]?.ToString();
            if (!string.IsNullOrWhiteSpace(field8))
            {
                if (DateTime.TryParseExact(field8, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedBis))
                {
                    bis = parsedBis;
                }
            }

            // Der angegebene Wert bezieht sich auf den Zeitraum von Beginn bis Ende des Schuljahres
            // Wenn der Zeitraum eingeschränkt ist, wird der Wert entsprechend angepasst

            if (von > new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1))
            {
                wert = wert * (bis - von).TotalDays / (bis - new DateTime(Convert.ToInt32(Global.AktSj[0]), 8, 1)).TotalDays;
            }
            else if (bis < new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31))
            {
                wert = wert * (bis - von).TotalDays / (new DateTime(Convert.ToInt32(Global.AktSj[1]), 7, 31) - von).TotalDays;
            }

            wertSoll += wert;
        }

        // Es wird auf zwei Nachkommastellen gerundet
        wertSoll = Math.Round(wertSoll, 2);
        return wertSoll;
    }

    internal double GetWertSonderzeiten(string? grund, List<dynamic> lehrkraefteSonderzeiten)
    {
        var l = lehrkraefteSonderzeiten
                    .Where(rec => ((IDictionary<string, object>)rec)["Grund"].ToString() == grund)
                    .Where(rec => ((IDictionary<string, object>)rec)["Lehrkraft"].ToString() == Kürzel)
                    .Select(rec => ((IDictionary<string, object>)rec)["Anzahl Stunden"].ToString())
                    .FirstOrDefault();
        if (l != null && double.TryParse(l.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double wert))
        {
            return Math.Round(wert, 2);
        }
        else
        {
            return 0.0;
        }
    }
}