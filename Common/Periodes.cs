using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Exception = System.Exception;

public class Periodes : List<Periode>
{
    public Periodes(IConfiguration configuration)
    {
        SqlConnection sqlConnection = new SqlConnection(configuration["ConnectionStringUntis"]);

        try
        {
            sqlConnection.Open();

            string queryString = @"SELECT DISTINCT
Terms.TERM_ID, 
Terms.Name, 
Terms.Longname, 
Terms.DateFrom, 
Terms.DateTo
FROM Terms
WHERE (((Terms.SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] +
                                 ")  AND ((Terms.SCHOOL_ID)=177659)) ORDER BY Terms.TERM_ID;";

            using (SqlCommand sqlCommand = new SqlCommand(queryString, sqlConnection))
            {
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

                while (sqlDataReader.Read())
                {
                    Periode periode = new Periode()
                    {
                        IdUntis = sqlDataReader.GetInt32(0),
                        Name = Global.SafeGetString(sqlDataReader, 1),
                        Langname = Global.SafeGetString(sqlDataReader, 2),
                        Von = DateTime.ParseExact((sqlDataReader.GetInt32(3)).ToString(), "yyyyMMdd",
                            System.Globalization.CultureInfo.InvariantCulture),
                        Bis = DateTime.ParseExact((sqlDataReader.GetInt32(4)).ToString(), "yyyyMMdd",
                            System.Globalization.CultureInfo.InvariantCulture)
                    };

                    this.Add(periode);
                }
                // Korrektur des Periodenendes

                for (int i = 0; i < this.Count - 1; i++)
                {
                    this[i].Bis = this[i + 1].Von.AddDays(-1);
                }

                sqlDataReader.Close();
                //Global.ZeileSchreiben("Perioden", this.Count().ToString(), ConsoleColor.Green, ConsoleColor.White);
            }
        }
        catch
        {
        }
        finally
        {
            sqlConnection.Close();
        }
    }

    public int GetAktuellePeriode()
    {
        var aktuellePeriode = (from p in this where p.Bis >= DateTime.Now.Date where DateTime.Now.Date >= p.Von select p.IdUntis)
            .FirstOrDefault();

        if (aktuellePeriode == 0)
        {
            Console.WriteLine("Es kann keine aktuelle Periode ermittelt werden. Das ist z. B. während der Sommerferien der Fall.");
        }
        else
        {
            //Global.ZeileSchreiben("Aktuelle Periode", this.Count().ToString(), ConsoleColor.Green, ConsoleColor.White);
        }

        return aktuellePeriode;
    }
}