using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Spectre.Console;

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

public class Raums : List<Raum>
{
    public int Anzahl { get; internal set; }

    public Raums()
    {
    }

    public Raums(int periode, IConfiguration configuration)
    {
        using (SqlConnection odbcConnection = new SqlConnection(configuration["ConnectionStringUntis"]))
        {
            try
            {
                string queryString = @"SELECT Room.ROOM_ID, 
                                                Room.Name,  
                                                Room.Longname,
                                                Room.Capacity
                                                FROM Room
                                                WHERE (((Room.SCHOOLYEAR_ID)= " + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Room.SCHOOL_ID)=177659) AND  ((Room.TERM_ID)=" + periode + "))";

                SqlCommand odbcCommand = new SqlCommand(queryString, odbcConnection);
                odbcConnection.Open();
                SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                while (sqlDataReader.Read())
                {
                    Raum raum = new Raum()
                    {
                        IdUntis = sqlDataReader.GetInt32(0),
                        Raumnummer = Global.SafeGetString(sqlDataReader, 1)
                    };

                    this.Add(raum);
                }

                sqlDataReader.Close();
            }
            catch (Microsoft.Data.SqlClient.SqlException ex)
            {
                Console.WriteLine("Netzwerkbezogener oder instanzspezifischer Fehler beim Herstellen einer Verbindung mit SQL Server. Der Server wurde nicht gefunden, oder auf ihn kann nicht zugegriffen werden.");
                Console.ReadKey();
            }
            finally
            {
                odbcConnection.Close();
                //Global.ZeileSchreiben("Räume", this.Count().ToString(), ConsoleColor.Green, ConsoleColor.White);
            }
        }
    }
}