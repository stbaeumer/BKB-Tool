using Microsoft.Data.SqlClient;
public class Faechers : List<Fach>
{
    public Faechers(int periode)
    {
        using (SqlConnection odbcConnection = new SqlConnection(Global.ConnectionStringUntis))
        {
            try
            {
                string queryString = @"SELECT DISTINCT 
Subjects.Subject_ID,
Subjects.Name,
Subjects.Longname,
Subjects.Text,
Description.Name
FROM Description RIGHT JOIN Subjects ON Description.DESCRIPTION_ID = Subjects.DESCRIPTION_ID
WHERE Subjects.Schoolyear_id = " + Global.AktSj[0] + Global.AktSj[1] +
                                     " AND Subjects.Deleted='false'  AND ((Subjects.SCHOOL_ID)=177659) ORDER BY Subjects.Name;";

                SqlCommand odbcCommand = new SqlCommand(queryString, odbcConnection);
                odbcConnection.Open();
                SqlDataReader sqlDataReader = odbcCommand.ExecuteReader();

                while (sqlDataReader.Read())
                {
                    Fach fach = new Fach()
                    {
                        IdUntis = sqlDataReader.GetInt32(0),
                        FachKürzelUntis = Global.SafeGetString(sqlDataReader, 1),
                        FachLangnameUntis = Global.SafeGetString(sqlDataReader, 3),
                        FachBeschrUntis = Global.SafeGetString(sqlDataReader, 4)
                    };

                    this.Add(fach);
                }

                sqlDataReader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                odbcConnection.Close();
                Global.ZeileSchreiben( "Fächer aus Untis", this.Count().ToString(),ConsoleColor.DarkCyan);
            }
        }
    }
}