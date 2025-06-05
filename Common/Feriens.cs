    using Microsoft.Data.SqlClient;
    public class Feriens : List<Ferien>
    {
        public Feriens()
        {
            SqlConnection sqlConnection = new SqlConnection(Global.ConnectionStringUntis);

            try
            {
                sqlConnection.Open();

                string queryString = @"SELECT DISTINCT Holiday.Holiday_ID,
Holiday.Name, 
Holiday.Longname, 
Holiday.DateFrom, 
Holiday.DateTo, 
Holiday.Flags
FROM Holiday 
WHERE (((Holiday.SCHOOLYEAR_ID)=" + Global.AktSj[0] + Global.AktSj[1] + ") AND ((Holiday.SCHOOL_ID)=177659));";

                using (SqlCommand sqlCommand = new SqlCommand(queryString, sqlConnection))
                {
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();

                    while (sqlDataReader.Read())
                    {
                        Ferien ferien = new Ferien
                        {
                            Name = Global.SafeGetString(sqlDataReader, 1),
                            LangName = Global.SafeGetString(sqlDataReader, 2),
                            Von = DateTime.ParseExact((sqlDataReader.GetInt32(3)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Bis = DateTime.ParseExact((sqlDataReader.GetInt32(4)).ToString(), "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                            Feiertag = Global.SafeGetString(sqlDataReader, 5) == "F" ? true : false
                        };
                        
                        this.Add(ferien);
                    }
                    
                    sqlDataReader.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                sqlConnection.Close();
                Global.ZeileSchreiben(("Ferien").PadRight(70, '.'), this.Count().ToString(), ConsoleColor.Green, ConsoleColor.White);
            }
        }

        internal bool IstFerienTag(DateTime tag)
        {
            foreach (var ferien in this)
            {
                if (ferien.Von.Date <= tag.Date && tag.Date <= ferien.Bis.Date)
                {
                    return true;
                }
            }
            return false;
        }
    }