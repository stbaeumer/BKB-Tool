using System.Data.OleDb;

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

internal class DataAccess
{
    public string ConnectionString { get; set; }

    public DataAccess(String connectionString)
    {
        ConnectionString = connectionString;
    }
            
    internal DateTime GetSchildVersion()
    {
        #pragma warning disable CA1416 // Plattformkompatibilität unterdrücken
        using (OleDbConnection oleDbConnection = new OleDbConnection(ConnectionString))
        {
            try
            {
                DateTime schildVersion = new DateTime();

                string queryString = @"SELECT Schild_Verwaltung.Version FROM Schild_Verwaltung;";

                OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                oleDbConnection.Open();
                OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                while (oleDbDataReader.Read())
                {
                    schildVersion = Convert.ToDateTime((oleDbDataReader["Version"]).ToString());
                }
                oleDbDataReader.Close();
                return schildVersion;
            }                
            catch (Exception ex)
            {   
                Global.ZeileSchreiben("Fehler:", ex.Message, ConsoleColor.Red, ConsoleColor.White);
                return new DateTime(0);
            }
            finally
            {
                oleDbConnection.Close();                    
            }            
        }
    }
    /*
    public Klassen GetSchoolClasses()
    {
        using (OleDbConnection oleDbConnection = new OleDbConnection(connetionString))
        {
            try
            {
                Klassen klassen = new Klassen();
                int aktSj = DateTime.Now.Month >= 8 ? DateTime.Now.Year : DateTime.Now.AddYears(-1).Year;

                string queryString = @"SELECT Schueler.ID, Schueler.Klasse, Schueler.FotoVorhanden, Len([SchuelerFotos].[Foto]) AS Laenge, Schueler.SchulnrEigner 
FROM Schueler LEFT JOIN SchuelerFotos ON Schueler.ID = SchuelerFotos.Schueler_ID
WHERE (((Schueler.Geloescht)='-') AND ((Schueler.Status)=2) AND ((Schueler.AktSchuljahr)=" + aktSj + @"))ORDER BY Schueler.Klasse, Schueler.Name, Schueler.Vorname;";
            
                OleDbCommand oleDbCommand = new OleDbCommand(queryString, oleDbConnection);
                oleDbConnection.Open();
                OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();

                while (oleDbDataReader.Read())
                {
                    Student student = new Student()
                    {
                        IdSchild = (oleDbDataReader.GetInt32(0)).ToString(),
                        Klasse = SafeGetString(oleDbDataReader, 1),
                        FotoVorhanden = SafeGetString(oleDbDataReader, 2) == "-" ? false : true,
                        FotoBinary = oleDbDataReader.GetValue(3) == DBNull.Value ? false : true,
                        SchulnrEigner = oleDbDataReader.GetInt32(4)
                    };

                    Klasse klasse = (from kl in klassen where kl.Name == student.Klasse select kl).FirstOrDefault();

                    if (student.FotoBinary && !student.FotoVorhanden)
                    {
                        oleDbCommand = new OleDbCommand();
                        oleDbCommand = oleDbConnection.CreateCommand();
                        oleDbCommand.CommandText = "UPDATE Schueler SET FotoVorhanden = '+' WHERE ID = @id";
                        oleDbCommand.Parameters.AddWithValue("@id", student.IdSchild);
                        oleDbCommand.ExecuteNonQuery();
                        student.FotoVorhanden = true;
                    }
                    if (!student.FotoBinary && student.FotoVorhanden)
                    {
                        oleDbCommand = new OleDbCommand();
                        oleDbCommand = oleDbConnection.CreateCommand();
                        oleDbCommand.CommandText = "UPDATE Schueler SET FotoVorhanden = '-' WHERE ID = @id";
                        oleDbCommand.Parameters.AddWithValue("@id", student.IdSchild);
                        oleDbCommand.ExecuteNonQuery();
                        student.FotoVorhanden = true;
                    }

                    if (klasse != null)
                    {
                        klasse.Schuelers.Add(student);
                    }
                    else
                    {
                        klasse = new Klasse()
                        {
                            Name = student.Klasse
                        };
                        if (klasse.Name != "")
                        {
                            klasse.Schuelers = new Students(); 
                            {
                                student
                            };
                            klassen.Add(klasse);
                        }
                    }
                }
                oleDbDataReader.Close();
                return klassen;
            }
            catch (Exception ex)
            {
                
            }
            finally
            {
                oleDbConnection.Close();
            }                
        }
    }*/

    public string SafeGetString(OleDbDataReader reader, int colIndex)
    {
        if (!reader.IsDBNull(colIndex))
            return reader.GetString(colIndex);
        return string.Empty;
    }

    internal string InsertImage(Student student)
    {
        #pragma warning disable CA1416 // Plattformkompatibilität unterdrücken
        using (OleDbConnection oleDbConnection = new OleDbConnection(ConnectionString))
        {
            try
            {
                oleDbConnection.Open();
                OleDbCommand oleDbCommand = new OleDbCommand();
                oleDbCommand = oleDbConnection.CreateCommand();
                oleDbCommand.CommandText = "UPDATE Schueler SET FotoVorhanden = '+' WHERE ID = @id";
                oleDbCommand.Parameters.AddWithValue("@id", student.IdSchild);
                oleDbCommand.ExecuteNonQuery();
                student.FotoVorhanden = true;
                oleDbCommand = new OleDbCommand();
                oleDbCommand = oleDbConnection.CreateCommand();
                oleDbCommand.CommandText = "insert into SchuelerFotos (Schueler_ID, Foto, SchulnrEigner) values (@id, @foto, @schulnrEigner)";
                oleDbCommand.Parameters.AddWithValue("@id", student.IdSchild);
                oleDbCommand.Parameters.AddWithValue("@foto", student.Foto != null ? Convert.FromBase64String(student.Foto) : DBNull.Value);
                oleDbCommand.Parameters.AddWithValue("@schulnrEigner", student.SchulnrEigner);
                oleDbCommand.ExecuteNonQuery();
                student.FotoBinary = true;
                return "ok";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                oleDbConnection.Close();
            }
        }
    }

    internal string DeleteImage(Student student)
    {
        using (OleDbConnection oleDbConnection = new OleDbConnection(ConnectionString))
        {
            try
            {
                oleDbConnection.Open();
                OleDbCommand oleDbCommand = new OleDbCommand();
                oleDbCommand = oleDbConnection.CreateCommand();
                oleDbCommand.CommandText = "DELETE * FROM SchuelerFotos WHERE SchuelerFotos.Schueler_ID = @id";
                oleDbCommand.Parameters.AddWithValue("@id", student.IdSchild);
                oleDbCommand.ExecuteNonQuery();
                oleDbCommand = new OleDbCommand();
                oleDbCommand = oleDbConnection.CreateCommand();
                oleDbCommand.CommandText = "UPDATE Schueler SET FotoVorhanden = '-' WHERE ID = @id";
                oleDbCommand.Parameters.AddWithValue("@id", student.IdSchild);
                oleDbCommand.ExecuteNonQuery();
                student.FotoVorhanden = false;
                return "ok";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                oleDbConnection.Close();
            }
        }
    }
}
