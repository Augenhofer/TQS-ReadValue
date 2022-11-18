using log4net;
using log4net.Config;
using System;
using System.Data;
using System.IO;
using System.Reflection;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace TQS_Read
{
    // ==========================================================================
    //
    //  Programm:       WE_E_Mail
    //
    //  (c) Copyright by voestalpine Stahl Donawitz GmbH
    //  Kerpelystraße 199  A-8700 Leoben/Austria
    //  www.voestalpine.com  
    // --------------------------------------------------------------------------
    //  Description:
    //    the porgramm read the daily TS - Data from e-Mail and store it to an database
    //
    //  The code is structured into the following parts:
    // --------------------------------------------------------------------------
    //  History:
    //
    //	2022-11-15 Christoph Augenhofer
    //	- Erstellen des Programmes
    //
    // ==========================================================================
    class Program
    {
        public static string fileName = @"D:\Versandanalysen\TQSFile.txt";
        //public static string fileName = @"c:\temp\TQSFile.txt";
        public static string text;
        private static readonly ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            var logRepository = LogManager.GetRepository(Assembly.GetEntryAssembly());
            XmlConfigurator.Configure(logRepository, new FileInfo("log4net.config"));
            bool isNumeric = false;

            try
            {
                text = File.ReadAllText(fileName);
                string[] lines = text.Split(Environment.NewLine);

                for (int i = 1; i < lines.Length; i++)
                {
                    string[] tokens = lines[i].Split('|');
                    if (tokens.Length > 3)
                    {
                        text = tokens[3].ToString();
                        isNumeric = int.TryParse(tokens[1].ToString(), out int n);
                        if (tokens[3].ToString() != "   "  && isNumeric == true)
                            SaveData(tokens, log);
                    }
                }
            }
            catch (Exception e)
            {
                log.Error("Fehler: " + e.Message);
            }
        }
        public static void SaveData(string[] tokens, ILog log)
        {
            DateTime dt = new DateTime(DateTime.Now.Year, DateTime.Now.AddDays(-1).Month, Convert.ToInt32(tokens[1]));
            // define connection string and query
            string connectionString = "Data Source=2217DBDO01;Initial Catalog=XPSDetailData;User id=XPS_Detail;Password=!pwXPS_Detail";
            string query = @"IF NOT EXISTS(SELECT * FROM TQS_Data WHERE Datum = @Datum)
                                INSERT INTO TQS_Data(Datum, RE_IST, Chargen_SOLL, Chargen_IST, Rohstahl_fest) VALUES(@Datum, @RE_IST, @Chargen_SOLL, @Chargen_IST, @Rohstahl_fest);";

            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                // define the parameters - not sure just how large those 
                // string lengths need to be - use whatever is defined in the
                // database table here!
                

                cmd.Parameters.Add("@Datum", SqlDbType.DateTime, 100).Value = dt;
                cmd.Parameters.Add("@RE_IST", SqlDbType.Int, 200).Value = Convert.ToInt32(tokens[2]);
                cmd.Parameters.Add("@Chargen_SOLL", SqlDbType.Int, 200).Value = Convert.ToInt32(tokens[23]);
                cmd.Parameters.Add("@Chargen_IST", SqlDbType.Int, 200).Value = Convert.ToInt32(tokens[22]);
                cmd.Parameters.Add("@Rohstahl_fest", SqlDbType.Int, 200).Value = Convert.ToInt32(tokens[24]);

                // open connection, execute query, close connection
                try
                {
                    conn.Open();
                    int rowsAffected = cmd.ExecuteNonQuery();
                    Console.Write("MSSQL rows affected: " + rowsAffected);
                    log.Error("MSSQL rows affected: " + rowsAffected);
                    conn.Close();
                }
                catch (Exception e)
                {
                    log.Error("MSSQL-Fehler: " + e.Message);
                }
            }
            query = @"IF EXISTS(SELECT * FROM Aprodukt WHERE Datum = @Datum)
                        UPDATE Aprodukt 
                        SET Ist_Chargen = @Ist_Chargen
                        WHERE Datum = @Datum";
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlCommand cmd = new SqlCommand(query, conn))
            {
                // define the parameters - not sure just how large those 
                // string lengths need to be - use whatever is defined in the
                // database table here!
                cmd.Parameters.Add("@Datum", SqlDbType.DateTime, 100).Value = dt;
                cmd.Parameters.Add("@Ist_Chargen", SqlDbType.Int, 200).Value = Convert.ToInt32(tokens[22]);

                // open connection, execute query, close connection
                conn.Open();
                int rowsAffected = cmd.ExecuteNonQuery();
                conn.Close();
            }

        }

    }   
}