/* 
Author: Jimmy Erikson
Description: 
    Google Shopping Content Product [Categories].XSL file converted to .MDB file
    Added .MDB as OleDb DataConnection, using good old ADO reader to query product categories and their unique identifiers.
    https://www.google.com/basepages/producttype/taxonomy-with-ids.en-US.txt
*/
using System;
using System.Data;
using System.Data.OleDb;
namespace GoogleMdbTest
{
    class Program
    {
        public static string ConnectionString = (@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source = D:\dev\Github\Jenpbiz\Jenpbiz\GoogleCategories.mdb");
        public static string QueryString = ("SELECT * FROM GoogleCategories_Table WHERE F2 LIKE '%' + @Category +'%'");

        static void Main(string[] args) {
            ReadData(ConnectionString, "Software");
        }

        private static void ReadData(string connectionString, string searcString)
        {
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                OleDbCommand command = new OleDbCommand(QueryString, connection);
                command.Parameters.Add("@Category", OleDbType.VarChar,50).Value = searcString;

                Console.WriteLine("Establishing connection . . .");
                connection.Open();

                OleDbDataReader reader;
                reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        string
                            F0 = reader[0].ToString(),
                            F1 = reader[1].ToString(),
                            F2 = reader[2].ToString(),
                            F3 = reader[3].ToString(),
                            F4 = reader[4].ToString(),
                            F5 = reader[5].ToString(),
                            F6 = reader[6].ToString(),
                            F7 = reader[7].ToString();
                        Console.WriteLine(F0 + " " + F1 + " " + F2 + " " + F3 + " " + F4 + " " + F5 + " " + F6 + " " + F7);
                    }
                }
                catch (Exception e) {
                    Console.WriteLine("Error: " + e.Message);
                }
                Console.WriteLine("\nClosing connection . . . ");
                reader.Close();
                Console.WriteLine("Connection closed . . .");
            }
        }
    }
}

