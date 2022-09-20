using DataMacroWi.DB;
using DataMacroWi.Model;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Service
{
    class YAxisService
    {
        public int GetYAxis(string unit)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM YAxis Where unit = '"+unit+"'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    return reader.GetInt32(reader.GetOrdinal("Value"));
                }
                conn.Close();
                return -1;
            }
            catch (Exception e)
            {
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
            return -1;
        }


    }
}
