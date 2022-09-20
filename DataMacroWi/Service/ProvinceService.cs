using DataMacroWi.DB;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Service
{
    class ProvinceService
    {
        public void Insert(string province, string region)
        {

            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "insert into provinces(province,region) values('"
                + province + "','"
                + region + "') RETURNING id;";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            try
            {
                conn.Open();

                int id = (int)cmd.ExecuteScalar();

                conn.Close();
            }
            catch (Exception a)
            {
                string err = a.Message;
            }
            finally
            {
                conn.Close();
            }
        }

        public string GetRegion(string province)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM provinces WHERE province='" + province + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                string result = "";

                while (reader.Read())
                {
                    result =  reader.GetString(reader.GetOrdinal("region"));
                }
                conn.Close();

                return result;
            }
            catch (Exception e)
            {
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
            return "";
        }
    }
}
