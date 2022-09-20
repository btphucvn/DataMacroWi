using DataMacroWi.DB;
using DataMacroWi.Model;
using MySql.Data.MySqlClient;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Service
{
    class AllKeyService
    {
        public int InsertPG(string keyID, string nameVi)
        {
            if(nameVi=="")
            {
                string bug = "";
            }
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            nameVi = nameVi.Replace("\"", "");
            string query = "insert into AllKeys(key_id,name_vi) values('"
                + keyID + "','"
                + nameVi + "') RETURNING id;";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            try
            {
                conn.Open();

                    int id = (int)cmd.ExecuteScalar();
                
                conn.Close();
                return id;
            }
            catch (Exception a)
            {
                string err = a.Message;
            }
            finally
            {
                conn.Close();
            }
            return -1;

        }

        public int Insert(string keyID, string nameVi)
        {
            DBConnect dBConnect = new DBConnect();
            MySqlConnection conn = dBConnect.ConnectMySQL();
            nameVi = nameVi.Replace("\"", "");
            string query = "insert into AllKeys(key_id,namevi) values('"
                + keyID + "','"
                + nameVi + "')";

            MySqlCommand cmd = new MySqlCommand(query, conn);
            try
            {
                conn.Open();
                cmd.ExecuteReader();
                int id = (int)cmd.LastInsertedId;
                conn.Close();
                return id;
            }
            catch (Exception a)
            {
                string err = a.Message;
            }
            finally
            {
                conn.Close();
            }
            return -1;

        }

        public AllKey GetAllKeyByKeyID(string keyID)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM allkeys WHERE key_id='"+ keyID+"'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();


                AllKey allKey = new AllKey();

                while (reader.Read())
                {
                    allKey.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    allKey.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    allKey.NameVi = reader.GetString(reader.GetOrdinal("name_vi"));
                    break;
                }
                conn.Close();
                return allKey;
            }
            catch (Exception e)
            {
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
            return null;
        }

        public bool CheckExitsAllKeyByKeyID(string keyID)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM allkeys WHERE key_id='" + keyID + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();



                while (reader.Read())
                {
                    conn.Close();
                    return true;
                }
                
            }
            catch (Exception e)
            {
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
            return false;
        }


        public string Update(string keyID, string nameVi)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            nameVi = nameVi.Replace("\"", "");
            string query = "UPDATE AllKeys "
                + "SET  name_vi = '" + nameVi + "' "
                + "WHERE key_id = '" + keyID+"'";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            if (CheckExitsAllKeyByKeyID(keyID)==true)
            {
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
                catch (Exception a)
                {
                    return a.Message;
                }
                finally
                {
                    conn.Close();
                }
            }
            else
            {
                if (InsertPG(keyID, nameVi) > 0) {
                    return "Cập nhật thất bại";
                }
            }
            return "Thành công";

        }

    }
}
