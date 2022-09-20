using DataMacroWi.DB;
using DataMacroWi.Extension;
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
    class RowDataLevel3Service
    {
        public int Insert(Row_Data_Level3 row_Data_Level3)
        {
            DBConnect dBConnect = new DBConnect();
            MySqlConnection conn = dBConnect.ConnectMySQL();

            string query = "insert into Row_Data_Level3s(key_id,unit,id_row_data_level2) values('"
                + row_Data_Level3.KeyID + "','"
                + row_Data_Level3.Unit + "','"
                + row_Data_Level3.IdRowDataLevel2 + "');";

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
        public int InsertPG(Row_Data_Level3 row_Data_Level3)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "insert into Row_Data_Level3s(key_id,unit,stt,id_row_data_level2) values('"
                + row_Data_Level3.KeyID + "','"
                + row_Data_Level3.Unit + "','"
                + row_Data_Level3.Stt + "','"
                + row_Data_Level3.IdRowDataLevel2 + "') RETURNING id;";
            AllKeyService allKeyService = new AllKeyService();
            if (!allKeyService.CheckExitsAllKeyByKeyID(row_Data_Level3.KeyID))
            {
                allKeyService.InsertPG(row_Data_Level3.KeyID, row_Data_Level3.Name);
            }
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

        public List<Row_Data_Level3> Get_RowDataLevel3_By_IdRowLevel2(int id_row_level2)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level3s WHERE id_row_data_level2='" + id_row_level2 + "' ORDER BY stt ASC";
            
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level3> list = new List<Row_Data_Level3>();
                while (reader.Read())
                {
                    Row_Data_Level3 row_Data_Level = new Row_Data_Level3();

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    row_Data_Level.IdRowDataLevel2 = reader.GetInt32(reader.GetOrdinal("id_row_data_level2"));

                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(row_Data_Level.KeyID);
                    row_Data_Level.Name = allKey.NameVi;
                    try
                    {
                        row_Data_Level.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    }
                    catch { }
                    row_Data_Level.Unit = reader.GetString(reader.GetOrdinal("unit"));

                    list.Add(row_Data_Level);
                }
                conn.Close();
                return list;
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

        public Row_Data_Level3 Get_RowDataLevel3_By_IdRowLevel2_KeyID(int id_row_level2,string keyID)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level3s WHERE " +
                " key_id='"+keyID+"'" +
                " AND id_row_data_level2='" + id_row_level2 + "' ORDER BY stt ASC";

            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level3> list = new List<Row_Data_Level3>();
                Row_Data_Level3 row_Data_Level = new Row_Data_Level3();

                while (reader.Read())
                {

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    row_Data_Level.IdRowDataLevel2 = reader.GetInt32(reader.GetOrdinal("id_row_data_level2"));

                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(row_Data_Level.KeyID);
                    row_Data_Level.Name = allKey.NameVi;
                    try
                    {
                        row_Data_Level.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    }
                    catch { }
                    row_Data_Level.Unit = reader.GetString(reader.GetOrdinal("unit"));

                }
                conn.Close();
                return row_Data_Level;
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

    }
}
