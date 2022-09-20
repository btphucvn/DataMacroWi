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
    class RowDataLevel2Service
    {
        public int Insert(Row_Data_Level2 row_Data_Level2)
        {
            DBConnect dBConnect = new DBConnect();
            MySqlConnection conn = dBConnect.ConnectMySQL();

            string query = "insert into Row_Data_Level2s(keyID,unit,idRowDataLevel1) values('"
                + row_Data_Level2.KeyID + "','"
                + row_Data_Level2.Unit + "','"
                + row_Data_Level2.IdRowDataLevel1 + "');";

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

        public int InsertPG(Row_Data_Level2 row_Data_Level2)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();


            string query = "insert into Row_Data_Level2s(key_id,unit,stt,id_row_data_level1) values('"
                + row_Data_Level2.KeyID + "','"
                + row_Data_Level2.Unit + "','"
                + row_Data_Level2.Stt + "','"
                + row_Data_Level2.IdRowDataLevel1 + "') RETURNING id;";
            AllKeyService allKeyService = new AllKeyService();
            if (!allKeyService.CheckExitsAllKeyByKeyID(row_Data_Level2.KeyID))
            {
                allKeyService.InsertPG(row_Data_Level2.KeyID, row_Data_Level2.Name);
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

        public List<Row_Data_Level2> Get_RowDataLevel2_By_IdRowLevel1(int id_row_level1)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level2s WHERE id_row_data_level1='" + id_row_level1 + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level2> list = new List<Row_Data_Level2>();
                while (reader.Read())
                {
                    Row_Data_Level2 row_Data_Level = new Row_Data_Level2();

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    row_Data_Level.IdRowDataLevel1 = reader.GetInt32(reader.GetOrdinal("id_row_data_level1"));

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

        public Row_Data_Level2 Get_RowDataLevel2_By_IdRowLevel1_KeyID(int id_row_level1,string keyID)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level2s WHERE " +
                "key_id ='"+keyID+"'"+
                " AND id_row_data_level1='" + id_row_level1 + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level2> list = new List<Row_Data_Level2>();
                Row_Data_Level2 row_Data_Level = new Row_Data_Level2();

                while (reader.Read())
                {

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    row_Data_Level.IdRowDataLevel1 = reader.GetInt32(reader.GetOrdinal("id_row_data_level1"));

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

        public Row_Data_Level2 Get_RowDataLevel2_By_IDTable_KeyIDLevel1_KeyIDLevel2(int idTable,string keyIDLevel1, string keyIDLevel2)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
            Row_Data_Level1 row_Data_Level1 = new Row_Data_Level1();
            row_Data_Level1 = rowDataLevel1Service.Get_RowDataLevel1_By_IdTable_KeyID(idTable, keyIDLevel1);
            string query = "SELECT * FROM row_data_level2s WHERE " +
                "key_id ='" + keyIDLevel2 + "'" +
                " AND id_row_data_level1='" + row_Data_Level1.Id + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Row_Data_Level2 row_Data_Level = new Row_Data_Level2();

                while (reader.Read())
                {

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    row_Data_Level.IdRowDataLevel1 = reader.GetInt32(reader.GetOrdinal("id_row_data_level1"));

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
                if (row_Data_Level.KeyID!=null)
                {
                    return row_Data_Level;
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
            return null;
        }

    }
}
