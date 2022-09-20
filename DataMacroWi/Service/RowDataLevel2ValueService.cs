using DataMacroWi.DB;
using DataMacroWi.Extension;
using DataMacroWi.Model;
using MySql.Data.MySqlClient;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DataMacroWi.Service
{
    class RowDataLevel2ValueService
    {
        public int Insert(Row_Data_Level2_Value row_Data_Level2_Value)
        {
            DBConnect dBConnect = new DBConnect();
            MySqlConnection conn = dBConnect.ConnectMySQL();

            string query = "insert into Row_Data_Level2_Values(Value,TimeStamp,IdRowDataLevel2) values('"
                + row_Data_Level2_Value.Value + "','"
                + row_Data_Level2_Value.TimeStamp + "','"
                + row_Data_Level2_Value.IdRowDataLevel2 + "');";

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
        public int InsertPG(Row_Data_Level2_Value row_Data_Level2_Value)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "insert into Row_Data_Level2_Values(value,timestamp,id_row_data_level2) values('"
                + row_Data_Level2_Value.Value + "','"
                + row_Data_Level2_Value.TimeStamp + "','"
                + row_Data_Level2_Value.IdRowDataLevel2 + "') RETURNING id;";

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

        public Row_Data_Level2_Value Get_RowDataLevel2Value_By_IDTable_KeyIDLevel1_KeyIDLevel2_TimeStamp(int idTable, string keyIDRowLevel1,string keyIDRowLevel2, double timeStamp)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();

            RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
            Row_Data_Level1 row_Data_Level1 = rowDataLevel1Service.Get_RowDataLevel1_By_IdTable_KeyID(idTable, keyIDRowLevel1);

            RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
            Row_Data_Level2 row_Data_Level2 = rowDataLevel2Service.Get_RowDataLevel2_By_IdRowLevel1_KeyID(row_Data_Level1.Id, keyIDRowLevel2);

            

            string query = "SELECT * FROM row_data_level2_values WHERE " +
                "id_row_data_level2 ='" + row_Data_Level2.Id + "'" +
                " AND timestamp='" + timeStamp + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Row_Data_Level2_Value row_Data_Level = new Row_Data_Level2_Value();

                while (reader.Read())
                {

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.IdRowDataLevel2 = reader.GetInt16(reader.GetOrdinal("id_row_data_level2"));
                    row_Data_Level.Value = reader.GetDouble(reader.GetOrdinal("Value"));
                    row_Data_Level.TimeStamp = timeStamp;
                    
                   
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


        public string Update_Row_Data_Level2_Value(Row_Data_Level2_Value row_Data_Level2_Value)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "UPDATE row_data_level2_values SET value='"
                + row_Data_Level2_Value.Value + "'"
                + " WHERE id='" + row_Data_Level2_Value.Id + "'";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return "success";
            }
            catch (Exception a)
            {
                conn.Close();
                string err = a.Message;
                return err;
            }
            finally
            {
                conn.Close();
            }
            return "failed";

        }

        public List<Row_Data_Level2_Value> Get_All()
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level2_values ORDER BY id ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;

                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level2_Value> list = new List<Row_Data_Level2_Value>();

                while (reader.Read())
                {
                    Row_Data_Level2_Value row_Data_Level = new Row_Data_Level2_Value();
                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.IdRowDataLevel2 = reader.GetInt32(reader.GetOrdinal("id_row_data_level2"));
                    row_Data_Level.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    row_Data_Level.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
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

        public List<Row_Data_Level2_Value> Get_RowDataLevel2Value_By_IDTable_KeyIDLevel1_KeyIDLevel2(int idTable, string keyIDRowLevel1, string keyIDRowLevel2)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();

            RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
            Row_Data_Level1 row_Data_Level1 = rowDataLevel1Service.Get_RowDataLevel1_By_IdTable_KeyID(idTable, keyIDRowLevel1);

            RowDataLevel2Service rowDataLevel2Service = new RowDataLevel2Service();
            Row_Data_Level2 row_Data_Level2 = rowDataLevel2Service.Get_RowDataLevel2_By_IdRowLevel1_KeyID(row_Data_Level1.Id, keyIDRowLevel2);



            string query = "SELECT * FROM row_data_level2_values WHERE " +
                "id_row_data_level2 ='" + row_Data_Level2.Id + "' ORDER BY timestamp DESC ";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level2_Value> list = new List<Row_Data_Level2_Value>();
                while (reader.Read())
                {
                    Row_Data_Level2_Value row_Data_Level = new Row_Data_Level2_Value();

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.IdRowDataLevel2 = reader.GetInt16(reader.GetOrdinal("id_row_data_level2"));
                    row_Data_Level.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    row_Data_Level.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
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

        public string Insert_YTD(List<dynamic> list_YTD, List<Row_Data_Level2_Value> list_Data)
        {
            try
            {
                for (int i = 0; i < list_Data.Count; i++)
                {
                    for (int k = 0; k < list_YTD.Count; k++)
                    {
                        if (!Tool.Check_Exist_List_Data(list_YTD[k].TimeStamp, list_Data.Cast<dynamic>().ToList()))
                        {
                            Row_Data_Level2_Value row_Data_Level2_Value = new Row_Data_Level2_Value();
                            row_Data_Level2_Value.IdRowDataLevel2 = list_Data[i].IdRowDataLevel2;
                            row_Data_Level2_Value.TimeStamp = list_YTD[k].TimeStamp;
                            row_Data_Level2_Value.Value = list_YTD[k].Value;
                            InsertPG(row_Data_Level2_Value);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                return "Lỗi: " + e.Message;
            }
            return "Thêm YTD Thành công";
        }

    }
}
