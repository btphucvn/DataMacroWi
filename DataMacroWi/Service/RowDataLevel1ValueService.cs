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
    class RowDataLevel1ValueService
    {
        public int Insert(Row_Data_Level1_Value row_Data_Level1_Value)
        {
            DBConnect dBConnect = new DBConnect();
            MySqlConnection conn = dBConnect.ConnectMySQL();

            string query = "insert into Row_Data_Level1_Values(Value,TimeStamp,IdRowDataLevel1) values('"
                + row_Data_Level1_Value.Value + "','"
                + row_Data_Level1_Value.TimeStamp + "','"
                + row_Data_Level1_Value.IdRowDataLevel1 + "');";

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
        public int InsertPG(Row_Data_Level1_Value row_Data_Level1_Value)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "insert into Row_Data_Level1_Values(value,timestamp,id_row_data_level1) values('"
                + row_Data_Level1_Value.Value + "','"
                + row_Data_Level1_Value.TimeStamp + "','"
                + row_Data_Level1_Value.IdRowDataLevel1 + "') RETURNING id;";
            
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
    
        public Row_Data_Level1_Value Get_RowDataLevel1Value_By_IDRowDataLevel1_TimeStamp(int IDRowDataLevel1, double timestamp)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level1_values WHERE " +
                "id_row_data_level1 ='" + IDRowDataLevel1 + "'" +
                " AND timestamp='" + timestamp + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Row_Data_Level1_Value row_Data_Level = new Row_Data_Level1_Value();

                while (reader.Read())
                {

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.IdRowDataLevel1 = reader.GetInt32(reader.GetOrdinal("id_row_data_level1"));
                    row_Data_Level.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    row_Data_Level.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));


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

        public List<Row_Data_Level1_Value> Get_RowDataLevel1Value_By_IDRowDataLevel1(int IDRowDataLevel1)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level1_values WHERE " +
                "id_row_data_level1 ='" + IDRowDataLevel1 + "'" +
                " ORDER BY timestamp DESC ";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level1_Value> list_Row_Data_Level1_Value = new List<Row_Data_Level1_Value>();
                while (reader.Read())
                {
                    Row_Data_Level1_Value row_Data_Level = new Row_Data_Level1_Value();

                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.IdRowDataLevel1 = reader.GetInt32(reader.GetOrdinal("id_row_data_level1"));
                    row_Data_Level.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    row_Data_Level.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
                    list_Row_Data_Level1_Value.Add(row_Data_Level);

                }
                conn.Close();
                return list_Row_Data_Level1_Value;
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

        public Row_Data_Level1_Value Get_RowDataLevel1Value_By_IDTable_KeyIDRowDataLevel1_TimeStamp(int idTable
            ,string keyIDRowDataLevel1
            , double timestamp)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
            Row_Data_Level1 rowDataLevel1 = rowDataLevel1Service.Get_RowDataLevel1_By_IdTable_KeyID(idTable,keyIDRowDataLevel1);

            Row_Data_Level1_Value row_Data_Level1_Value = Get_RowDataLevel1Value_By_IDRowDataLevel1_TimeStamp(rowDataLevel1.Id, timestamp);
            return row_Data_Level1_Value;
        }
        public string Update_Row_Data_Level1_Value(Row_Data_Level1_Value row_Data_Level1_Value)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "UPDATE row_data_level1_values SET value='"
                + row_Data_Level1_Value.Value + "'"
                + " WHERE id='" + row_Data_Level1_Value.Id + "'";

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

        public List<Row_Data_Level1_Value> Get_All()
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level1_values ORDER BY id ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;

                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level1_Value> list = new List<Row_Data_Level1_Value>();

                while (reader.Read())
                {
                    Row_Data_Level1_Value row_Data_Level = new Row_Data_Level1_Value();
                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.IdRowDataLevel1 = reader.GetInt32(reader.GetOrdinal("id_row_data_level1"));
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
        
        public string Insert_YTD(List<dynamic> list_YTD, List<Row_Data_Level1_Value> list_Data)
        {
            try
            {
                for (int i = 0; i < list_Data.Count; i++)
                {
                    for (int k = 0; k < list_YTD.Count; k++)
                    {
                        if (!Tool.Check_Exist_List_Data(list_YTD[k].TimeStamp, list_Data.Cast<dynamic>().ToList()))
                        {
                            Row_Data_Level1_Value row_Data_Level1_Value = new Row_Data_Level1_Value();
                            row_Data_Level1_Value.IdRowDataLevel1 = list_Data[i].IdRowDataLevel1;
                            row_Data_Level1_Value.TimeStamp = list_YTD[k].TimeStamp;
                            row_Data_Level1_Value.Value = list_YTD[k].Value;
                            InsertPG(row_Data_Level1_Value);
                        }
                    }
                }
            }
            catch(Exception e)
            {
                return "Lỗi: " + e.Message;
            }
            return "Thêm YTD Thành công";
        }

        public List<Row_Data_Level1_Value> Get_RowDataLevel1Value_By_IDTable_KeyIDRowDataLevel1(int idTable
                    , string keyIDRowDataLevel1)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            RowDataLevel1Service rowDataLevel1Service = new RowDataLevel1Service();
            Row_Data_Level1 rowDataLevel1 = rowDataLevel1Service.Get_RowDataLevel1_By_IdTable_KeyID(idTable, keyIDRowDataLevel1);

            List<Row_Data_Level1_Value> list_Row_Data_Level1_Value = Get_RowDataLevel1Value_By_IDRowDataLevel1(rowDataLevel1.Id);
            return list_Row_Data_Level1_Value;
        }
        
         
    }
}
