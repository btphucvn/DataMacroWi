using DataMacroWi.DB;
using DataMacroWi.Model;
using MySql.Data.MySqlClient;
using Npgsql;
using System;
using System.Collections.Generic;

namespace DataMacroWi.Service
{
    class RowDataLevel3ValueService
    {
        public int Insert(Row_Data_Level3_Value row_Data_Level3_Value)
        {
            DBConnect dBConnect = new DBConnect();
            MySqlConnection conn = dBConnect.ConnectMySQL();

            string query = "insert into Row_Data_Level3_Values(Value,TimeStamp,id_row_data_level3) values('"
                + row_Data_Level3_Value.Value + "','"
                + row_Data_Level3_Value.TimeStamp + "','"
                + row_Data_Level3_Value.IdRowDataLevel3 + "');";

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

        public int InsertPG(Row_Data_Level3_Value row_Data_Level3_Value)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "insert into Row_Data_Level3_Values(value,timestamp,id_row_data_level3) values('"
                + row_Data_Level3_Value.Value + "','"
                + row_Data_Level3_Value.TimeStamp + "','"
                + row_Data_Level3_Value.IdRowDataLevel3 + "') RETURNING id;";

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

        public List<Row_Data_Level3_Value> Get_All()
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_data_level3_values ORDER BY id ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;

                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Data_Level3_Value> list = new List<Row_Data_Level3_Value>();

                while (reader.Read())
                {
                    Row_Data_Level3_Value row_Data_Level = new Row_Data_Level3_Value();
                    row_Data_Level.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    row_Data_Level.IdRowDataLevel3 = reader.GetInt32(reader.GetOrdinal("id_row_data_level3"));
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

    }
}
