using DataMacroWi.DB;
using DataMacroWi.Extension;
using DataMacroWi.Model;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Service
{
    class RowValueService
    {
        public int Insert(Row_Value row_value)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            //List<Row_Value> row_Values = GetAll();
            //row_value.ID = Tool.GetLastestID(row_Values.Cast<dynamic>().ToList());
            //string query = "insert into row_value(id,id_row,value,timestamp) values('"
            //    + row_value.ID + "','"
            //    + row_value.ID_Row + "','"
            //    + row_value.Value + "','"
            //    + row_value.TimeStamp + "') RETURNING id;";
            string query = "insert into row_value(id_row,value,timestamp) values('"
                + row_value.ID_Row + "','"
                + row_value.Value + "','"
                + row_value.TimeStamp + "') RETURNING id;";
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

        public int Insert_Update(Row_Value row_value)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            //List<Row_Value> row_Values = GetAll();
            Row_Value row_Value_Check = Get_Row_Value_By_IDRow_TimeStamp(row_value.ID_Row, row_value.TimeStamp);
            string query = "";
            if(row_value.TimeStamp == 1654016400000)
            {
                int test = -1;
            }
            if (row_Value_Check.TimeStamp == 0)
            {
                //row_value.ID = Tool.GetLastestID(row_Values.Cast<dynamic>().ToList());
                query = "insert into row_value(id_row,value,timestamp) values('"
                    + row_value.ID_Row + "','"
                    + row_value.Value + "','"
                    + row_value.TimeStamp + "') RETURNING id;";
            }
            else
            {
                query = "UPDATE row_value SET " +
                    "value='" + row_value.Value + "' " +
                    " WHERE id_row='" + row_value.ID_Row + "' " +
                    " AND timestamp = '"+row_value.TimeStamp+"' " +
                    "RETURNING id";
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


        public List<Row_Value> Get_Row_Value_By_IDRow(int idRow)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_value WHERE id_row='" + idRow + "'  ORDER BY timestamp ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Value> list = new List<Row_Value>();
                while (reader.Read())
                {
                    Row_Value row = new Row_Value();
                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    row.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
                    row.ID_Row = reader.GetInt32(reader.GetOrdinal("id_row"));
                    row.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    list.Add(row);

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

        public List<Row_Value> GetAll()
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_value";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row_Value> list = new List<Row_Value>();
                while (reader.Read())
                {
                    Row_Value row = new Row_Value();
                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    row.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
                    row.ID_Row = reader.GetInt32(reader.GetOrdinal("id_row"));
                    row.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    list.Add(row);

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

        public Row_Value Get_Row_Value_By_IDRow_TimeStamp(int idRow,double timeStamp)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM row_value WHERE id_row='"+ idRow + "' AND timestamp='"+timeStamp+"'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Row_Value row = new Row_Value();
                while (reader.Read())
                {
                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    row.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
                    row.ID_Row = reader.GetInt32(reader.GetOrdinal("id_row"));
                    row.Value = reader.GetDouble(reader.GetOrdinal("value"));

                }
                conn.Close();
                return row;
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

        public void Clear(int idRow)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "DELETE FROM row_value WHERE id_row='"+idRow+"'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;

                command.ExecuteReader();



                conn.Close();
            }
            catch (Exception e)
            {
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
