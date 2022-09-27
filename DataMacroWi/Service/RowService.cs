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
    class RowService
    {
        public int Insert(Row row)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "insert into rows(key_id,level,unit,stt,id_string,name,yaxis,id_table) values('"
                + row.Key_ID + "','"
                + row.Level + "','"
                + row.Unit + "','"
                + row.Stt + "','"
                + row.ID_String + "','"
                + row.Name + "','"
                + row.YAxis + "','"
                + row.ID_Table + "') RETURNING id;";

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

        public int Insert_And_Update_STT(Row row)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            row.Name = row.Name.Replace("'", "\'");
            string query = "insert into rows(key_id,level,unit,stt,id_string,name,yaxis,id_table) values('"
                + row.Key_ID + "','"
                + row.Level + "','"
                + row.Unit + "','"
                + row.Stt + "','"
                + row.ID_String + "','"
                + row.Name + "','"
                + row.YAxis + "','"
                + row.ID_Table + "') RETURNING id;";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            try
            {
                conn.Open();
                int id = (int)cmd.ExecuteScalar();
                List<Row> list = Get_Rows_By_IdTable(row.ID_Table);
                for(int i = 0; i < list.Count; i++)
                {
                    Row row_update = list[i];
                    row_update.Stt = i;
                    Update(row_update);
                }
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


        public List<Row> Get_Rows_By_IdTable(int idTable)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM rows WHERE id_table='" + idTable + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row> list = new List<Row>();
                while (reader.Read())
                {
                    Row row = new Row();

                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    row.Key_ID = reader.GetString(reader.GetOrdinal("key_id"));
                    row.ID_Table = reader.GetInt32(reader.GetOrdinal("id_table"));
                    try
                    {
                        row.Level = reader.GetInt32(reader.GetOrdinal("level"));
                    }
                    catch { }
                    row.ID_String = reader.GetString(reader.GetOrdinal("id_string"));
                    row.Name = reader.GetString(reader.GetOrdinal("name"));
                    try
                    {
                        row.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    }
                    catch { }
                    row.Unit = reader.GetString(reader.GetOrdinal("unit"));

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

        public Row Get_Row_By_KeyID_Unit_IDTable(string keyID,string unit, int idTable)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM rows WHERE " +
                " unit='"+unit+"' AND "+
                " id_table='" + idTable + "' AND " +
                " key_id='" + keyID + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Row row = new Row();

                while (reader.Read())
                {

                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    row.Key_ID = reader.GetString(reader.GetOrdinal("key_id"));
                    row.ID_Table = reader.GetInt32(reader.GetOrdinal("id_table"));
                    try
                    {
                        row.Level = reader.GetInt32(reader.GetOrdinal("level"));
                    }
                    catch { }
                    row.ID_String = reader.GetString(reader.GetOrdinal("id_string"));
                    row.Name = reader.GetString(reader.GetOrdinal("name"));
                    try
                    {
                        row.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    }
                    catch { }
                    row.Unit = reader.GetString(reader.GetOrdinal("unit"));

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

        public List<Row> Get_Row_By_ContainKeyID_Unit_IDTable(string containContinientKeyID, string unit, int idTable)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "";
            query = "SELECT * FROM rows WHERE " +
                " unit='" + unit + "' AND " +
                " id_table='" + idTable + "' AND " +
                " key_id like '%\\_" + containContinientKeyID + "\\_%' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Row> list = new List<Row>();
                Row row_chau_luc = Get_Row_By_KeyID_Unit_IDTable("xuat-khau_Value__" + containContinientKeyID, "Triệu USD", idTable);
                list.Add(row_chau_luc);
                while (reader.Read())
                {
                    Row row = new Row();
                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    row.Key_ID = reader.GetString(reader.GetOrdinal("key_id"));
                    row.ID_Table = reader.GetInt32(reader.GetOrdinal("id_table"));
                    try
                    {
                        row.Level = reader.GetInt32(reader.GetOrdinal("level"));
                    }
                    catch { }
                    row.ID_String = reader.GetString(reader.GetOrdinal("id_string"));
                    row.Name = reader.GetString(reader.GetOrdinal("name"));
                    try
                    {
                        row.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    }
                    catch { }
                    row.Unit = reader.GetString(reader.GetOrdinal("unit"));
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


        public Row Get_Row_By_KeyID(string keyID)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM rows WHERE " +
                " key_id='" + keyID + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Row row = new Row();

                while (reader.Read())
                {

                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    row.Key_ID = reader.GetString(reader.GetOrdinal("key_id"));
                    row.ID_Table = reader.GetInt32(reader.GetOrdinal("id_table"));
                    try
                    {
                        row.Level = reader.GetInt32(reader.GetOrdinal("level"));
                    }
                    catch { }
                    row.ID_String = reader.GetString(reader.GetOrdinal("id_string"));
                    row.Name = reader.GetString(reader.GetOrdinal("name"));
                    try
                    {
                        row.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    }
                    catch { }
                    row.Unit = reader.GetString(reader.GetOrdinal("unit"));

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

        public void Clear(Table table)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "DELETE FROM rows WHERE id_table='" + table.Id + "'";
            List<Row> listRow = Get_Rows_By_IdTable(table.Id);
            RowValueService rowValueService = new RowValueService();
            foreach(var row in listRow)
            {
                rowValueService.Clear(row.ID);
            }

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

        public string Update(Row row)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "UPDATE rows "
                + "SET  key_id = '" + row.Key_ID + "', "
                + " level='" + row.Level + "', "
                + " unit='" + row.Unit + "', "
                + " stt='" + row.Stt + "', "
                + " id_table='" + row.ID_Table + "', "
                + " id_string='" + row.ID_String + "', "
                + " yaxis='" + row.YAxis + "', "
                + " name='" + row.Name + "' "

                + " WHERE id = '" + row.ID + "'";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);

            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return "Thành công";
            }
            catch (Exception a)
            {
                return a.Message;
            }
            finally
            {
                conn.Close();
            }

            return "Thất bại";

        }

        public string Delete(Row row)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "DELETE FROM rows where id ="+row.ID;
            string query_rowValue = "DELETE FROM row_value where id_row =" + row.ID;

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            NpgsqlCommand cmd_value = new NpgsqlCommand(query_rowValue, conn);

            try
            {
                conn.Open();
                cmd_value.ExecuteNonQuery();
                cmd.ExecuteNonQuery();

                conn.Close();
                return "Thành công";
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


        public List<Row_Value> Sum_Contain_KeyID_By_TimeStamp(int idTable, string keyID)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();

            string query = "WITH rows_table AS " +
                "(" +
                "SELECT * FROM rows WHERE id_table="+idTable+" AND key_id like '%"+ keyID + "%'" +
                "), row_value_table AS " +
                "( " +
                "SELECT * FROM rows_table INNER JOIN row_value ON rows_table.id=row_value.id_row " +
                ") " +
                "SELECT timestamp, SUM(row_value_table.value) as value " +
                "FROM row_value_table " +
                "GROUP BY row_value_table.timestamp " +
                "ORDER BY timestamp DESC";
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
                    Row_Value row_value = new Row_Value();

                    row_value.TimeStamp = reader.GetDouble(reader.GetOrdinal("timestamp"));
                    row_value.Value = reader.GetDouble(reader.GetOrdinal("value"));
                    list.Add(row_value);
                }
                conn.Close();
                return list;
            }
            catch (Exception e)
            {
                Form1._Form1.updateTxtBug("Lỗi: " + e.Message);
                conn.Close();
            }
            finally
            {
                conn.Close();
            }
            return null;
        }


        public double Sum_Contain_KeyID(string keyID, int idTable,double timeStamp)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "WITH row_table AS( " +
                " SELECT * FROM rows " +
                " WHERE key_id LIKE '%\\_" + keyID + "\\_%' AND" +
                " id_table=" + idTable + " " +
                "),row_table_value AS( " +
                "SELECT * FROM row_table INNER JOIN row_value ON row_table.id = row_value.id_row WHERE timestamp="+timeStamp+") " +
                "SELECT SUM(row_table_value.value) " +
                "FROM row_table_value";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Row row = new Row();
                double result = double.NaN;
                while (reader.Read())
                {

                     result= reader.GetDouble(reader.GetOrdinal("sum"));


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
            return double.NaN;
        }

    }
}
