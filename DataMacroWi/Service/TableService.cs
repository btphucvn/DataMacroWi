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
    class TableService
    {
        public int Insert(Table table)
        {
            DBConnect dBConnect = new DBConnect();
            MySqlConnection conn = dBConnect.ConnectMySQL();

            string query = "insert into tables(keyID,valueType,dateType,stt,idMacroType) values('"
                + Tool.titleToKeyID(table.KeyID) + "','"
                + table.ValueType + "','"
                + table.DateType + "','"
                + table.Stt + "','"
                + table.IdMacroType + "');";

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
        public int InsertPG(Table table)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "insert into Tables(key_id,value_type,date_type,stt,id_macro_type,table_type,name,unit) values('"
                + Tool.titleToKeyID(table.KeyID) + "','"
                + table.ValueType + "','"
                + table.DateType + "','"
                + table.Stt + "','"
                + table.IdMacroType + "','"
                + table.TableType + "','"
                + table.Name + "','"
                + table.Unit + "') RETURNING id;";

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

        public List<Table> Get_Table_By_IDMacroType(int id_macro_type)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM tables WHERE id_macro_type='" + id_macro_type + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Table> listTable = new List<Table>();
                while (reader.Read())
                {
                    Table table = new Table();

                    table.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    table.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    table.IdMacroType= reader.GetInt32(reader.GetOrdinal("id_macro_type"));
                    table.ValueType= reader.GetString(reader.GetOrdinal("value_type"));
                    table.DateType = reader.GetString(reader.GetOrdinal("date_type"));
                    table.Unit = reader.GetString(reader.GetOrdinal("unit"));
                    try { table.TableType = reader.GetString(reader.GetOrdinal("table_type")); } catch { }

                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(table.KeyID);
                    table.Name = allKey.NameVi;

                    table.Stt = reader.GetInt32(reader.GetOrdinal("stt"));

                    listTable.Add(table);
                }
                conn.Close();
                return listTable;
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

        public List<Table> Get_Table_By_IDTable(int id_macro_type)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM tables WHERE id='" + id_macro_type + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Table> listTable = new List<Table>();
                while (reader.Read())
                {
                    Table table = new Table();

                    table.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    table.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    table.IdMacroType = reader.GetInt32(reader.GetOrdinal("id_macro_type"));
                    table.ValueType = reader.GetString(reader.GetOrdinal("value_type"));
                    table.DateType = reader.GetString(reader.GetOrdinal("date_type"));
                    table.Unit = reader.GetString(reader.GetOrdinal("unit"));
                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(table.KeyID);
                    table.Name = allKey.NameVi;

                    table.Stt = reader.GetInt32(reader.GetOrdinal("stt"));

                    listTable.Add(table);
                }
                conn.Close();
                return listTable;
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

        public Table Get_Table_By_KeyID_ValueType(string keyIDTable,string valueType)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM tables WHERE key_id='" + keyIDTable + "' AND value_type='"+ valueType + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Table table = new Table();

                while (reader.Read())
                {

                    table.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    table.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    table.IdMacroType = reader.GetInt32(reader.GetOrdinal("id_macro_type"));
                    table.ValueType = reader.GetString(reader.GetOrdinal("value_type"));
                    table.DateType = reader.GetString(reader.GetOrdinal("date_type"));
                    table.Unit = reader.GetString(reader.GetOrdinal("unit"));
                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(table.KeyID);
                    table.Name = allKey.NameVi;

                    table.Stt = reader.GetInt32(reader.GetOrdinal("stt"));

                    return table;
                }
                conn.Close();
                return table;
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

        public Table Get_Table_By_KeyID_TableType(string keyIDTable, string tableType)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM tables WHERE key_id='" + keyIDTable + "' AND table_type='" + tableType + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Table table = new Table();

                while (reader.Read())
                {

                    table.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    table.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    table.IdMacroType = reader.GetInt32(reader.GetOrdinal("id_macro_type"));
                    table.ValueType = reader.GetString(reader.GetOrdinal("value_type"));
                    table.DateType = reader.GetString(reader.GetOrdinal("date_type"));
                    table.Unit = reader.GetString(reader.GetOrdinal("unit"));
                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(table.KeyID);
                    table.Name = allKey.NameVi;

                    table.Stt = reader.GetInt32(reader.GetOrdinal("stt"));

                    return table;
                }
                conn.Close();
                return table;
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

        public Table Get_Table_By_KeyID_TableType_ValueType(string keyIDTable, string tableType, string valueType)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM tables WHERE key_id='" + keyIDTable 
                + "' AND table_type='" + tableType + "'"
                + " AND value_type='" + valueType+"'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Table table = new Table();

                while (reader.Read())
                {

                    table.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    table.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    table.IdMacroType = reader.GetInt32(reader.GetOrdinal("id_macro_type"));
                    table.ValueType = reader.GetString(reader.GetOrdinal("value_type"));
                    table.DateType = reader.GetString(reader.GetOrdinal("date_type"));
                    table.Unit = reader.GetString(reader.GetOrdinal("unit"));
                    table.Name = reader.GetString(reader.GetOrdinal("name"));
                    table.TableType = reader.GetString(reader.GetOrdinal("table_type"));

                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(table.KeyID);
                    table.Name = allKey.NameVi;

                    table.Stt = reader.GetInt32(reader.GetOrdinal("stt"));

                    return table;
                }
                conn.Close();
                return table;
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
        public Table Get_Table_By_KeyID_TableType_ValueType_KeyIDMacroType(string keyIDTable, string tableType, string valueType,string keyIDMacroType)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();

            MacroTypeService macroTypeService = new MacroTypeService();
            MacroType macroType = macroTypeService.Get_MacroType_By_KeyID(keyIDMacroType);

            string query = "SELECT * FROM tables WHERE key_id='" + keyIDTable
                + "' AND table_type='" + tableType + "'"
                + " AND id_macro_type='" + macroType.Id + "'"
                + " AND value_type='" + valueType + "'";

            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                Table table = new Table();

                while (reader.Read())
                {

                    table.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    table.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    table.IdMacroType = reader.GetInt32(reader.GetOrdinal("id_macro_type"));
                    table.ValueType = reader.GetString(reader.GetOrdinal("value_type"));
                    table.DateType = reader.GetString(reader.GetOrdinal("date_type"));
                    table.Unit = reader.GetString(reader.GetOrdinal("unit"));
                    table.Name = reader.GetString(reader.GetOrdinal("name"));
                    table.TableType = reader.GetString(reader.GetOrdinal("table_type"));

                    AllKeyService allKeyService = new AllKeyService();
                    AllKey allKey = new AllKey();
                    allKey = allKeyService.GetAllKeyByKeyID(table.KeyID);
                    table.Name = allKey.NameVi;

                    table.Stt = reader.GetInt32(reader.GetOrdinal("stt"));

                    return table;
                }
                conn.Close();
                return table;
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

        public string Update(Table table)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "UPDATE tables "
                + "SET  key_id = '" + table.KeyID + "', "
                + " value_type='" + table.ValueType + "', "
                + " date_type='" + table.DateType + "', "
                + " stt='" + table.Stt + "', "
                + " id_macro_type='" + table.IdMacroType + "', "
                + " unit='" + table.Unit + "', "
                + " table_type='" + table.TableType + "', "
                + " name='" + table.Name + "' "
                + " WHERE id = '" + table.Id + "'";

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

        public void Delete(Table table)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "DELETE FROM tables WHERE id="+table.Id;
            RowService rowService = new RowService();
            rowService.Clear(table);
            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);

            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception a)
            {
                Form1._Form1.updateTxtBug("Xóa table thất bại, lỗi: "+a.Message);
            }
            finally
            {
                conn.Close();
            }
        }
        
    }
}
