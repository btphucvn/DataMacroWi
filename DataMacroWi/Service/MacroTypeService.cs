using DataMacroWi.DB;
using DataMacroWi.Model;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.Service
{
    class MacroTypeService
    {
        public List<MacroType> Get_MacroType_By_KeyIDMacro(string keyIDMacro)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM macro_types WHERE key_id_macro='"+keyIDMacro+"' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<MacroType> listMacroType = new List<MacroType>();
                while (reader.Read())
                {
                    MacroType macroType = new MacroType();

                    macroType.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    macroType.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    //AllKeyService allKeyService = new AllKeyService();
                    //AllKey allKey = new AllKey();
                    //allKey = allKeyService.GetAllKeyByKeyID(macroType.KeyID);
                    //macroType.Name = allKey.NameVi;
                    try
                    {
                        macroType.Name = reader.GetString(reader.GetOrdinal("name"));
                    }
                    catch { }
                    try
                    {
                        macroType.IdDetail = reader.GetInt32(reader.GetOrdinal("id_detail"));
                    }
                    catch { }
                    macroType.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    macroType.KeyIDMacro = reader.GetString(reader.GetOrdinal("key_id_macro"));

                    listMacroType.Add(macroType);
                }
                conn.Close();
                return listMacroType;
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

        public MacroType Get_MacroType_By_KeyID(string keyID)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM macro_types WHERE key_id='" + keyID + "' ORDER BY stt ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                MacroType macroType = new MacroType();

                while (reader.Read())
                {

                    macroType.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    macroType.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    //AllKeyService allKeyService = new AllKeyService();
                    //AllKey allKey = new AllKey();
                    //allKey = allKeyService.GetAllKeyByKeyID(macroType.KeyID);
                    //macroType.Name = allKey.NameVi;
                    try
                    {
                        macroType.Name = reader.GetString(reader.GetOrdinal("name"));
                    }
                    catch { }
                    try
                    {
                        macroType.IdDetail = reader.GetInt32(reader.GetOrdinal("id_detail"));
                    }
                    catch { }
                    macroType.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    macroType.KeyIDMacro = reader.GetString(reader.GetOrdinal("key_id_macro"));

                }
                conn.Close();
                return macroType;
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



        public List<MacroType> Get_All_MacroType()
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM macro_types";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<MacroType> listMacroType = new List<MacroType>();
                while (reader.Read())
                {
                    MacroType macroType = new MacroType();

                    macroType.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    macroType.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    macroType.Name = reader.GetString(reader.GetOrdinal("name"));
                    try
                    {
                        macroType.IdDetail = reader.GetInt32(reader.GetOrdinal("id_detail"));
                    }
                    catch { }
                    macroType.Stt = reader.GetInt32(reader.GetOrdinal("stt"));
                    macroType.KeyIDMacro = reader.GetString(reader.GetOrdinal("key_id_macro"));

                    listMacroType.Add(macroType);
                }
                conn.Close();
                return listMacroType;
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

        public string Update(MacroType macroType)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "UPDATE macro_types "
                + "SET  key_id = '" + macroType.KeyID + "', " 
                +" id_detail='"+macroType.IdDetail+"', "
                + " key_id_macro='" + macroType.KeyIDMacro + "', "
                + " stt='" + macroType.Stt + "', "
                + " name='" + macroType.Name + "'"
                + " WHERE id = '" + macroType.Id + "'";

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

        public string Insert(MacroType macroType)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();


            string query = "insert into macro_types(key_id,id_detail,stt,name,key_id_macro) values('"
                            + macroType.KeyID + "','"
                            + macroType.IdDetail + "','"
                             + macroType.Stt + "','"
                            + macroType.Name + "','"
                            + macroType.KeyIDMacro + "') RETURNING id;";

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


        }

        

        public void ClearAllTable(MacroType macroType)
        {
            TableService tableService = new TableService();
            RowService rowService = new RowService();
            RowValueService rowValueService = new RowValueService();
            List<Table> listTable = tableService.Get_Table_By_IDMacroType(macroType.Id);
            foreach(Table table in listTable)
            {
                List<Row> listRow = rowService.Get_Rows_By_IdTable(table.Id);
                foreach(Row row in listRow)
                {
                    rowValueService.Clear(row.ID);

                }
                rowService.Clear(table);
            }
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();


            string query = "DELETE FROM tables WHERE id_macro_type="+macroType.Id;

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception a)
            {
                Form1._Form1.updateTxtBug("Lỗi: " + a.Message);
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
