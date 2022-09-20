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
    class MacroService
    {
        public List<Macro> GetAllMacro()
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM macros ORDER BY id ASC";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<Macro> listMacro = new List<Macro>();
                while (reader.Read())
                {
                    Macro macro = new Macro();

                    macro.Id = reader.GetInt32(reader.GetOrdinal("id"));
                    macro.KeyID = reader.GetString(reader.GetOrdinal("key_id"));
                    macro.Name = reader.GetString(reader.GetOrdinal("name"));
                    listMacro.Add(macro);
                }
                conn.Close();
                return listMacro;
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
    
        public string Update(Macro macro)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();

            string query = "UPDATE macros SET key_id='"
                + macro.KeyID + "', "
                +" name= '"+macro.Name+"'"
                + " WHERE id='" + macro.Id + "'";
            AllKeyService allKeyService = new AllKeyService();
            AllKey allKey = allKeyService.GetAllKeyByKeyID(macro.KeyID);
            if (allKeyService.CheckExitsAllKeyByKeyID(macro.KeyID))
            {
                allKey.NameVi = macro.Name;
                allKeyService.Update(allKey.KeyID, allKey.NameVi);
            }
            else
            {
                allKey.KeyID = macro.KeyID;
                allKey.NameVi = macro.Name;

            }

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

    }
}
