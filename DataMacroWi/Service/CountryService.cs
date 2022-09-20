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
    class CountryService
    {
        public void Insert(string chau, string nuoc)
        {

            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "insert into countries(continent,country) values('"
                + chau + "','"
                + nuoc + "') RETURNING id;";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);
            try
            {
                conn.Open();

                int id = (int)cmd.ExecuteScalar();

                conn.Close();
            }
            catch (Exception a)
            {
                string err = a.Message;
            }
            finally
            {
                conn.Close();
            }
        }

        public string GetContinent(string country)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM countries WHERE country='" + country + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                string result = "";

                while (reader.Read())
                {
                    result = reader.GetString(reader.GetOrdinal("continent"));
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
            return "";
        }

        public string GetContinent_By_KeyID(string keyid)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM countries WHERE key_id='" + keyid + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                string result = "";

                while (reader.Read())
                {
                    result = reader.GetString(reader.GetOrdinal("continent"));
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
            return "";
        }

        public CountryModel Get_By_Country(string country)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM countries WHERE LOWER(country) like '%" + country + "%'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                string result = "";

                CountryModel contryModel = new CountryModel();

                while (reader.Read())
                {
                    contryModel.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    contryModel.Continent = reader.GetString(reader.GetOrdinal("continent"));

                    contryModel.Country = reader.GetString(reader.GetOrdinal("country"));
                    try
                    {
                        contryModel.Country_Name_Vi = reader.GetString(reader.GetOrdinal("country_name_vi"));
                    }
                    catch { }

                }
                conn.Close();

                return contryModel;
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

        public CountryModel Get_By_Country_Name_Vi(string country)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM countries WHERE country_name_vi='"+country+"'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                string result = "";

                CountryModel contryModel = new CountryModel();

                while (reader.Read())
                {
                    contryModel.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    contryModel.Continent = reader.GetString(reader.GetOrdinal("continent"));

                    contryModel.Country = reader.GetString(reader.GetOrdinal("country"));
                    try
                    {
                        contryModel.Key_ID = reader.GetString(reader.GetOrdinal("key_id"));
                    }
                    catch { }
                    try
                    {
                        contryModel.Country_Name_Vi = reader.GetString(reader.GetOrdinal("country_name_vi"));
                    }
                    catch { }

                }
                conn.Close();

                return contryModel;
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

        public List<CountryModel> Get_All()
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM countries";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                List<CountryModel> list = new List<CountryModel>();
                while (reader.Read())
                {
                    CountryModel row = new CountryModel();

                    row.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    try
                    {
                        row.Country = reader.GetString(reader.GetOrdinal("country"));
                    }
                    catch { }
                    try
                    {
                        row.Key_ID = reader.GetString(reader.GetOrdinal("key_id"));
                    }
                    catch { }
                    try
                    {
                        row.Country_Name_Vi = reader.GetString(reader.GetOrdinal("country_name_vi"));
                    }
                    catch { }
                    try
                    {
                        row.Continent = reader.GetString(reader.GetOrdinal("continent"));
                    }
                    catch { }
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

        public string Update_KeyID(CountryModel country)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "UPDATE countries "
                + "SET key_id = '" + country.Key_ID + "' "
                + "WHERE id = '" + country.ID + "'";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);

            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception a)
            {
                return a.Message;
            }
            finally
            {
                conn.Close();
            }

            return "Thành công";

        }

        public string Update(CountryModel country)
        {
            DBConnect dBConnect = new DBConnect();
            NpgsqlConnection conn = dBConnect.ConnectPG();
            string query = "UPDATE countries "
                + "SET country_name_vi = '" + country.Country_Name_Vi + "', " +
                "key_id='" + country.Key_ID + "' "
                + "WHERE id = '" + country.ID + "'";

            NpgsqlCommand cmd = new NpgsqlCommand(query, conn);

            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception a)
            {
                return a.Message;
            }
            finally
            {
                conn.Close();
            }

            return "Thành công";

        }

        public bool Check_Exist_Country_Name_Vi(string country)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM countries WHERE country_name_vi='" + country + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                bool result = false;

                while (reader.Read())
                {
                    result = true;
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
            return false;
        }

        public CountryModel Get_By_Country_KeyID(string key_id)
        {
            DBConnect connect = new DBConnect();
            NpgsqlConnection conn = connect.ConnectPG();
            string query = "SELECT * FROM countries WHERE key_id='" + key_id + "'";
            try
            {
                conn.Open();
                NpgsqlCommand command = conn.CreateCommand();
                command.CommandText = query;
                command.Connection = conn;
                NpgsqlDataReader reader = command.ExecuteReader();
                string result = "";

                CountryModel contryModel = new CountryModel();

                while (reader.Read())
                {
                    contryModel.ID = reader.GetInt32(reader.GetOrdinal("id"));
                    contryModel.Continent = reader.GetString(reader.GetOrdinal("continent"));

                    contryModel.Country = reader.GetString(reader.GetOrdinal("country"));
                    contryModel.Key_ID = reader.GetString(reader.GetOrdinal("key_id"));

                    try
                    {
                        contryModel.Country_Name_Vi = reader.GetString(reader.GetOrdinal("country_name_vi"));
                    }
                    catch { }

                }
                conn.Close();

                return contryModel;
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
