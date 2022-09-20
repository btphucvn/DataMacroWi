using MySql.Data.MySqlClient;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataMacroWi.DB
{
    class DBConnect
    {
        public MySqlConnection ConnectMySQL()
        {

            string host = "localhost";
            int port = 3307;
            string database = "dbvnmacro";
            string username = "root";
            string password = "";


            String connString = "Server=" + host + ";Database=" + database
                + ";port=" + port + ";User Id=" + username + ";password=" + password;

            MySqlConnection conn = new MySqlConnection(connString);

            return conn;
        }
        public NpgsqlConnection ConnectPG()
        {

            string host = "127.0.0.1";
            int port = 3308;
            string database = "dbvnmacro";
            string username = "postgres";
            string password = "Thanhphuc123@";


            String connString =  "Host="+ host + ":"+ port + ";" +
                    "Username="+ username + ";" +
                    "Password="+ password + ";" +
                    "Database="+ database;

            NpgsqlConnection  connection = new NpgsqlConnection(connString);

            return connection;
        }
    }
}
