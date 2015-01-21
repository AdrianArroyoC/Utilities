using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using Oracle.DataAccess.Client;
using FirebirdSql.Data.FirebirdClient;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Utilities
{
    public class utils
    {
        public static Boolean confirmationBox(string message, string title)
        {
            bool confirm = false;
            DialogResult result = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Yes)
            {
                confirm = true;
            }
            return confirm;
        }
        //
    }

    public class dataBase
    {
        /*Parameters for te connection if db == true { serv = oracle service name; dir = oracle host (ip) } 
            else {serv = firebird datasource (server); dir = firebird database (file)}*/
        //public string[] connectionValues = new string[4] { "user", "pass", "serv", "dir" };

        public static String connectionString(string[] connectionValues, bool db = false)
        {
            string sql = "";
            if (db == true)
            {
                sql = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + connectionValues[3] + ")(PORT=1521))(CONNECT_DATA=(SERVICE_NAME="
                    + connectionValues[2] + ")));User Id=" + connectionValues[0] + ";Password=" + connectionValues[1];
            }
            else
            {
                sql = "user=" + connectionValues[0] + "; password=" + connectionValues[1] + "; database=" + connectionValues[3] + "; datasource= " + connectionValues[2] + ";";
            }
            return sql;
        }

        public static OracleConnection connectOracle(string[] connectionValues)
        {
            try
            {
                string oracleString = connectionString(connectionValues, true);
                OracleConnection remoteConnection = new OracleConnection(oracleString);
                remoteConnection.Open();
                return remoteConnection;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return null;
            }
        }

        public static void closeOracle(OracleConnection connection)
        {
            try
            {
                if (connection != null)
                {
                    connection.Close();
                    OracleConnection.ClearPool(connection);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        public static FbConnection connectFirebird(string[] connectionValues)
        {
            try
            {
                string firebirdString = connectionString(connectionValues);
                FbConnection remoteConnection = new FbConnection(firebirdString);
                remoteConnection.Open();
                return remoteConnection;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return null;
            }
        }

        public static void closeFirebird(FbConnection connection)
        {
            try
            {
                if (connection != null)
                {
                    connection.Close();
                    FbConnection.ClearPool(connection);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        public static DataTable oracleData(string[] connectionValues, string instruction) //Oracle select
        {
            DataTable dt = new DataTable();
            try
            {
                OracleConnection conn = connectOracle(connectionValues);
                OracleDataAdapter adapter = new OracleDataAdapter();
                adapter.SelectCommand = new OracleCommand(instruction, conn);
                adapter.Fill(dt);
                closeOracle(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return dt;
        }

        public static void oraStatement(string instruction, string[] connectionValues = null, OracleConnection conn = null) //Oracle update, insert or delete
        {
            try
            {
                if (conn == null)
                {
                    conn = connectOracle(connectionValues);
                }
                OracleCommand cmd = new OracleCommand(instruction, conn);
                cmd.ExecuteNonQuery();
                closeOracle(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        public static DataTable fbData(string[] connectionValues, string instruction) //Firebird select
        {
            DataTable dt = new DataTable();
            try
            {
                FbConnection conn = connectFirebird(connectionValues);
                FbDataAdapter adapter = new FbDataAdapter();
                adapter.SelectCommand = new FbCommand(instruction, conn);
                adapter.Fill(dt);
                closeFirebird(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return dt;
        }

        public static void fbStatement(string instruction, string[] connectionValues = null, FbConnection conn = null) //Firebird insert, update or delete
        {
            try
            {
                if (conn == null)
                {
                    conn = connectFirebird(connectionValues);
                }
                FbCommand cmd = new FbCommand(instruction, conn);
                cmd.ExecuteNonQuery();
                closeFirebird(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        public static OracleDataReader oraReader(OracleConnection conn, string instruction)
        {
            OracleDataReader reader = null;
            try
            {
                OracleCommand cmd = conn.CreateCommand();
                cmd.CommandText = instruction;
                reader = cmd.ExecuteReader();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return (reader);
        }

        public static FbDataReader fbReader(FbConnection conn, string instruction)
        {
            FbDataReader reader = null;
            try
            {
                FbCommand cmd = conn.CreateCommand();
                cmd.CommandText = instruction;
                reader = cmd.ExecuteReader();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return (reader);
        }
    }

    public class nextsNumbers : dataBase
    {
        //Get the next number by a sql query
        public static Int32 nextId (string[] connectionValues, string instruction, bool db = false) 
        {
            Int32 id = 0;
            try
            {
                if (db == true)
                {
                    OracleConnection conn = connectOracle(connectionValues);
                    OracleDataReader reader = oraReader(conn, instruction);
                    reader.Read();
                    id = reader.GetInt32(0);
                    closeOracle(conn);
                }
                else
                {
                    FbConnection conn = connectFirebird(connectionValues);
                    FbDataReader reader = fbReader(conn, instruction);
                    reader.Read();
                    id = reader.GetInt32(0);
                    closeFirebird(conn);
                }
                id++;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return id;
        }

        //When the sql querys returns string and you want to know the next number 
        public static String nextFolio(string[] connectionValues, string instruction, bool db = false)
        {
            string folio = "";
            DataTable dt = null;
            try
            {
                if (db == true)
                {
                    dt = oracleData(connectionValues, instruction);
                }
                else
                {
                    dt = fbData(connectionValues, instruction);
                }
                dt.Columns.Add("folios");
                foreach (DataRow dtRow in dt.Rows)
                {
                    dtRow["folios"] = Convert.ToInt32(dtRow["folios"]);
                }
                DataView view = new DataView(dt);
                view.Sort = "folios Desc";
                dt.Clear();
                dt = view.Table;
                folio = dt.Rows[0].ToString();
            }
            catch (Exception error)

            {
                MessageBox.Show(error.Message);
            }
            return folio;
        }
    }

    public class logs
    {
        public static void Log(string file, string line = "", string path = "", int lines = 0, string[] text = null)
        {
            StreamWriter w = File.AppendText(@path + @file);
            w.Write("\r\nLog : ");
            w.WriteLine(DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
            if (lines != 0 && text != null)
            {
                for (int i = 0; i < lines; i++)
                {
                    w.WriteLine(" [Información] :-->" + text[i]);
                }
            }
            else
            {
                w.WriteLine(" [Información] :-->" + line);
            }
            w.WriteLine ("-------------------------------");
        }
    }

    private class fileIni
    {
        [DllImport("kernel32")]
        public static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
    }

    public class configIni : fileIni
    {
        public static String readValue (string file, string section, string key, string path = "")
        {
            string value = "";
            StringBuilder cantidad = new StringBuilder();
            if (File.Exists(@path + @file))
            {
                GetPrivateProfileString(section,
                                             key,
                                             "",
                                             cantidad,
                                             cantidad.Capacity,
                                             file);
                value = cantidad.ToString();
            }
            return value;
        }
        
    }
}
