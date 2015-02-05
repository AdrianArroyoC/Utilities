using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using data = System.Data;
using Oracle.DataAccess.Client;
using FirebirdSql.Data.FirebirdClient;
using System.IO; //
using System.Runtime.InteropServices; //
using System.Diagnostics;
using System.ComponentModel;
using Microsoft.VisualBasic;
using excel = Microsoft.Office.Interop.Excel;
using conf = System.Configuration;

namespace Utilities
{
    public class utils
    {
        public static DialogResult continueBox(string message, string title)
        {
            DialogResult result = MessageBox.Show(message, title, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            return result;
        }

        public static String inputBox(string message, string title, string def)
        {
            string text = Microsoft.VisualBasic.Interaction.InputBox(message, title, def);
            return text;
        }
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

        public static data.DataTable oracleData(string instruction, bool close, OracleConnection conn = null, string[] connectionValues = null) //Oracle select
        {
            data.DataTable dt = new data.DataTable();
            try
            {
                if (conn == null)
                {
                    conn = connectOracle(connectionValues);
                }
                OracleDataAdapter adapter = new OracleDataAdapter();
                adapter.SelectCommand = new OracleCommand(instruction, conn);
                adapter.Fill(dt);
                if (close == true)
                {
                    closeOracle(conn);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return dt;
        }

        public static void oraStatement(string instruction, bool close, OracleConnection conn = null, string[] connectionValues = null) //Oracle update, insert or delete
        {
            try
            {
                if (conn == null)
                {
                    conn = connectOracle(connectionValues);
                }
                OracleCommand cmd = new OracleCommand(instruction, conn);
                cmd.ExecuteNonQuery();
                if (close == true)
                {
                    closeOracle(conn);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        public static data.DataTable fbData(string instruction, bool close, FbConnection conn = null, string[] connectionValues = null) //Firebird select
        {
            data.DataTable dt = new data.DataTable();
            try
            {
                if (conn == null)
                {
                    conn = connectFirebird(connectionValues);    
                }
                FbDataAdapter adapter = new FbDataAdapter();
                adapter.SelectCommand = new FbCommand(instruction, conn);
                adapter.Fill(dt);
                if (close == true)
                {
                    closeFirebird(conn);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return dt;
        }

        public static void fbStatement(string instruction, bool close, FbConnection conn = null, string[] connectionValues = null) //Firebird insert, update or delete
        {
            try
            {
                if (conn == null)
                {
                    conn = connectFirebird(connectionValues);
                }
                FbCommand cmd = new FbCommand(instruction, conn);
                cmd.ExecuteNonQuery();
                if (close == true)
                {
                    closeFirebird(conn);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        public static OracleDataReader oraReader(string instruction, OracleConnection conn = null, string[] connectionValues = null)
        {
            OracleDataReader reader = null;
            try
            {
                if (conn == null)
                {
                    conn = connectOracle(connectionValues);
                }
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

        public static FbDataReader fbReader(string instruction, FbConnection conn = null, string[] connectionValues = null)
        {
            FbDataReader reader = null;
            try
            {
                if (conn == null)
                {
                    conn = connectFirebird(connectionValues);
                }
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
        public static int oraNextId (string instruction, bool close, OracleConnection conn = null, string[] connectionValues = null) 
        {
            int id = 0;
            try
            {
                if (conn == null)
                {
                    conn = connectOracle(connectionValues);
                }
                OracleDataReader reader = oraReader(instruction, conn);
                reader.Read();
                id = Convert.ToInt32(reader.GetValue(0));
                if (close == true)
                {
                    closeOracle(conn);
                }
                id++;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return id;
        }

        public static int fbNextId(string instruction, bool close, string[] connectionValues = null, FbConnection conn = null)
        {
            int id = 0;
            try
            {
                if (conn == null)
                {
                    conn = connectFirebird(connectionValues);
                }
                FbDataReader reader = fbReader(instruction, conn);
                reader.Read();
                id = Convert.ToInt32(reader.GetValue(0));
                if (close == true)
                {
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
        public static String oraNextFolio(string instruction, bool close, OracleConnection conn = null, string[] connectionValues = null)
        {
            string folio = "";
            try
            {
                if (conn == null)
                {
                    conn = connectOracle(connectionValues);
                }
                folio = sortedDt(oracleData(instruction, close, conn));
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return folio;
        }

        public static String fbNextFolio( string instruction, bool close, FbConnection conn = null, string[] connectionValues = null)
        {
            string folio = "";
            try
            {
                if (conn == null)
                {
                    conn = connectFirebird(connectionValues);
                }
                folio = sortedDt(fbData(instruction, close, conn));
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return folio;
        }

        public static string sortedDt (data.DataTable dt)
        {
            string folio = "";
            dt.Columns.Add("folios");
            foreach (data.DataRow dtRow in dt.Rows)
            {
                dtRow["folios"] = Convert.ToInt32(dtRow["folios"]);
            }
            data.DataView view = new data.DataView(dt);
            view.Sort = "folios Desc";
            dt.Clear();
            dt = view.Table;
            folio = dt.Rows[0].ItemArray[0].ToString();
            return folio;
        }
    }

    public class excelWorksheet
    {
        public static object missVal = System.Reflection.Missing.Value;

        public static excel.Application start()
        {
            excel.Application xlApp = new excel.Application();
            return xlApp;
        }

        public static bool verifyExcel(excel.Application xlApp)
        {
            if (xlApp == null)
            {
                MessageBox.Show("Necesitas instalar Excel");
                return false;
            }
            return true;
        }

        public static excel.Workbook createExcel(excel.Application xlApp)
        {
            if (verifyExcel(xlApp))
            {
                excel.Workbook xlWorkBook = xlApp.Workbooks.Add(missVal);
                xlApp.Visible = true;
                return xlWorkBook;
            }
            return null;
        }

        public static excel.Worksheet createWoorkSheet(excel.Application xlApp, excel.Workbook xlWorkBook, data.DataTable dt = null, string[] columns = null, DataGridView dgv = null)
        {
            excel.Worksheet xlWorkSheet = new excel.Worksheet();
            xlWorkSheet = (excel.Worksheet)xlWorkBook.Sheets[1];
            fillExcel(xlWorkSheet, dt, dgv);
            xlWorkSheet.Activate();
            xlWorkBook.Saved = false;
            return xlWorkSheet;
        }

        public static void fillExcel (excel.Worksheet xlWorkSheet, data.DataTable dt = null, DataGridView dgv = null)
        {
            if (dt == null)
            {
                dt = new data.DataTable();
                foreach(DataGridViewColumn column in dgv.Columns)
                {    
                    data.DataColumn col = new data.DataColumn(column.Name);
                    dt.Columns.Add(col);
                }
                foreach(DataGridViewRow row in dgv.Rows)
                {
                    data.DataRow dr = dt.NewRow();
                    for (int i = 0; i < dgv.ColumnCount; i++)
                    {
                        dr[i] = row.Cells[i].Value.ToString();
                    }
                    dt.Rows.Add(dr);
                }
            }
            int c = 0;
            foreach (data.DataColumn column in dt.Columns)
            {
                xlWorkSheet.Cells[0][c] = column.ColumnName.ToString();
                c++;
            }
            for (int i = 1; i <= dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    xlWorkSheet.Cells[i][j] = dt.Rows[i - 1].ItemArray[j].ToString();
                }
            }
        }
    }

    public class config
    {
        public static String[] readAllSettings()
        {
            string[] settings = null;
            var appSettings = conf.ConfigurationManager.AppSettings;
            try
            {
                if (appSettings.Count == 0)
                {
                    MessageBox.Show("Archivo de configuración vacio");
                }
                else
                {
                    settings = new string[conf.ConfigurationManager.AppSettings.Count];
                    int i = 0;
                    foreach (var key in conf.ConfigurationManager.AppSettings.AllKeys)
                    {
                        settings[i] = conf.ConfigurationManager.AppSettings[key].ToString();
                        i++;
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return settings;
        }
        
        public static String[] readSettings(string[] keys)
        {
            var appSettings = conf.ConfigurationManager.AppSettings;
            string[] settings = null;
            try
            {
                if (appSettings.Count == 0)
                {
                    MessageBox.Show("Archivo de configuración vacio");
                }
                else
                {
                    settings = new string[keys.Length];
                    int j = 0;
                    foreach (var key in appSettings.AllKeys)
                    {
                        for (int i = 0; i < keys.Length; i++)
                        {
                            if (keys[i] == key.ToString())
                            {
                                settings[j] = appSettings[key].ToString();
                                j++;
                            }
                        }
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return settings;
        }

        public static String readSetting(string key)
        {
            var appSettings = conf.ConfigurationManager.AppSettings;
            string setting = "";
            try
            {
                setting = appSettings[key] ?? "";
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return setting;
        }

        public static void addUpdateSetting(string key, string value)
        {
            var configFile = conf.ConfigurationManager.OpenExeConfiguration(conf.ConfigurationUserLevel.None);
            var settigns = configFile.AppSettings.Settings;
            try
            {
                if (settigns[key] == null)
                {
                    settigns.Add(key, value);
                }
                else
                {
                    settigns[key].Value = value;
                }
                configFile.Save(conf.ConfigurationSaveMode.Modified);
                conf.ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }
    }
}
