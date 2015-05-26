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
        
        public static String openPath(string[] filters)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            string filter = "", path = "";
            foreach (string i in filters)
                filter += i + " files (*." + i + ")|*." + i + "|";
            openFileDialog1.Filter = filter + "All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.Multiselect = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                path = openFileDialog1.FileName;
            return path;
        }

        public static String savePath(string[] filters)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            string filter = "", path = "";
            foreach (string i in filters)
                filter += i + " files (*." + i + ")|*." + i +"|"; 
            saveDialog.Filter = filter + "All files (*.*)|*.*"; 
            saveDialog.FilterIndex = 0;
            saveDialog.RestoreDirectory = true;
            if(saveDialog.ShowDialog() == DialogResult.OK)
                path = Path.GetFullPath(saveDialog.FileName); 
            return path;
        }

        public static String inputBox(string message, string title, string def)
        {
            return Microsoft.VisualBasic.Interaction.InputBox(message, title, def);
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
                sql = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + connectionValues[3] + ")(PORT=1521))(CONNECT_DATA=(SERVICE_NAME="
                    + connectionValues[2] + ")));User Id=" + connectionValues[0] + ";Password=" + connectionValues[1];
            else
                sql = "user=" + connectionValues[0] + "; password=" + connectionValues[1] + "; database=" + connectionValues[3] + "; datasource= " + connectionValues[2] + ";";
                //sql = "user=" + connectionValues[0] + "; password=" + connectionValues[1] + "; database=" + connectionValues[3] + "; datasource= " + connectionValues[2] + 
                    //"port=3050; dialect=3; charset=NONE; role=; connection lifetime=15; pooling=false; packet Size=8192; serverType=0";
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

        public static data.DataTable oracleData(string instruction, OracleConnection conn = null, string[] connectionValues = null) //Oracle select
        {
            data.DataTable dt = new data.DataTable();
            try
            {
                if (conn == null)
                    conn = connectOracle(connectionValues);
                OracleDataAdapter adapter = new OracleDataAdapter();
                adapter.SelectCommand = new OracleCommand(instruction, conn);
                adapter.Fill(dt);
                if (connectionValues != null)
                    closeOracle(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return dt;
        }

        public static void oraStatement(string instruction, OracleConnection conn = null, string[] connectionValues = null) //Oracle update, insert or delete
        {
            try
            {
                if (conn == null)
                    conn = connectOracle(connectionValues);
                OracleCommand cmd = new OracleCommand(instruction, conn);
                cmd.ExecuteNonQuery();
                if (connectionValues != null)
                    closeOracle(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        public static data.DataTable fbData(string instruction, FbConnection conn = null, string[] connectionValues = null) //Firebird select
        {
            data.DataTable dt = new data.DataTable();
            try
            {
                if (conn == null)
                    conn = connectFirebird(connectionValues);    
                FbDataAdapter adapter = new FbDataAdapter();
                adapter.SelectCommand = new FbCommand(instruction, conn);
                adapter.Fill(dt);
                if (connectionValues != null)
                    closeFirebird(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return dt;
        }

        public static void fbStatement(string instruction, FbConnection conn = null, string[] connectionValues = null) //Firebird insert, update or delete
        {
            try
            {
                if (conn == null)

                    conn = connectFirebird(connectionValues);
                FbCommand cmd = new FbCommand(instruction, conn);
                cmd.ExecuteNonQuery();
                if (connectionValues != null)
                    closeFirebird(conn);
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
                    conn = connectOracle(connectionValues);
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
                    conn = connectFirebird(connectionValues);
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
        public static int oraNextId (string field, string table, OracleConnection conn = null, string[] connectionValues = null, string conditions = null) 
        {
            string instruction = "select nvl(max(" + field + "),0) from " + table;
            if (conditions != null)
                instruction += " where " + conditions;
            int id = 0;
            try
            {
                if (conn == null)
                    conn = connectOracle(connectionValues);
                OracleDataReader reader = oraReader(instruction, conn);
                reader.Read();
                //if (reader.GetValue(0) == null)
                    //id = 0;
                //else
                    id = Convert.ToInt32(reader.GetValue(0));
                if (connectionValues != null)
                    closeOracle(conn);
                id++;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return id;
        }

        public static int fbNextId(string field, string table, FbConnection conn = null, string[] connectionValues = null, string conditions = null)
        {
            string instruction = "select max(" + field + ") from " + table;
            if (conditions != null)
                instruction += " where " + conditions;
            int id = 0;
            try
            {
                if (conn == null)
                    conn = connectFirebird(connectionValues);
                FbDataReader reader = fbReader(instruction.Replace("max", "count"), conn); //
                reader.Read();
                if (Convert.ToInt32(reader.GetValue(0)) == 0) //
                    id = 0;
                else ////
                {
                    reader = fbReader(instruction, conn);
                    reader.Read();
                    id = Convert.ToInt32(reader.GetValue(0));
                }
                if (connectionValues != null)
                    closeFirebird(conn);
                id++;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return id;
        }
        //When the sql querys returns string and you want to know the next number 
        public static String oraNextFolio(string field, string field2, string table, OracleConnection conn = null, string[] connectionValues = null, string conditions = null)
        {
            string instruction = "select nvl(" + field + ",0) from " + table; //
            if (conditions != null)
                instruction += " where " + conditions;
            int folio1 = 0, folio2 = 0, folio = 0;
            try
            {
                if (conn == null)
                    conn = connectOracle(connectionValues);
                folio1 = sortedDt(oracleData(instruction, conn));
                if (folio1 != 0) //
                {
                    data.DataTable dt = oracleData((instruction + " and " + field2 + " = " + (oraNextId(field2, table, conn, null, conditions) - 1).ToString()), conn);
                    folio2 = Convert.ToInt32(dt.Rows[0].ItemArray[0]);
                }
                if (folio1 >= folio2)
                    folio = folio1;
                else
                    folio = folio2;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return (folio + 1).ToString();
        }

        public static String fbNextFolio(string field1, string field2, string table, FbConnection conn = null, string[] connectionValues = null, string conditions = null)
        {
            string instruction = "select " + field1 + " from " + table, instruction2 = "select count(" + field1 + ") from " + table; //
            if (conditions != null)
            {
                instruction += " where " + conditions;
                instruction2 += " where " + conditions; //
            }
            int folio1 = 0, folio2 = 0, folio = 0;
            try
            {
                if (conn == null)
                    conn = connectFirebird(connectionValues);
                FbDataReader reader = fbReader(instruction2, conn); //
                reader.Read();
                if (Convert.ToInt32(reader.GetValue(0)) == 0) //
                    folio = 0;
                else //
                {
                    folio1 = sortedDt(fbData(instruction, conn));
                    data.DataTable dt = fbData((instruction + " and " + field2 + " = " + (fbNextId(field2, table, conn, null, conditions) - 1).ToString()), conn);
                    folio2 = Convert.ToInt32(dt.Rows[0].ItemArray[0]);
                    if (folio1 >= folio2)
                        folio = folio1;
                    else
                        folio = folio2;
                }
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return (folio + 1).ToString();
        }

        public static int sortedDt(data.DataTable dt)
        {
            int folio;
            if (dt == null)
                folio = 0;
            else
            {
                try
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                        dt.Rows[i].ItemArray[0] = Convert.ToInt32(dt.Rows[i].ItemArray[0].ToString());
                    data.DataView view = dt.DefaultView;
                    view.Sort = dt.Columns[0].ColumnName + " desc";
                    dt = view.ToTable();
                    folio = Convert.ToInt32(dt.Rows[0].ItemArray[0].ToString().TrimStart('0'));
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                    folio = 0;
                }
            }            
            return folio;
        }
    }

    public class excelWorksheet
    {
        public static object missVal = System.Reflection.Missing.Value;

        public static excel.Application start()
        {
            try
            {
                excel.Application xlApp = new excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("Necesitas instalar Excel");
                    return null;
                }
                return xlApp;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return null;
            }
        }

        public static excel.Workbook createExcel()
        {
            excel.Application xlApp = start();
            try
            {
                excel.Workbook xlWorkBook = xlApp.Workbooks.Add(missVal);
                xlApp.Visible = true;
                return xlWorkBook;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return null;
            }
        }

        public static void readWoorkSheet(DataGridView dgv, string path)
        {
            data.DataTable dt = new data.DataTable();
            excel.Application xlApp = start();
            try
            {
                excel.Workbook xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                excel.Worksheet xlWorkSheet = (excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                excel.Range range = xlWorkSheet.UsedRange;
                for (int i = 1; i <= range.Columns.Count; i++)
                    dt.Columns.Add(range.Cells[1, i].Value2);
                for (int i = 2; i <= range.Rows.Count; i++)
                {
                    data.DataRow dr = dt.NewRow();
                    for (int j = 1; j <= range.Columns.Count; j++)
                        dr[j - 1] = ((range.Cells[i, j] as excel.Range).Value2).ToString();
                    dt.Rows.Add(dr);
                }
                dgv.DataSource = dt;
                xlWorkBook.Close(false, null, null);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }   
        }

        public static excel.Worksheet createWoorkSheet(data.DataTable dt = null, DataGridView dgv = null)
        {
            excel.Workbook xlWorkBook = createExcel();
            excel.Worksheet xlWorkSheet = new excel.Worksheet();
            try
            {
                if (xlWorkBook != null)
                {
                    xlWorkSheet = (excel.Worksheet)xlWorkBook.Sheets[1];
                    fillExcel(xlWorkSheet, dt, dgv);
                    xlWorkSheet.Activate();
                    xlWorkBook.Saved = false;
                }
                return xlWorkSheet;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return null;
            }
        }

        public static void fillExcel (excel.Worksheet xlWorkSheet, data.DataTable dt = null, DataGridView dgv = null)
        {
            try
            {
                if (dt == null)
                {
                    dt = new data.DataTable();
                    foreach (DataGridViewColumn column in dgv.Columns)
                    {
                        data.DataColumn col = new data.DataColumn(column.Name);
                        dt.Columns.Add(col);
                    }
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        data.DataRow dr = dt.NewRow();
                        for (int i = 0; i < dgv.ColumnCount; i++)
                            dr[i] = row.Cells[i].Value.ToString();
                        dt.Rows.Add(dr);
                    }
                }
                int c = 1;
                foreach (data.DataColumn column in dt.Columns)
                {
                    xlWorkSheet.Cells[1, c] = column.ColumnName;
                    c++;
                }
                for (int i = 2; i <= dt.Rows.Count; i++)
                    for (int j = 1; j <= dt.Columns.Count; j++)
                        xlWorkSheet.Cells[i, j] = dt.Rows[i - 2].ItemArray[j - 1].ToString();
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
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
                    MessageBox.Show("Archivo de configuración vacio");
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
                    MessageBox.Show("Archivo de configuración vacio");
                else
                {
                    settings = new string[keys.Length];
                    int j = 0;
                    foreach (var key in appSettings.AllKeys)
                        for (int i = 0; i < keys.Length; i++)
                            if (keys[i] == key.ToString())
                            {
                                settings[j] = appSettings[key].ToString();
                                j++;
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
                    settigns.Add(key, value);
                else
                    settigns[key].Value = value;
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
