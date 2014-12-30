using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
//using System.Data.OracleClient;
using Oracle.DataAccess.Client;
using FirebirdSql.Data.FirebirdClient;


namespace Utilities
{
    public class DataBase
    {
        public static OracleConnection connectOracle(string ip, string service, string user, string pass)
        {
            try
            {
                string connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + ip + ")(PORT=1521))(CONNECT_DATA=(SERVICE_NAME="
                    + service + ")));User Id=" + user + ";Password=" + pass;
                OracleConnection remoteConnection = new OracleConnection(connectionString);
                remoteConnection.Open();
                return remoteConnection;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return null;
            }
        }

        public static void closeOracle(OracleConnection connection) //OracleDataReader dr //Si abrieramos 
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

        public static FbConnection connectFirebird(string user, string pass, string database, string datasource)
        {
            try
            {
                string connectionString = @"user=" + user + "; pass=" + pass + "; database=" + database + "; datasource=" + datasource + ";";
                FbConnection remoteConnection = new FbConnection(connectionString);
                remoteConnection.Open();
                return remoteConnection;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
                return null;
            }
        }

        public static void closeFirebird(FbConnection connection) //Pudiera ir reader
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

        public static DataTable oracleData(string ip, string service, string user, string pass, string instruction)
        {
            DataTable dt = new DataTable();
            try
            {
                OracleConnection conn = connectOracle(ip, service, user, pass);
                //OracleCommand cmd = new OracleCommand(instruction, conn);
                OracleDataAdapter adapter = new OracleDataAdapter();
                adapter.SelectCommand = new OracleCommand(instruction, conn);
                //OracleDataReader reader = cmd.ExecuteReader();
                adapter.Fill(dt);
                closeOracle(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
            return dt;
        }

        public static void oraStatement(string ip, string service, string user, string pass, string instruction) //times
        {
            try
            {
                OracleConnection conn = connectOracle(ip, service, user, pass);
                OracleCommand cmd = new OracleCommand(instruction, conn);
                cmd.ExecuteNonQuery();
                closeOracle(conn);
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

    }
}