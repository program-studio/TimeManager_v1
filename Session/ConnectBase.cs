using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

namespace TimeTimer
{
    public class ConnectBase
    {
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TimeDB.mdb;";
        //public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\bin\Debug\TimeDB.mdb";
        public string timeBase = "TimeBase";
        private OleDbConnection myConnection;

        void ConnectTo()
        {
            myConnection = new OleDbConnection(connectString);
            //myConnection.Open();

            //myConnection.Close();
        }

        public DataTable Select()
        {
            try
            {
                myConnection.Open();
                //string query = "SELECT * FROM ";
                OleDbCommand cmd = myConnection.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from " + timeBase;
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                myConnection.Close();
                //return dataGridView1.DataSource = dt; ///  переробити
                return dt; ///  переробити
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
    }

}
