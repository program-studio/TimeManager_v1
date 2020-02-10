using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;  // для повного імені користувача
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using Microsoft.SqlServer;
using System.Data.SqlServerCe;


namespace TimeTimer
{
    public class ConnectBase
    {


        private static string dbName = "MyDB.sdf"; //наша БД
        string connectString = "Data Source = " + dbName + "; Password =";
        SqlCeConnection myConnection;
        SqlCeCommand cmd;
        public string timeBase = "TimeBase";

        //Timer tempTimer = new Timer();


        public ConnectBase()
        {
            if (!File.Exists(dbName))
            {
                try
                {
                    SqlCeEngine engine = new SqlCeEngine(connectString);
                    engine.CreateDatabase();
                    RunSQL("CREATE TABLE " + timeBase + "(ID int IDENTITY(1,1) PRIMARY KEY, User_Name nvarchar(40), Break_Type nvarchar(40), Break_Notes nvarchar(100), Start_Time datetime, End_Time datetime, Is_Active int)");  // User_Name, Break_Type, Break_Notes, Start_Time, End_Time
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    myConnection.Close();
                }
            }
        }



        private SqlCeConnection Conection { get { return new SqlCeConnection(connectString); } }
        public DataTable GetData(string qyery)
        {
            DataTable table = new DataTable();
            DataSet ds = new DataSet();
            try
            {
                myConnection = Conection;
                using (myConnection)
                using (SqlCeDataAdapter adapter = new SqlCeDataAdapter(qyery, myConnection))
                {
                    adapter.Fill(ds);
                    table = ds.Tables[0];
                    return table;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return table;
            }
        }
        public bool RunSQL(string qyery)
        {
            try
            {
                myConnection = Conection;
                cmd = new SqlCeCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = qyery;
                cmd.Connection = myConnection;
                myConnection.Open();
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            finally
            {
                if (myConnection != null) myConnection.Close();
            }
        }



        int countConnection = 0;
        //Random rand = new Random();

        MTAThreadAttribute thread = new MTAThreadAttribute();
        private void ConnectTo(bool mErorr = true)
        {
            //InitializeTimerUpdates();
            try
            {


                myConnection = new SqlCeConnection(connectString);
                myConnection.Open();  // 15-31 miliseconds

                countConnection = 0;

            }
            catch (Exception ex)
            {

                if (countConnection < 50000)
                {
                    //rand.Next(0, 5);
                    countConnection++;
                    ConnectTo();
                }
                else
                    if (mErorr) MessageBox.Show(ex.Message);
                //return null;
            }
        }




        public DataTable Select(string query, string userName = null)
        {
            try
            {
                //string query = "SELECT * FROM ";
                ConnectTo();

                SqlCeCommand cmd = myConnection.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = query;
                if (userName != null) cmd.Parameters.AddWithValue("@uName", userName);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlCeDataAdapter da = new SqlCeDataAdapter(cmd);
                da.Fill(dt);
                myConnection.Close();

                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public void Insert(string uName, string breakType, string breakNotes, string startTime, string endTime, string isActive)
        {
            try
            {
                ConnectTo();

                SqlCeCommand cmd = new SqlCeCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO " + timeBase + "(User_Name, Break_Type, Break_Notes, Start_Time, End_Time, Is_Active) VALUES(@uName, @breakType, @breakNotes, @startTime, @endTime, @Is_Active)";
                cmd.Connection = myConnection;
                cmd.Parameters.AddWithValue("@uName", uName);
                cmd.Parameters.AddWithValue("@breakType", breakType);
                cmd.Parameters.AddWithValue("@breakNotes", breakNotes);
                cmd.Parameters.AddWithValue("@startTime", startTime);
                cmd.Parameters.AddWithValue("@endTime", endTime);
                cmd.Parameters.AddWithValue("@Is_Active", isActive);
                cmd.ExecuteNonQuery();
                myConnection.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
        private string ScreenText(string text)
        {
            return text.Replace("'", "''");
        }
        public string Uname()
        {
            UserPrincipal userPrincipal = UserPrincipal.Current;   // повне імя користувача
            string uName = userPrincipal.DisplayName;

            if (uName == null || uName == "")
            {
                uName = "New User";
            }
            else
            {
                uName = userPrincipal.DisplayName;
            }
           
            
            return uName;
                
            //}

            //UserPrincipal userPrincipal = UserPrincipal.Current;   // повне імя користувача
            //string uName = userPrincipal.DisplayName;
            //return uName;
        }

        public DataTable SelectInfo()
        {
            try
            {
                ConnectTo();
                SqlCeCommand cmd = myConnection.CreateCommand();
                cmd.CommandType = CommandType.Text;
                //cmd.CommandText = "SELECT * FROM " + timeBase;
                cmd.CommandText = @"SELECT t1.ID, t1.UserName, t1.Time_Date, @i AS WorkTime, @i:= t1.Time_Date - @i AS Time_Action FROM " + timeBase + @" t1 JOIN(SELECT @i:= 0) var";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlCeDataAdapter da = new SqlCeDataAdapter(cmd);
                da.Fill(dt);
                myConnection.Close();

                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public void Updates(string userName, string endTime, string breakType, string breakNotes, string breakNotesNew, string isActive)
        {
            try
            {
                ConnectTo();

                SqlCeCommand cmd = new SqlCeCommand("UPDATE " + timeBase + " SET [End_Time]=@End_Time, [Break_Notes]=@breakNotesNew, [Is_Active]=@Is_Active  WHERE [Break_Notes]=@Break_Notes AND [Break_Type]=@Break_Type AND [User_Name]=@User_Name", myConnection);

                cmd.Parameters.AddWithValue("@User_Name", userName);
                cmd.Parameters.AddWithValue("@Break_Notes", breakNotes);
                cmd.Parameters.AddWithValue("@Break_Type", breakType);
                cmd.Parameters.AddWithValue("@breakNotesNew", breakNotesNew);
                cmd.Parameters.AddWithValue("@End_Time", endTime);
                cmd.Parameters.AddWithValue("@Is_Active", isActive);

                cmd.Connection = myConnection;
                cmd.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        public void UpdatesWork(string userName, string endTime, string breakType, string breakNotes, string breakNotesNew, string isActive, bool mError = true)
        {
            try
            {
                ConnectTo(mError);

                SqlCeCommand cmd = new SqlCeCommand("UPDATE " + timeBase + " SET [End_Time]=@End_Time, [Break_Notes]=@breakNotesNew, [Is_Active]=@Is_Active  WHERE [Break_Notes]=@Break_Notes AND [Break_Type]=@Break_Type AND [User_Name]=@User_Name", myConnection);

                cmd.Parameters.AddWithValue("@User_Name", userName);
                cmd.Parameters.AddWithValue("@Break_Notes", breakNotes);
                cmd.Parameters.AddWithValue("@Break_Type", breakType);
                cmd.Parameters.AddWithValue("@breakNotesNew", breakNotesNew);
                cmd.Parameters.AddWithValue("@End_Time", endTime);
                cmd.Parameters.AddWithValue("@Is_Active", isActive);

                cmd.Connection = myConnection;
                cmd.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception ex)
            {
                if (mError)
                    MessageBox.Show(ex.Message);
            }
        }

        // Start Form check
        public void UpdatesStart(string userName, string breakType, string breakNotes, string isActive)
        {
            try
            {
                ConnectTo();

                SqlCeCommand cmd = new SqlCeCommand("UPDATE " + timeBase + " SET [Break_Notes]='', [Is_Active]=@Is_Active  WHERE [Break_Notes]=@Break_Notes AND [Break_Type]=@Break_Type AND [User_Name]=@User_Name", myConnection);

                cmd.Parameters.AddWithValue("@User_Name", userName);
                cmd.Parameters.AddWithValue("@Break_Notes", breakNotes);
                cmd.Parameters.AddWithValue("@Break_Type", breakType);
                cmd.Parameters.AddWithValue("@Is_Active", isActive);

                cmd.Connection = myConnection;
                cmd.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }
        public void UpdatesStart(string userName, string breakNotes, string isActive)
        {
            try
            {
                ConnectTo();

                SqlCeCommand cmd = new SqlCeCommand("UPDATE " + timeBase + " SET [Break_Notes]=@Break_Notes, [Is_Active]=@Is_Active  WHERE [User_Name]=@User_Name AND [Break_Notes] LIKE @Type", myConnection);

                cmd.Parameters.AddWithValue("@User_Name", userName);  // [Break_Notes]=@Break_Notes,
                cmd.Parameters.AddWithValue("@Break_Notes", breakNotes.Replace("(The break is ongoing..)", ""));
                //cmd.Parameters.AddWithValue("@End_Time", endTime);
                cmd.Parameters.AddWithValue("@Type", "%(The break is ongoing..)%");
                //cmd.Parameters.AddWithValue("@Break_Type", breakType);
                cmd.Parameters.AddWithValue("@Is_Active", isActive);

                cmd.Connection = myConnection;
                cmd.ExecuteNonQuery();
                myConnection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        //public void InsertTest(DataTable query)
        //{
        //    ConnectTo();

        //    //SqlCeCommand cmd = new SqlCeCommand();
        //    //cmd.CommandType = CommandType.Text;

        //    //cmd.CommandText = "INSERT INTO NewTab SELECT * FROM @MoodleUserTablee";
        //    //cmd.Connection = myConnection;
        //    ////cmd.Parameters.AddWithValue("@uName", uName);
        //    ////cmd.Parameters.AddWithValue("@breakType", breakType);
        //    ////cmd.Parameters.AddWithValue("@breakNotes", breakNotes);
        //    ////cmd.Parameters.AddWithValue("@startTime", startTime);
        //    ////cmd.Parameters.AddWithValue("@endTime", endTime);
        //    //cmd.Parameters.Add(new SqlCeParameter("@MoodleUserTable", query));
        //    //cmd.ExecuteNonQuery();
        //    //myConnection.Close();

        //    SqlCeCommand sqlcmd = new SqlCeCommand("ProductBulkInsertion", myConnection);
        //    sqlcmd.CommandType = CommandType.StoredProcedure;
        //    //sqlcmd.CommandType = CommandType.TableDirect;
        //    //sqlcmd.CommandType = CommandType.Text;
        //    sqlcmd.CommandText = "CREATE PROCEDURE ProductBulkInsertion @product udtProduct readonly AS BEGIN INSERT INTO Product (ProductID, ProductName, ProductCode) SELECT ProductID, ProductName, ProductCode FROM @product END";

        //    sqlcmd.CommandText = "INSERT INTO NewTab SELECT * FROM @MoodleUserTablee";
        //    sqlcmd.Connection = myConnection;
        //    sqlcmd.Parameters.AddWithValue("@MoodleUserTable", query);
        //    sqlcmd.ExecuteNonQuery();
        //    myConnection.Close();

        //}

      

        //private bool IsDatabaseInUse()
        //{
        //    using (SqlConnection sqlConnection = new SqlConnection(connectString))
        //    using (SqlCommand sqlCmd = new SqlCommand())
        //    {
        //        sqlCmd.Connection = sqlConnection;
        //        sqlCmd.CommandText =
        //            @"select count(*)
        //        from sys.dm_tran_locks
        //        where resource_database_id = db_id(@database_name);";
        //        sqlCmd.Parameters.Add(new SqlParameter("@database_name", SqlDbType.NVarChar, 128)
        //        {
        //            Value = dbName
        //        });
        //        sqlConnection.Open();
        //        int sessionCount = Convert.ToInt32(sqlCmd.ExecuteScalar());
        //        if (sessionCount > 0)
        //            return true;
        //        else
        //            return false;
        //    }
        //}


        //private void tempTimer_Tick(object sender, EventArgs e)     // Update кожні 60 секунд !!!!!
        //{

        //        //MessageBox.Show("120");

        //}
        //private void InitializeTimerUpdates()
        //{

        //    tempTimer.Enabled = true;
        //    tempTimer.Start();
        //    tempTimer.Interval = 500;   // 60000 = 60 секунд







    }

}
