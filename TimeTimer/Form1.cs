using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;  // для повного імені користувача
using System.Data.SqlClient;
using Microsoft.Win32;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.Reflection;
using ADGV;
using IWshRuntimeLibrary;
//using System.Runtime.InteropServices;
//using Microsoft.Office.Interop.Outlook;


namespace TimeTimer
{
    public partial class TimeManager : Form
    {

        private static ConnectBase db = new ConnectBase();
        private static string userName = db.Uname();
        private List<UserInfo> usersInfo = null;
        private ADGVManger dgv;
        private Button bt;
        private DateTime startWorkTime;
        private DateTime startTime;
        private bool sessionLock = false;
        private TimeSpan workInterval;
        private TimeSpan pauseInterval;
        private TimeSpan workIntervalTemp;
        //private System.Windows.Forms.Timer timerUpdate = new System.Windows.Forms.Timer();


        public TimeManager()
        {
            InitializeComponent();

            userNametxt.Text += userName;

            DateTime startTime = DateTime.Now;
            StartTime2txt.Text = startTime.ToLongTimeString();

            Pausebtm.Enabled = false;
            Notestxt.Enabled = false;

            BreakDinnerbtm.Click += mainBt_Click;
            BreakPausebtm.Click += mainBt_Click;
            BreakMeetingbtm.Click += mainBt_Click;
            BreakStudybtm.Click += mainBt_Click;
            BreakNotebtm.Click += mainBt_Click;
            BreakDoctorbtm.Click += mainBt_Click;

            //BreakDinnerbtm.Enabled = false;
            //BreakPausebtm.Enabled = false;
            //BreakMeetingbtm.Enabled = false;
            //BreakStudybtm.Enabled = false;
            //BreakNotebtm.Enabled = false;
            //BreakDoctorbtm.Enabled = false;

            tableLayoutPanel1.Enabled = false;

            label11.Visible = false;



            ChangeButton();
            SystemEvents.SessionSwitch += SESS;    // блокування компютера - монітора

            


        }





        private void Form1_Load(object sender, EventArgs e)
        {

            //   !!!! не стирати 
            tabControl1.TabPages.Remove(tabPage3);
            //tabControl1.TabPages.Add(tabPage2);
            
            ////////////////////////////////////////////////////////////////////////////////////

            dgv = new ADGVManger(panel2);

            //if (Environment.UserName.Equals("tssavka") || Environment.UserName.Equals("tarassmith"))
            if (Environment.UserName.Equals("tssavka"))
            {
            DataTable dt = db.Select("SELECT DISTINCT (User_Name) as UserName FROM " + db.timeBase);
            comboBoxUsers.Items.Add("All employees");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                    comboBoxUsers.Items.Add(dt.Rows[i]["UserName"].ToString());
            }
            }
            else
            {
                //comboBoxUsers.Items.Add(userName);
                
                comboBoxUsers.Items.Add("All employees");
                DataTable dt = db.Select("SELECT DISTINCT (User_Name) as UserName FROM " + db.timeBase);
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                        comboBoxUsers.Items.Add(dt.Rows[i]["UserName"].ToString());
                }
            }
            comboBoxUsers.SelectedIndex = 0;


            // Start Form check !

            string startActiveNote;
            DataTable dtStart = db.Select("SELECT ID, Break_Notes, Break_Type  FROM " + db.timeBase + " WHERE User_Name = '" + userName + "' AND Is_Active = 1");
            if (dtStart != null)
            {
                for (int i = 0; i < dtStart.Rows.Count; i++)
                {
                    dtStart.Rows[i]["ID"].ToString();
                    startActiveNote = dtStart.Rows[i]["Break_Notes"].ToString();

                    if (startActiveNote.Contains("(Work is ongoing..)")) { db.UpdatesStart(userName, "The break is ongoing", "(Work is ongoing..)", "0"); }
                        
                    if (startActiveNote.Contains("(The break is ongoing..)")) {db.UpdatesStart(userName, startActiveNote, "0");}

                }
            }


            //InitializeTimerUpdate();   // оновлення інформації в БД по таймеру (апдейт)

            //timerUpdate.Enabled = true;

            AutoRun(0); // Автозапуск програми



        }



        //int countTimer = 0;  // поч час
        //int countTimer2 = 0;  // поч час
        //int countTimer3 = 0;  // поч час

        //DateTime t1, t2;
        //DateTime t3, t4;
        //DateTime t5, t6;


        private void Startbtm_Click(object sender, EventArgs e)
        {
            InitializeTimerUpdate();   // оновлення інформації в БД по таймеру (апдейт)
            //tabControl1.TabPages.Add(tabPage3);
            startWorkTime = DateTime.Now;
            StartTime2txt.Text = startWorkTime.ToLongTimeString();


            //timerStart.Enabled = !timerStart.Enabled;
            timerStart.Start();


            if (Startbtm.Text == "Start work")
            {
                Startbtm.Text = "Work started";
                Pausebtm.Enabled = false;
                Startbtm.Enabled = false;
            }
            else
            {
                Startbtm.Text = "Start work";

            }

            //BreakDinnerbtm.Enabled = true;
            //BreakPausebtm.Enabled = true;
            //BreakMeetingbtm.Enabled = true;
            //BreakStudybtm.Enabled = true;
            //BreakNotebtm.Enabled = true;
            //BreakDoctorbtm.Enabled = true;

            tableLayoutPanel1.Enabled = true;

            db.Insert(userName, "Period of work", "(Work is ongoing..)", startWorkTime.ToString(), DateTime.Now.AddSeconds(1).ToString(), "1");  // "00:00:00"

        }

        private void Pausebtm_Click(object sender, EventArgs e)
        {

            //timerPause.Enabled = !timerPause.Enabled;  //  +
            //timerStart.Enabled = !timerPause.Enabled;


            if (Pausebtm.Text == "Break")
            {
                //timerStart.Stop();
                //timerPause.Start();

                Pausebtm.Text = "End the break";
                Pausebtm.BackColor = Color.Red;
                Notestxt.Enabled = false;
                startTime = DateTime.Now;
                //if (bt != null) db.Insert(userName, bt.Text, Notestxt.Text, startWorkTime.ToString(), DateTime.Now.ToString());
                //if (bt != null) db.Insert(userName, DateTime.Now.ToString(), workTimetxt.Text, pauseTimetxt.Text, bt.Text, Notestxt.Text);  // "Перерва"
                if (Startbtm.Enabled != true) db.Insert(userName, bt.Text.Replace("\r\n", ""), Notestxt.Text + " (The break is ongoing..)", startTime.ToString(), DateTime.Now.AddSeconds(1).ToString(), "1");

                //BreakDinnerbtm.Enabled = false;
                //BreakPausebtm.Enabled = false;
                //BreakMeetingbtm.Enabled = false;
                //BreakStudybtm.Enabled = false;
                //BreakNotebtm.Enabled = false;
                //BreakDoctorbtm.Enabled = false;

                tableLayoutPanel1.Enabled = false;

                label11.Visible = false;

            }
            else
            {
                //timerPause.Stop();
                //timerStart.Start();
                workIntervalTemp += DateTime.Now - startTime;
                if (Startbtm.Enabled != true) db.Updates(userName, DateTime.Now.AddSeconds(1).ToString(), bt.Text.Replace("\r\n", ""), Notestxt.Text + " (The break is ongoing..)", Notestxt.Text, "0");
                // db.UpdatesPause(userName, Notestxt.Text, DateTime.Now.ToString());

                

                Pausebtm.Text = "Break";
                Pausebtm.BackColor = Color.DarkOrange;
                Pausebtm.Enabled = false;
                //if (bt != null) db.Insert(userName, bt.Text.Replace("\r\n", ""), Notestxt.Text, startTime.ToString(), DateTime.Now.ToString());
                Notestxt.BackColor = Color.White;
                Notestxt.Text = "";
                //if (Startbtm.Enabled != true) db.UpdatesPause(userName, "%(Перерва триває)%",  DateTime.Now.ToString());

                //BreakDinnerbtm.Enabled = true;
                //BreakPausebtm.Enabled = true;
                //BreakMeetingbtm.Enabled = true;
                //BreakStudybtm.Enabled = true;
                //BreakNotebtm.Enabled = true;
                //BreakDoctorbtm.Enabled = true;

                tableLayoutPanel1.Enabled = true;

                bt.BackColor = Color.White;
            }


        }

        private void Updatebtm_Click(object sender, EventArgs e)
        {
            Updatebtm.Enabled = false;
            if (tabControl1.SelectedIndex == 1 && tabControl2.SelectedIndex == 0)
            {
                SelectBase();
                ChartUpdate(usersInfo);
            }
            else
            {
                SelectBase();
                ChartUpdateC(usersInfo);
            }
            Updatebtm.Enabled = true;


        }   // +++

        private void ChangeButton()
        {
            string breakDinner = "Icon\\coffee.png"; // шлях до іконки
            string breakMeeting = "Icon\\people.png";
            string breakPause = "Icon\\break.png";
            string breakStudy = "Icon\\study.png";
            string breakNote = "Icon\\note.png";
            string breakDoc = "Icon\\doctor.png";
            string closePic = "Icon\\close.png";
            string minimizePic = "Icon\\minimize.png";
            string startTimePic = "Icon\\start_time.png";
            string endTimePic = "Icon\\end_time.png";
            string compPic = "Icon\\comp.png";


            Bitmap image = Bitmap.FromFile(breakDinner) as Bitmap;  // присвоюємо іконку по шляху oFile        
            BreakDinnerbtm.Image = new Bitmap(image, new Size(18, 18));  // зміна розміру іконки
            label02.Image = new Bitmap(image, new Size(12, 12));

            Bitmap image2 = Bitmap.FromFile(breakPause) as Bitmap;  // присвоюємо іконку по шляху oFile
            BreakPausebtm.Image = new Bitmap(image2, new Size(20, 20));  // зміна розміру іконки
            label03.Image = new Bitmap(image2, new Size(12, 12));

            Bitmap image3 = Bitmap.FromFile(breakMeeting) as Bitmap;  // присвоюємо іконку по шляху oFile
            BreakMeetingbtm.Image = new Bitmap(image3, new Size(20, 20));  // зміна розміру іконки
            label04.Image = new Bitmap(image3, new Size(13, 13));

            Bitmap image4 = Bitmap.FromFile(breakStudy) as Bitmap;  // присвоюємо іконку по шляху oFile
            BreakStudybtm.Image = new Bitmap(image4, new Size(20, 20));  // зміна розміру іконки
            label05.Image = new Bitmap(image4, new Size(13, 13));

            Bitmap image5 = Bitmap.FromFile(breakNote) as Bitmap;  // присвоюємо іконку по шляху oFile
            BreakNotebtm.Image = new Bitmap(image5, new Size(18, 18));  // зміна розміру іконки
            label06.Image = new Bitmap(image5, new Size(9, 9));

            Bitmap image6 = Bitmap.FromFile(breakDoc) as Bitmap;  // присвоюємо іконку по шляху oFile
            BreakDoctorbtm.Image = new Bitmap(image6, new Size(18, 18));  // зміна розміру іконки
            label07.Image = new Bitmap(image6, new Size(12, 12));

            Bitmap image7 = Bitmap.FromFile(closePic) as Bitmap;  // присвоюємо іконку по шляху oFile
            closeBtm.Image = new Bitmap(image7, new Size(15, 15));  // зміна розміру іконки

            Bitmap image8 = Bitmap.FromFile(minimizePic) as Bitmap;  // присвоюємо іконку по шляху oFile
            minimizeBtm.Image = new Bitmap(image8, new Size(15, 15));  // зміна розміру іконки

            Bitmap image9 = Bitmap.FromFile(startTimePic) as Bitmap;  // присвоюємо іконку по шляху oFile
            Startbtm.Image = new Bitmap(image9, new Size(20, 20));  // зміна розміру іконки


            Bitmap image10 = Bitmap.FromFile(endTimePic) as Bitmap;  // присвоюємо іконку по шляху oFile
            Pausebtm.Image = new Bitmap(image10, new Size(20, 20));  // зміна розміру іконки

            Bitmap image11 = Bitmap.FromFile(compPic) as Bitmap;  // присвоюємо іконку по шляху oFile
            startTimelb.Image = new Bitmap(image11, new Size(19, 23));  // зміна розміру іконки
            label01.Image = new Bitmap(image11, new Size(12, 12));



            //BreakDinnerbtm.ImageAlign = ContentAlignment.TopCenter;  //  зміна позиції кнопки !!!!!!!!!!!!
            //BreakDinnerbtm.TextAlign = ContentAlignment.BottomCenter;   //  зміна позиції тексту !!!!!!!!!!!!
            //BreakDinnerbtm.Location = new System.Drawing.Point(10, 10);    // зміщення позиції кнопки
            //BreakDinnerbtm.Size = new System.Drawing.Size(100, 100);   // зміна розміру кнопки


        }   //  Зміна Кнопок Іконки

        
        private void timerStart_Tick(object sender, EventArgs e)  // +++
        {
            //workInterval = DateTime.Now - startWorkTime - workIntervalTemp;
            //if (Startbtm.Enabled != true) workTimetxt.Text = String.Format("{0:hh}:{0:mm}:{0:ss}", workInterval);
            workInterval = DateTime.Now - startWorkTime - workIntervalTemp;
            if (Startbtm.Enabled != true && Pausebtm.Text == "Break") workTimetxt.Text = String.Format("{0:hh}:{0:mm}:{0:ss}", workInterval);

            pauseInterval = (DateTime.Now - startTime) + workIntervalTemp;
            //pauseInterval = workIntervalTemp;
            if (Pausebtm.Enabled != false && Pausebtm.Text != "Break") pauseTimetxt.Text = String.Format("{0:hh}:{0:mm}:{0:ss}", pauseInterval);

            //countTimer++;  // поч час  ++


            //t1 = new DateTime(DateTime.Now.Day);

            //t2 = t1.AddHours((double)countTimer);
            //t2 = t1.AddMinutes((double)countTimer);
            //t2 = t1.AddSeconds((double)countTimer);


            //if (t2.Hour < 10)

            //    workTimetxt.Text = "0" + t2.Hour.ToString() + ":";
            //else
            //    workTimetxt.Text = t2.Hour.ToString() + ":";
            //if (t2.Minute < 10)
            //    workTimetxt.Text += "0" + t2.Minute.ToString() + ":";
            //else
            //    workTimetxt.Text += t2.Minute.ToString() + ":";
            //if (t2.Second < 10)
            //    workTimetxt.Text += "0" + t2.Second.ToString();
            //else
            //    workTimetxt.Text += t2.Second.ToString();



            /////////


            //Starttxt.Text = "00" + "." + countTimer.ToString();

            //second += 1;


            //if (second == 60)
            //{
            //    second = 0;
            //    minute += 1;
            //    //Starttxt.Text = "0" + second;
            //}
            ////Starttxt.Text = "0" + second;
            //if (minute == 60)
            //{
            //    minute = 0;
            //    hour += 1;
            //    //Starttxt.Text = "0" + minute;
            //}
            ////Starttxt.Text += "0" + minute;
            //if (hour == 10)
            //{
            //    hour = 0;

            //}
            //Starttxt.Text += "0" + hour;

            //Starttxt.Text = "0" + hour;
            //Starttxt.Text = hour + "." + minute + "." + second;
            //int tt = 0;

        }

        //private void timerPause_Tick(object sender, EventArgs e)  // +++
        //{
        //    //pauseInterval = (DateTime.Now - startTime) + workIntervalTemp; 
        //    //if (Pausebtm.Enabled != false &&  Pausebtm.Text != "Перерва") pauseTimetxt.Text = String.Format("{0:hh}:{0:mm}:{0:ss}", pauseInterval);
           
        //    //countTimer2++;  // поч час  ++


        //    //t3 = new DateTime(DateTime.Today.Day);

        //    //t4 = t3.AddHours((double)countTimer2);
        //    //t4 = t3.AddMinutes((double)countTimer2);
        //    //t4 = t3.AddSeconds((double)countTimer2);

        //    //if (t4.Hour < 10)

        //    //    pauseTimetxt.Text = "0" + t4.Hour.ToString() + ":";
        //    //else
        //    //    pauseTimetxt.Text = t4.Hour.ToString() + ":";
        //    //if (t4.Minute < 10)
        //    //    pauseTimetxt.Text += "0" + t4.Minute.ToString() + ":";
        //    //else
        //    //    pauseTimetxt.Text += t4.Minute.ToString() + ":";
        //    //if (t4.Second < 10)
        //    //    pauseTimetxt.Text += "0" + t4.Second.ToString();
        //    //else
        //    //    pauseTimetxt.Text += t4.Second.ToString();
        //}


        private void timerUpdate_Tick(object sender, EventArgs e)     // Update кожні 60 секунд !!!!!
        {
            if (Startbtm.Enabled != true && sessionLock == false)
            {
                //if (e.Reason == SessionSwitchReason.SessionLock)
                //if (e.Reason == SessionSwitchReason.SessionUnlock

                //timerUpdate.Interval = 3;
                //countTimer3++;  // поч час  ++
                //t5 = new DateTime(DateTime.Today.Day);
                //t6 = t5.AddSeconds((double)countTimer3);
                //if (t6.Second == 60)
                //{

                if (Startbtm.Enabled != true) db.UpdatesWork(userName, DateTime.Now.ToString(), "Period of work", "(Work is ongoing..)", "(Work is ongoing..)", "1", false);
                //if (Startbtm.Enabled != true && Pausebtm.Text == "Завершити перерву") db.UpdatesWork(userName, DateTime.Now.ToString(), bt.Text.Replace("\r\n", ""),
                //    Notestxt.Text + " (Перерва триває..)", Notestxt.Text + " (Перерва триває..)", "1", false);

                //}

                //MessageBox.Show("120");
            }
        }


        private void InitializeTimerUpdate()
        {

            timerUpdate.Enabled = true;
            timerUpdate.Start();
            timerUpdate.Interval = 120000;   // 60000 = 60 секунд

            timerStart.Enabled = true;
            timerStart.Start();
            timerStart.Interval = 1000;

            //timerPause.Enabled = true;
            //timerPause.Interval = 1000;

        }   //  Запуск Таймерів



        private void Notestxt_Click(object sender, EventArgs e)
        {
            Notestxt.BackColor = Color.White;
            Notestxt.ForeColor = Color.Black;
            Notestxt.Text = "";
        }


        private void closeBtm_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }


        private void minimizeBtm_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }


        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Left += e.X - lastPoint.X;
                this.Top += e.Y - lastPoint.Y;
            }
        }

        Point lastPoint;
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            lastPoint = new Point(e.X, e.Y);
        }


        // Завершення роботи / закриття форми !


        private void TimeManager_FormClosing(object sender, FormClosingEventArgs e)  // Завершення роботи / закриття форми !
        {
            ////string startWorkTimeRep = startWorkTime.ToString("dd.MM.yyyy").Replace(".", "/");  //  "MM.dd.yyyy H:mm:ss"
            string startWorkTimeRep = startWorkTime.ToString("MM.dd.yyyy HH:mm").Replace(".", "/");  //  "MM.dd.yyyy H:mm:ss"
            if (Startbtm.Enabled != true) db.Updates(userName, DateTime.Now.AddSeconds(1).ToString(), "Period of work", "(Work is ongoing..)", "", "0");
            if (Startbtm.Enabled != true && Pausebtm.Text == "End the break") db.Updates(userName, DateTime.Now.AddSeconds(1).ToString(), bt.Text.Replace("\r\n", ""), Notestxt.Text + " (The break is ongoing..)", Notestxt.Text, "0");


        }




        private void minimizeBtm_MouseLeave(object sender, EventArgs e)
        {
            minimizeBtm.BackColor = Color.Transparent;
        }

        private void minimizeBtm_MouseHover(object sender, EventArgs e)
        {
            minimizeBtm.BackColor = Color.Silver;
        }

        private void closeBtm_MouseHover(object sender, EventArgs e)
        {
            closeBtm.BackColor = Color.Silver;
        }
        private void closeBtm_MouseLeave(object sender, EventArgs e)
        {
            closeBtm.BackColor = Color.Transparent;
        }




        private void mainBt_Click(object sender, EventArgs e)
        {
            bt = (Button)sender;
            //bt.BackColor = Color.LightSalmon;

            foreach (Button control in this.tableLayoutPanel1.Controls.OfType<Button>())
            {

                control.BackColor = Color.White;

            }
            bt.BackColor = Color.LightGray;


            Pausebtm.Enabled = true;
            label11.Visible = true;
            //string commentText = string.Empty;
            if (bt.Name.Equals(BreakDinnerbtm.Name) || bt.Name.Equals(BreakPausebtm.Name) || bt.Name.Equals(BreakDoctorbtm.Name) || bt.Name.Equals(BreakNotebtm.Name))
            {

                Notestxt.Enabled = false;
                Notestxt.BackColor = Color.White;
                Notestxt.Text = "";
            }

            else
            {
                Notestxt.Enabled = true;
                Notestxt.BackColor = Color.LightSalmon;
                Notestxt.ForeColor = Color.Red;
                Notestxt.Text = " Enter text";
            }


        }


        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)  // ОНОВЛЕННЯ Графіка на вкладці
        {
            if (tabControl1.SelectedIndex == 1 && tabControl2.SelectedIndex == 0)
            {
                SelectBase();
                ChartUpdate(usersInfo);
            }
            if (tabControl1.SelectedIndex == 1 && tabControl2.SelectedIndex == 1)
            {
                SelectBase();
                ChartUpdateC(usersInfo);
            }
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1 && tabControl2.SelectedIndex == 1)
            {
                SelectBase();
                ChartUpdateC(usersInfo);
            }
            if (tabControl1.SelectedIndex == 1 && tabControl2.SelectedIndex == 0)
            {
                SelectBase();
                ChartUpdate(usersInfo);
            }
        }




        private void ChartUpdate(List<UserInfo> users)
        {
            chartTime.Series.Clear();
            chartTime.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.BrightPastel;
            System.Drawing.Font chtFont = new System.Drawing.Font("Arial", 8);
            System.Drawing.Font chtFont2 = new System.Drawing.Font("Arial", 8);
            chartTime.ChartAreas["ChartArea1"].AxisX.LabelStyle.Font = chtFont2;
            chartTime.ChartAreas["ChartArea1"].AxisY.LabelStyle.Font = chtFont2;
            chartTime.ChartAreas["ChartArea1"].AxisX.Interval = 1;

            chartTime.Series.Add("workTime");
            chartTime.Series["workTime"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
            chartTime.Series["workTime"].Color = Color.MediumSeaGreen;
            chartTime.Series["workTime"].LegendText = "Work time";
            chartTime.Series["workTime"].Font = chtFont;

            chartTime.Series.Add("breakTime");
            chartTime.Series["breakTime"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.StackedBar100;
            chartTime.Series["breakTime"].Color = Color.DarkOrange;
            chartTime.Series["breakTime"].LegendText = "Break";
            chartTime.Series["breakTime"].Font = chtFont;


            //chartTime.ChartAreas["ChartArea1"].AxisX.TitleFont = new System.Drawing.Font("Trebuchet MS", 2.25F, System.Drawing.FontStyle.Bold);
            //chartTime.Series["breakTime"].Label = "#VAL\n %";
            int i = 0;

            if (usersInfo != null && usersInfo.Count > 0)
            {
                foreach (UserInfo user in usersInfo)
                {
                    //if (user.WorkPeriods() != 0)
                    //{
                    chartTime.Series["workTime"].Points.AddXY(user.Name, Convert.ToInt32(user.WorkPeriods()));
                    chartTime.Series["workTime"].Points[i].Label = "#VAL%\n(" + (user.PeriodWork - (user.BreakDinner + user.BreakPause + user.BreakMeeting + user.BreakNote + user.BreakStudy + user.BreakDoctor)).ToString() + ")";
                    //    chartTime.Series["workTime"].Points.AddXY(user.Name, user.WorkPeriods());
                    //    chartTime.Series["workTime"].Points[i].Label = "#VAL%\n(" + user.PeriodWork.ToString() + ")";

                    chartTime.Series["breakTime"].Points.AddXY(user.Name, Convert.ToInt32(user.BreakPeriods()));
                    chartTime.Series["breakTime"].Points[i].Label = "#VAL%\n(" + user.PeriodBreaktxt().ToString() + ")";
                    //chartTime.ChartAreas.["breakTime"].AxisY.LabelAutoFitStyle = LabelAutoFitStyles.None;


                    //chartTime.Series["breakTime"].Points[i].Label = user.BreakPeriods().ToString();

                    i++;
                }
            }
        }


        private void ChartUpdateC(List<UserInfo> users)
        {

            chartTimeC.Series.Clear();
            chartTimeC.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Pastel;
            chartTimeC.Series.Add("workTime");
            chartTimeC.Series["workTime"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;

            TimeSpan PeriodWorkAll = TimeSpan.Zero;
            TimeSpan BreakDinnerAll = TimeSpan.Zero;
            TimeSpan BreakPauseAll = TimeSpan.Zero;
            TimeSpan BreakMeetingAll = TimeSpan.Zero;
            TimeSpan BreakStudyAll = TimeSpan.Zero;
            TimeSpan BreakNoteAll = TimeSpan.Zero;
            TimeSpan BreakDoctorAll = TimeSpan.Zero;
            TimeSpan PeriodWorkTotal = TimeSpan.Zero;
            int procentWorkAll = 0;
            int procentDinnerAll = 0;
            int procentPauseAll = 0;
            int procentMeetingAll = 0;
            int procentStudyAll = 0;
            int procentNoteAll = 0;
            int procentDoctorAll = 0;
            int w = 0;
            int d = 0;
            int p = 0;
            int m = 0;
            int s = 0;
            int n = 0;
            int dc = 0;

            if (usersInfo != null && usersInfo.Count == 1)  // > 0
            {
                foreach (UserInfo user in usersInfo)
                {
                    label1.Text = (user.PeriodWork - (user.BreakDinner + user.BreakPause + user.BreakMeeting + user.BreakNote + user.BreakStudy + user.BreakDoctor)).ToString() + " (" + user.ProcentWorks("w").ToString("0") + "%)"; ;
                    label2.Text = user.BreakDinner.ToString() + " (" + user.ProcentWorks("d").ToString("0") + "%)";
                    label3.Text = user.BreakPause.ToString() + " (" + user.ProcentWorks("p").ToString("0") + "%)";
                    label4.Text = user.BreakMeeting.ToString() + " (" + user.ProcentWorks("m").ToString("0") + "%)";
                    label5.Text = user.BreakStudy.ToString() + " (" + user.ProcentWorks("s").ToString("0") + "%)";
                    label6.Text = user.BreakNote.ToString() + " (" + user.ProcentWorks("n").ToString("0") + "%)";
                    label7.Text = user.BreakDoctor.ToString() + " (" + user.ProcentWorks("dc").ToString("0") + "%)";
                    label8.Text = user.PeriodWork.ToString();


                    chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(user.ProcentWorks("w")));
                    //chartTimeC.Series["workTime"].Points[0].Color = Color.Red;
                    //chartTimeC.Series["workTime"].Points[0].Label = "#VAL%\n(" + user.PeriodWork.ToString() + ")";
                    chartTimeC.Series["workTime"].Points[0].Label = "#VAL%";
                    chartTimeC.Series["workTime"].Points[0].LegendText = "Work";
                    chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(user.ProcentWorks("d")));
                    chartTimeC.Series["workTime"].Points[1].Label = "#VAL%";
                    chartTimeC.Series["workTime"].Points[1].LegendText = "Lunch";
                    chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(user.ProcentWorks("p")));
                    chartTimeC.Series["workTime"].Points[2].Label = "#VAL%";
                    chartTimeC.Series["workTime"].Points[2].LegendText = "Break";
                    chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(user.ProcentWorks("m")));
                    chartTimeC.Series["workTime"].Points[3].Label = "#VAL%";
                    chartTimeC.Series["workTime"].Points[3].LegendText = "Metting";
                    chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(user.ProcentWorks("s")));
                    chartTimeC.Series["workTime"].Points[4].Label = "#VAL%";
                    chartTimeC.Series["workTime"].Points[4].LegendText = "Study";
                    chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(user.ProcentWorks("n")));
                    chartTimeC.Series["workTime"].Points[5].Label = "#VAL%";
                    chartTimeC.Series["workTime"].Points[5].LegendText = "Exit note";
                    chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(user.ProcentWorks("dc")));
                    chartTimeC.Series["workTime"].Points[6].Label = "#VAL%";
                    chartTimeC.Series["workTime"].Points[6].LegendText = "To the doctor";

                }
            }

            else if (usersInfo != null && usersInfo.Count > 1)
            {

                foreach (UserInfo user in usersInfo)
                {

                    PeriodWorkAll += user.PeriodWork - (user.BreakDinner + user.BreakPause + user.BreakMeeting + user.BreakNote + user.BreakStudy + user.BreakDoctor);
                    BreakDinnerAll += user.BreakDinner;
                    BreakPauseAll += user.BreakPause;
                    BreakMeetingAll += user.BreakMeeting;
                    BreakStudyAll += user.BreakStudy;
                    BreakNoteAll += user.BreakNote;
                    BreakDoctorAll += user.BreakDoctor;
                    PeriodWorkTotal += user.PeriodWork;

                    if (user.ProcentWorks("w") != 0) { procentWorkAll += Convert.ToInt32(user.ProcentWorks("w")); w += 1; }
                    if (user.ProcentWorks("d") != 0) { procentDinnerAll += Convert.ToInt32(user.ProcentWorks("d")); d += 1; }
                    if (user.ProcentWorks("p") != 0) { procentPauseAll += Convert.ToInt32(user.ProcentWorks("p")); p += 1; }
                    if (user.ProcentWorks("m") != 0) { procentMeetingAll += Convert.ToInt32(user.ProcentWorks("m")); m += 1; }
                    if (user.ProcentWorks("s") != 0) { procentStudyAll += Convert.ToInt32(user.ProcentWorks("s")); s += 1; }
                    if (user.ProcentWorks("n") != 0) { procentNoteAll += Convert.ToInt32(user.ProcentWorks("n")); n += 1; }
                    if (user.ProcentWorks("dc") != 0) { procentDoctorAll += Convert.ToInt32(user.ProcentWorks("dc")); dc += 1; }

                }

                if (w != 0) label1.Text = PeriodWorkAll.ToString() + " (" + Convert.ToInt32((procentWorkAll / w)) + "%)"; else label1.Text = PeriodWorkAll.ToString() + " (0%)";
                //label1.Text = PeriodWorkAll.ToString();
                if (d != 0) label2.Text = BreakDinnerAll.ToString() + " (" + Convert.ToInt32((procentDinnerAll / d)) + "%)"; else label2.Text = BreakDinnerAll.ToString() + " (0%)";
                if (p != 0) label3.Text = BreakPauseAll.ToString() + " (" + Convert.ToInt32((procentPauseAll / p)) + "%)"; else label3.Text = BreakPauseAll.ToString() + " (0%)";
                if (m != 0) label4.Text = BreakMeetingAll.ToString() + " (" + Convert.ToInt32((procentMeetingAll / m)) + "%)"; else label4.Text = BreakMeetingAll.ToString() + " (0%)";
                if (s != 0) label5.Text = BreakStudyAll.ToString() + " (" + Convert.ToInt32((procentStudyAll / s)) + "%)"; else label5.Text = BreakStudyAll.ToString() + " (0%)";
                if (n != 0) label6.Text = BreakNoteAll.ToString() + " (" + Convert.ToInt32((procentNoteAll / n)) + "%)"; else label6.Text = BreakNoteAll.ToString() + " (0%)";
                if (dc != 0) label7.Text = BreakDoctorAll.ToString() + " (" + Convert.ToInt32((procentDoctorAll / dc)) + "%)"; else label7.Text = BreakDoctorAll.ToString() + " (0%)";
                label8.Text = PeriodWorkTotal.ToString();

                if (w != 0) chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(procentWorkAll / w)); else chartTimeC.Series["workTime"].Points.Add(procentWorkAll);
                chartTimeC.Series["workTime"].Points[0].Label = "#VAL%";
                chartTimeC.Series["workTime"].Points[0].LegendText = "Work";
                if (d != 0) chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(procentDinnerAll / d)); else chartTimeC.Series["workTime"].Points.Add(procentDinnerAll);
                chartTimeC.Series["workTime"].Points[1].Label = "#VAL%";
                chartTimeC.Series["workTime"].Points[1].LegendText = "Lunch";
                if (p != 0) chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(procentPauseAll / p)); else chartTimeC.Series["workTime"].Points.Add(procentPauseAll);
                chartTimeC.Series["workTime"].Points[2].Label = "#VAL%";
                chartTimeC.Series["workTime"].Points[2].LegendText = "Break";
                if (m != 0) chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(procentMeetingAll / m)); else chartTimeC.Series["workTime"].Points.Add(procentMeetingAll);
                chartTimeC.Series["workTime"].Points[3].Label = "#VAL%";
                chartTimeC.Series["workTime"].Points[3].LegendText = "Meeting";
                if (s != 0) chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(procentStudyAll / s)); else chartTimeC.Series["workTime"].Points.Add(procentStudyAll);
                chartTimeC.Series["workTime"].Points[4].Label = "#VAL%";
                chartTimeC.Series["workTime"].Points[4].LegendText = "Study";
                if (n != 0) chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(procentNoteAll / n)); else chartTimeC.Series["workTime"].Points.Add(procentNoteAll);
                chartTimeC.Series["workTime"].Points[5].Label = "#VAL%";
                chartTimeC.Series["workTime"].Points[5].LegendText = "Exit note";
                if (dc != 0) chartTimeC.Series["workTime"].Points.Add(Convert.ToInt32(procentDoctorAll / dc)); else chartTimeC.Series["workTime"].Points.Add(procentDoctorAll);
                chartTimeC.Series["workTime"].Points[6].Label = "#VAL%";
                chartTimeC.Series["workTime"].Points[6].LegendText = "To the doctor";

            }
            else
            {
                label1.Text = PeriodWorkAll.ToString();
                label2.Text = BreakDinnerAll.ToString();
                label3.Text = BreakPauseAll.ToString();
                label4.Text = BreakMeetingAll.ToString();
                label5.Text = BreakStudyAll.ToString();
                label6.Text = BreakNoteAll.ToString();
                label7.Text = BreakDoctorAll.ToString();
                label8.Text = PeriodWorkTotal.ToString();
            }
        }



        // робота з ЕКСЕЛЬ


        //private void button1_Click(object sender, EventArgs e)
        //{
        //    //Приложение
        //    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
        //    Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
        //    Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
        //    //Книга.
        //    ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
        //    //Таблица.
        //    ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
        //    ExcelApp.Cells[1,1] = "Звіт робочого часу працівників";
        //    ExcelApp.Cells[2, 1] = "За період";
        //    ExcelApp.Cells[2, 2] = " з " + dtpInfoDate.Value.ToString("dd.MM.yyyy") + " по "+ dtpInfoDateEnd.Value.ToString("dd.MM.yyyy");
        //    //ExcelApp.Cells[1, 1].Font.Color = Color.Blue;  // колір тексту
        //    //ExcelApp.Cells[2, 1].Font.Color = Color.Blue;  // колір тексту
        //    //ExcelApp.Cells[2, 2].Font.Color = Color.Blue;  // колір тексту

        //    ExcelApp.Cells[4, 1] = "Працівник";
        //    ExcelApp.Cells[4, 2] = "Активність";
        //    ExcelApp.Cells[4, 3] = "Коментар";
        //    ExcelApp.Cells[4, 4] = "Час Початку";
        //    ExcelApp.Cells[4, 5] = "Час завершення";
        //    ExcelApp.Range["A4", "E4"].Interior.Color = Color.PaleGreen;
        //    ExcelApp.Range["A4", "E4"].ColumnWidth = 16;  // ширина
        //    ExcelApp.Range["A4", "E4"].RowHeight = 40;   // висота
        //    ExcelApp.Range["A4", "E4"].HorizontalAlignment = HorizontalAlignment.Center;
        //    //ExcelApp.Range["A4", "E4"].VerticalAlignment = ali;

        //    //ExcelApp.Cells[4, 1].Interior.Color = 0xFF00;  // колір клітинки

        //    //ExcelApp.Cells[4, 1].Font.Color = Color.Red;  // колір тексту



        //    for (int i = 0; i < dataGridView1.Rows.Count; i++)
        //    {

        //        for (int j = 0; j < dataGridView1.ColumnCount; j++)
        //        {

        //            ExcelApp.Cells[i + 5, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
        //            //ExcelApp.Cells[4, i+1].Interior.Color = Color.PaleGreen; //  0xFF00  // колір клітинки
        //        }
        //    }
        //    //Вызываем нашу созданную эксельку.
        //    ExcelApp.Visible = true;
        //    ExcelApp.UserControl = true;
        //}





        // блокування компютера - монітора



        DateTime avtoStart = new DateTime();

        private void SESS(object sender, SessionSwitchEventArgs e)    // блокування компютера - монітора
        {
            //DateTime avtoStart = new DateTime();
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                // при блокуванні - занести в змінну час початку   !!! якщо було натиснуто Stars btm
                //db.Insert(userName, "Період роботи", "Блокировка сесии", DateTime.Now.ToString(), "30.12.1899");
                //DateTime avtoStart = new DateTime();
                avtoStart = DateTime.Now;
                sessionLock = true;
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock && Startbtm.Enabled != true && Pausebtm.Text == "Break")
            {
                sessionLock = false;
                // при розблокуванні - записати в БД перерву і час початку та час кінця   !!! якщо було натиснуто Stars btm
                var start = TimeSpan.Parse("13:00");
                var end = TimeSpan.Parse("14:00");
                //var now = DateTime.Now.TimeOfDay;
                var now = avtoStart.TimeOfDay;
                if (start <= now && now <= end)
                    db.Insert(userName, "Lunch", "(Automatically)", avtoStart.ToString(), DateTime.Now.AddSeconds(1).ToString(), "0");
                else
                    db.Insert(userName, "Break", "(Automatically)", avtoStart.ToString(), DateTime.Now.AddSeconds(1).ToString(), "0");
                //TimeSpan testtest = new TimeSpan();
                //testtest = startWorkTime - avtoStart;
                //MessageBox.Show(testtest.ToString());

            }
        }
        //SystemEvents.SessionSwitch += SESS;


        private void AutoRun(int isAutoRun)
        {
            if (isAutoRun == 1)
            {
                WshShell shell = new WshShell();

                //путь к ярлыку
                string autoRunPath = @"C:\Users\" + Environment.UserName + @"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\" + @"TimeManager.lnk";
                //string shortcutPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Блокнот.lnk";

                //создаем объект ярлыка
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(autoRunPath);

                //задаем свойства для ярлыка
                //описание ярлыка в всплывающей подсказке
                shortcut.Description = "Ярлык для TimeManager";
                //горячая клавиша
                //shortcut.Hotkey = "Ctrl+Shift+N";
                //путь к самой программе
                shortcut.TargetPath = @"S:\DATA\DebtCollectionProcess\Time Manager\" + @"TimeManager.exe";
                shortcut.WorkingDirectory = @"S:\DATA\DebtCollectionProcess\Time Manager\";   //путь к папке программи где БАЗА
                //Создаем ярлык
                shortcut.Save();
            }
        }






        private void SelectBase()
        {
            usersInfo = new List<UserInfo>();
            DataTable dt = db.Select("SELECT User_Name as Employee, Break_Type as Activity, Break_Notes as Comment, Start_Time as Start_time, End_Time as End_time FROM " + db.timeBase + " WHERE Start_Time BETWEEN '" + dtpInfoDate.Value.ToString("MM.dd.yyyy").Replace(".", "/") + "' AND '" + dtpInfoDateEnd.Value.AddDays(1).ToString("MM.dd.yyyy").Replace(".", "/") + "' " + (comboBoxUsers.SelectedIndex > 0 || comboBoxUsers.Text.Equals(userName) ? "AND User_Name=@uName" : "") + " ORDER BY Start_Time ASC, User_Name ASC", (comboBoxUsers.SelectedIndex > 0 || comboBoxUsers.Text.Equals(userName) ? comboBoxUsers.SelectedItem.ToString() : "")); // comboBoxUsers.Items[1].ToString()
            //DataTable dtt = db.Select("SELECT * FROM " + db.timeBase + "");
            //db.InsertTest(dtt);

            //DataTable dt = db.Select("SELECT User_Name as Працівник, Break_Type as Активність, Break_Notes as Коментар, Start_Time as Час_початку, End_Time as Час_завершення FROM " + db.timeBase + " WHERE Start_Time BETWEEN #" + dtpInfoDate.Value.ToString("MM.dd.yyyy").Replace(".", "/") + "# AND #" + dtpInfoDateEnd.Value.AddDays(1).ToString("MM.dd.yyyy").Replace(".", "/") + "#" + (comboBoxUsers.SelectedIndex > 0 || comboBoxUsers.Text.Equals(userName) ? " AND User_Name = '" + comboBoxUsers.SelectedItem + "'" : "") + " ORDER BY Start_Time ASC, User_Name ASC"); // comboBoxUsers.Items[1].ToString()
            //DataTable dt = db.Select("SELECT * FROM " + db.timeBase + " WHERE FORMAT(Start_Time,'DD.MM.YYYY') BETWEEN #" + dtpInfoDate.Value.ToString("MM.dd.yyyy").Replace(".", "/") + "# AND DateAdd ('d', 1, #" + dtpInfoDateEnd.Value.ToString("MM.dd.yyyy").Replace(".", "/") + "#)" + (comboBoxUsers.SelectedIndex > 0 || comboBoxUsers.Text.Equals(userName) ? " AND User_Name = '" + comboBoxUsers.SelectedItem + "'" : "") + " ORDER BY User_Name ASC"); // comboBoxUsers.Items[1].ToString()   DESC
            //DataTable dt = db.Select("SELECT * FROM " + db.timeBase + " WHERE FORMAT(Start_Time,'dd.mm.yyyy')=#" + dtpInfoDate.Value.ToString("MM.dd.yyyy").Replace(".", "/") + "#" + (comboBoxUsers.SelectedIndex > 0 || comboBoxUsers.Text.Equals(userName) ? " AND User_Name = '" + comboBoxUsers.SelectedItem + "'" : "") + " ORDER BY User_Name ASC"); // comboBoxUsers.Items[1].ToString()
            //dataGridView1.DataSource = dt;
            dgv.SetSourse(dt);
            string[] tempStrArr = string.Format("Work time report of employees;Data for the period from {0} to {1}", dtpInfoDate.Value.ToShortDateString(), dtpInfoDateEnd.Value.ToShortDateString()).Split(Convert.ToChar(";"));
            dgv.TitleArr = tempStrArr;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                UserInfo tempInfo = new UserInfo();
                if (usersInfo.Count > 0 && usersInfo.Count(el => el.Name.Equals(dt.Rows[i]["Employee"].ToString())) > 0)
                    tempInfo = usersInfo.Where(el => el.Name.Equals(dt.Rows[i]["Employee"].ToString())).First();
                else
                    usersInfo.Add(tempInfo);
                tempInfo.Name = dt.Rows[i]["Employee"].ToString();

                if (dt.Rows[i]["Activity"].ToString().Contains("Lunch"))
                {
                    tempInfo.BreakDinner += Convert.ToDateTime(dt.Rows[i]["End_time"].ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                    tempInfo.TypeDinner = dt.Rows[i]["Activity"].ToString();
                }
                if (dt.Rows[i]["Activity"].ToString().Contains("Break"))
                {
                    tempInfo.BreakPause += Convert.ToDateTime(dt.Rows[i]["End_time"].ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                    tempInfo.TypePause = dt.Rows[i]["Activity"].ToString();
                }
                if (dt.Rows[i]["Activity"].ToString().Contains("Meeting"))
                {
                    tempInfo.BreakMeeting += Convert.ToDateTime(dt.Rows[i]["End_time"].ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                    tempInfo.NoteMeeting += dt.Rows[i]["Comment"].ToString() + "; ";
                    tempInfo.TypeMeeting = dt.Rows[i]["Activity"].ToString();
                }
                if (dt.Rows[i]["Activity"].ToString().Contains("Study"))
                {
                    tempInfo.BreakStudy += Convert.ToDateTime(dt.Rows[i]["End_time"].ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                    tempInfo.NoteStudy += dt.Rows[i]["Comment"].ToString() + "; ";
                    tempInfo.TypeStudy = dt.Rows[i]["Activity"].ToString();
                }
                if (dt.Rows[i]["Activity"].ToString().Contains("Exit note"))
                {
                    tempInfo.BreakNote += Convert.ToDateTime(dt.Rows[i]["End_time"].ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                    tempInfo.TypeNote = dt.Rows[i]["Activity"].ToString();
                }
                if (dt.Rows[i]["Activity"].ToString().Contains("To the doctor"))
                {
                    tempInfo.BreakDoctor += Convert.ToDateTime(dt.Rows[i]["End_time"].ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                    tempInfo.TypeDoctor = dt.Rows[i]["Activity"].ToString();
                }
                if (dt.Rows[i]["Activity"].ToString().Contains("Period of work"))
                {
                    if (dt.Rows[i]["Comment"].ToString().Contains("(Work in progress..)"))
                    {
                        tempInfo.PeriodWork += Convert.ToDateTime(DateTime.Now.ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                        tempInfo.TypeWork = "Period of work";
                    }
                    else
                    {
                        tempInfo.PeriodWork += Convert.ToDateTime(dt.Rows[i]["End_time"].ToString()) - Convert.ToDateTime(dt.Rows[i]["Start_time"].ToString());
                        tempInfo.TypeWork = "Period of work";
                    }
                }



            }

        }

      
    }
    public static class ExtensionMethods
    {
        //потрібне для анулювання мерехтіння при прокручувані
        public static void DoubleBuffered(this AdvancedDataGridView dgv, bool setting)
        {
            try
            {
                typeof(AdvancedDataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, dgv, new object[] { true });
                //Type dgvType = dgv.GetType();
                //PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                //    BindingFlags.Instance | BindingFlags.NonPublic);
                //pi.SetValue(dgv, setting, null);
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            try { typeof(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, dgv, new object[] { true }); }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
    }





}
