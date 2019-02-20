using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Management;
using System.Diagnostics;
using System.IO;
using System.Threading;


namespace CheckBase
{
    public partial class Form1 : Form
    {
        private static System.Windows.Forms.Timer m_timer;
        private Int32 m_count;
        private string SPATH;
        private string SSRC;
        private bool WinID = true;
        private bool SQLID = false;
        private bool DoLogging = true;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr GetModuleHandle(string moduleName);

        [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr GetProcAddress(IntPtr hModule,
            [MarshalAs(UnmanagedType.LPStr)]string procName);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWow64Process(IntPtr hProcess, out bool wow64Process);

        static bool DoesWin32MethodExist(string moduleName, string methodName)
        {
            IntPtr moduleHandle = GetModuleHandle(moduleName);
            if (moduleHandle == IntPtr.Zero)
            {
                return false;
            }
            return (GetProcAddress(moduleHandle, methodName) != IntPtr.Zero);
        }

        static bool Is64BitOperatingSystem()
        {
            if (IntPtr.Size == 8)  // 64-bit programs run only on Win64
            {
                return true;
            }
            else  // 32-bit programs run on both 32-bit and 64-bit Windows
            {
                // Detect whether the current process is a 32-bit process 
                // running on a 64-bit system.
                bool flag;
                return ((DoesWin32MethodExist("kernel32.dll", "IsWow64Process") &&
                    IsWow64Process(GetCurrentProcess(), out flag)) && flag);
            }
        }

        
        private string PATH_property
        {
            get { return SPATH; }
            set { SPATH = value; }
        }
        private string SRC_property
        {
            get { return SSRC; }
            set { SSRC = value; }
        }

        private string ServerComboVal;

        public string ServerComboValue
        {
            get { return ServerComboVal; }
          
        }

        Process p;

        bool finish;

        public bool FinishErr
        {
            get { return finish; }
            set { finish = value; }
        }

       
        public Form1()
        {
            FinishErr = false;
            InitializeComponent();
           // PATH_property = PATH;
            SRC_property = @"c:";
            ReadParamNew();
            if (LocalCheckBox.Checked)
            {
                ServerCombo.Visible = true;
                NetServerBox.Visible = false;
            }
            if (NetCheckBox.Checked)
            {
                ServerCombo.Visible = false;
                NetServerBox.Visible = true;
            }



            ChooseBaselabel.Text = SRC_property + @"\MSDB\mssample.bak";
            if (ServerCombo.Items.Count > 0)
            {
                ServerCombo.SelectedIndex = 0;
            }

            ManagementScope scope = new ManagementScope("\\root\\cimv2");
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
            try
            {
                ManagementObjectCollection mobjects = searcher.Get();
                searcher.Dispose();
                foreach (ManagementObject mo in mobjects)
                {
                    CurrUserLabel.Text = "Текущий пользователь: " + mo["UserName"].ToString();
                }
         
                
            }
            catch (System.Exception)
            {
                CurrUserLabel.Text = "";
            }

        }

        private void WriteODBC()
        {
            try
            {
                if (Is64BitOperatingSystem())
                {
                    RegistryKey saveKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Wow6432Node\ODBC\ODBC.INI\" + BaseBox.Text);
                    saveKey.SetValue("Driver", Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\SQLSRV32.dll");
                    saveKey.SetValue("Description", "Подключение к программе \"Флагман\"");
                    saveKey.SetValue("Server", ServerCombo.SelectedItem);
                    saveKey.SetValue("QuotedId", "Yes");
                    saveKey.SetValue("Language", "русский");
                    saveKey.SetValue("LastUser", "SYSADM");
                    saveKey.SetValue("AutoTranslate", "No");
                    saveKey.SetValue("Database", BaseBox.Text);
                    saveKey.Close();
                    RegistryKey saveKey2 = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Wow6432Node\ODBC\ODBC.INI\ODBC Data Sources");
                    saveKey2.SetValue(BaseBox.Text, "SQL Server");
                    saveKey2.Close();
                }
                else
                {
                    RegistryKey saveKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\ODBC\ODBC.INI\" + BaseBox.Text);
                    saveKey.SetValue("Driver", Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\SQLSRV32.dll");
                    saveKey.SetValue("Description", "Подключение к программе \"Флагман\"");
                    saveKey.SetValue("Server", ServerCombo.SelectedItem);
                    saveKey.SetValue("QuotedId", "Yes");
                    saveKey.SetValue("Language", "русский");
                    saveKey.SetValue("LastUser", "SYSADM");
                    saveKey.SetValue("AutoTranslate", "No");
                    saveKey.SetValue("Database", BaseBox.Text);
                    saveKey.Close();
                    RegistryKey saveKey2 = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources");
                    saveKey2.SetValue(BaseBox.Text, "SQL Server");
                    saveKey2.Close();
                }

              
            }
            catch (Exception ex)
            {
                StatusLabel.Text = ex.Message;
                FinishErr = true;
            }
        }


        private void ReadParam()
        {
           // Читаем параметры из xml файла
            string Str;
            RegistryKey rk;
            try
            {
                ServerCombo.Items.Clear();
                XDocument xDocument = XDocument.Load(PATH_property + @"\param.xml");
                BaseBox.Text = xDocument.Element("Setup").Element("BASE").Value;
                UserBox.Text = xDocument.Element("Setup").Element("USER").Value;
                PassBox.Text = xDocument.Element("Setup").Element("PASS").Value;

                if (Is64BitOperatingSystem())
                {
                   

                    rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server");
                    String[] SubKeys = rk.GetValueNames();
                    bool found = false;
                    foreach (String SubKey in SubKeys)
                    {
                        if (SubKey == "InstalledInstances")
                            found = true;
                    }
                    if (found)
                    {

                        String[] instances = (String[])rk.GetValue("InstalledInstances");
                        if (instances.Length > 0)
                        {

                            foreach (String element in instances)
                            {
                                if (element == "MSSQLSERVER")
                                    Str = System.Environment.MachineName;
                                else
                                    Str = System.Environment.MachineName + @"\" + element;
                                ServerCombo.Items.Add(Str);
                            }
                        }
                    }

                    rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server");
                    SubKeys = rk.GetValueNames();
                    found = false;
                    foreach (String SubKey in SubKeys)
                    {
                        if (SubKey == "InstalledInstances")
                            found = true;
                    }
                    if (found)
                    {

                        String[] instances = (String[])rk.GetValue("InstalledInstances");
                        if (instances.Length > 0)
                        {
                            foreach (String element in instances)
                            {
                                if (element == "MSSQLSERVER")
                                    Str = System.Environment.MachineName;
                                else
                                    Str = System.Environment.MachineName + @"\" + element;
                                ServerCombo.Items.Add(Str);
                            }
                        }
                    }

                }
                else
                {
                 
                    rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server");
                    String[] SubKeys = rk.GetValueNames();
                     bool found = false;
                    foreach (String SubKey in SubKeys)
                    {
                        if (SubKey == "InstalledInstances")
                            found = true;
                    }
                    if (found)
                    {

                        String[] instances = (String[])rk.GetValue("InstalledInstances");
                        if (instances.Length > 0)
                        {
                            foreach (String element in instances)
                            {
                                if (element == "MSSQLSERVER")
                                    Str = System.Environment.MachineName;
                                else
                                    Str = System.Environment.MachineName + @"\" + element;
                                ServerCombo.Items.Add(Str);
                            }
                        }
                    }
                }
              
               



            }
            catch (Exception ex)
            {

                StatusLabel.Text = ex.Message;
                FinishErr = true;
            }
        }

        private void ReadParamNew()
        {
            // Читаем параметры из xml файла
            string Str;
            RegistryKey rk;
            try
            {
                ServerCombo.Items.Clear();
               
                BaseBox.Text = "ISMS";
                UserBox.Text = "sa";
                PassBox.Text = "1qaz2wsx";

                if (Is64BitOperatingSystem())
                {


                    rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server");
                    String[] SubKeys = rk.GetValueNames();
                    bool found = false;
                    foreach (String SubKey in SubKeys)
                    {
                        if (SubKey == "InstalledInstances")
                            found = true;
                    }
                    if (found)
                    {

                        String[] instances = (String[])rk.GetValue("InstalledInstances");
                        if (instances.Length > 0)
                        {

                            foreach (String element in instances)
                            {
                                if (element == "MSSQLSERVER")
                                    Str = System.Environment.MachineName;
                                else
                                    Str = System.Environment.MachineName + @"\" + element;
                                ServerCombo.Items.Add(Str);
                            }
                        }
                    }

                    rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server");
                    SubKeys = rk.GetValueNames();
                    found = false;
                    foreach (String SubKey in SubKeys)
                    {
                        if (SubKey == "InstalledInstances")
                            found = true;
                    }
                    if (found)
                    {

                        String[] instances = (String[])rk.GetValue("InstalledInstances");
                        if (instances.Length > 0)
                        {
                            foreach (String element in instances)
                            {
                                if (element == "MSSQLSERVER")
                                    Str = System.Environment.MachineName;
                                else
                                    Str = System.Environment.MachineName + @"\" + element;
                                ServerCombo.Items.Add(Str);
                            }
                        }
                    }

                }
                else
                {

                    rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server");
                    String[] SubKeys = rk.GetValueNames();
                    bool found = false;
                    foreach (String SubKey in SubKeys)
                    {
                        if (SubKey == "InstalledInstances")
                            found = true;
                    }
                    if (found)
                    {

                        String[] instances = (String[])rk.GetValue("InstalledInstances");
                        if (instances.Length > 0)
                        {
                            foreach (String element in instances)
                            {
                                if (element == "MSSQLSERVER")
                                    Str = System.Environment.MachineName;
                                else
                                    Str = System.Environment.MachineName + @"\" + element;
                                ServerCombo.Items.Add(Str);
                            }
                        }
                    }
                }





            }
            catch (Exception ex)
            {

                StatusLabel.Text = ex.Message;
                FinishErr = true;
            }
        }



        private static System.Windows.Forms.Timer GetTimer()
        {
            if (m_timer == null)
            {
                m_timer = new System.Windows.Forms.Timer();
                m_timer.Interval = 100;
                m_timer.Start();
            }

            return m_timer;
        }

        void WorkingForm_Tick(object sender, EventArgs e)
        {
            m_count++;
            m_count = m_count % progressBar.Maximum + progressBar.Minimum;
            progressBar.Value = m_count;
        }


        private void ConnectButton_Click(object sender, EventArgs e)
        {
            if (LocalCheckBox.Checked)
            {

                if (backgroundWorker.IsBusy != true)
                {
                    // Start the asynchronous operation.
                    if (ServerCombo.SelectedItem != null)
                    {
                        Okbutton.Enabled = false;
                        progressBar.Visible = true;
                        GetTimer().Tick += WorkingForm_Tick;
                        ServerComboVal = ServerCombo.SelectedItem.ToString();
                        backgroundWorker.RunWorkerAsync();
                    }
                    else
                    {
                        StatusLabel.Text = "Не выбран экземпляр сервера!..";
                    }

                }
            }
            else
            {
                if (backgroundWorker.IsBusy != true)
                {
                    // Start the asynchronous operation.
                    if (NetServerBox.Text != string.Empty || NetServerBox.Text != "")
                    {
                        Okbutton.Enabled = false;
                        progressBar.Visible = true;
                        GetTimer().Tick += WorkingForm_Tick;
                        ServerComboVal = NetServerBox.Text;
                        backgroundWorker.RunWorkerAsync();
                    }
                    else
                    {
                        StatusLabel.Text = "Не выбран экземпляр сервера!..";
                    }

                }
            }
         
               
           
        }


        private void PreInstallScript()
        {

            SqlConnection Sqlconn = new SqlConnection();
            try
            {
                if (WinID)
                {

                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master; Integrated Security=True;";
                }
                if (SQLID)
                {
                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master; User ID=" + UserBox.Text + ";Password=" + PassBox.Text;

                }

                Sqlconn.Open();
                if (Sqlconn.State == System.Data.ConnectionState.Open)
                {

                    // Предварительные скрипты
                    SqlCommand cmd1 = Sqlconn.CreateCommand();
                    cmd1.CommandText = @"use master
                                         if (select count(sid) from master.dbo.syslogins where name = 'SYSADM') = 0
                                          begin
                                            CREATE LOGIN SYSADM  WITH PASSWORD ='SYSADM', CHECK_POLICY =  OFF
                                            CREATE USER SYSADM
                                            exec master.dbo.sp_addsrvrolemember 'SYSADM', 'sysadmin'
                                          end
                                         
                                         ";
                    cmd1.CommandType = CommandType.Text;
                    cmd1.ExecuteNonQuery();
                    // создадим роль FLAGMAN_USER
                    SqlCommand cmd2 = Sqlconn.CreateCommand();
                    cmd2.CommandText = @"if not exists ( select uid from master.dbo.sysusers where name = 'FLAGMAN_USER' and issqlrole = 1 )
                                         begin
	                                        CREATE ROLE FLAGMAN_USER	
                                         end
                                         ";
                    cmd2.CommandType = CommandType.Text;
                    cmd2.ExecuteNonQuery();
                    // включим SYSADM'а в роль
                    SqlCommand cmd3 = Sqlconn.CreateCommand();
                    cmd3.CommandText = @"exec sp_addrolemember 'FLAGMAN_USER', 'SYSADM'
                                         ";
                    cmd3.CommandType = CommandType.Text;
                    cmd3.ExecuteNonQuery();


                    // проверка на необходимость создания backup устройства
                    SqlCommand cmd4 = Sqlconn.CreateCommand();
                    cmd4.CommandText = @"if exists(select * from master.dbo.sysdevices where name = 'ISMS_DEV')
                                         begin
	                                        exec sp_dropdevice 'ISMS_DEV'
                                         end
                                         ";
                    cmd4.CommandType = CommandType.Text;
                    cmd4.ExecuteNonQuery();

                    //add 2005 support
                    SqlCommand cmd5 = Sqlconn.CreateCommand();
                    cmd5.CommandText = @"GRANT VIEW ANY DEFINITION TO public
                                         ";
                    cmd5.CommandType = CommandType.Text;
                    cmd5.ExecuteNonQuery();







                }

            }
            catch (Exception ee)
            {
                StatusLabel.Text = ee.Message;
                           
            }
            finally
            {
                Sqlconn.Close();
                Sqlconn.Dispose();
            }

        
        
        
        }




        private void Okbutton_Click(object sender, EventArgs e)
        {
            // Предварительные скрипты
            PreInstallScript();



            // Записываем параметры
            string BPath;
            SqlConnection Sqlconn = new SqlConnection();
            //XDocument xDocument = XDocument.Load(PATH_property + @"\param.xml");
            // определяем путь к базам данных
            try
            {
                if (WinID)
                {

                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master;  Integrated Security=True;";
                }
                if (SQLID)
                {
                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master; User ID=" + UserBox.Text + ";Password=" + PassBox.Text;
                  
                }
                
                Sqlconn.Open();
                if (Sqlconn.State == System.Data.ConnectionState.Open)
                {

                   

                    SqlCommand cmd = Sqlconn.CreateCommand();
                    cmd.CommandText = @"select top 1 physical_name from sys.database_files";
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                    SqlDataReader r = cmd.ExecuteReader();
                    r.Read();
                    BPath =  r[0].ToString().Replace("master.mdf", "");
                    r.Close();
                    Sqlconn.Close();
                    Sqlconn.Dispose();

                 
                     string text = string.Empty;
                    if (DoLogging)
                    {
                        if (WinID)
                        {

                            text = @" -S " + ServerComboValue + @" -E -Q ""RESTORE DATABASE [" +
                           BaseBox.Text + @"] FROM  DISK = N'" + ChooseBaselabel.Text + @"' WITH  FILE = 1,  MOVE N'ISMS_Data' TO N'" + BPath + BaseBox.Text + @".mdf',  MOVE N'ISMS_Log' TO N'" + BPath + BaseBox.Text + @".ldf',  NOUNLOAD,  STATS = 10""" + @" -u -o c:\osql.log";
                        }
                        if (SQLID)
                        {
                            text = " -U " + UserBox.Text + " -P " + PassBox.Text + @" -S " + ServerComboValue + @" -Q ""RESTORE DATABASE [" +
                          BaseBox.Text + @"] FROM  DISK = N'" + ChooseBaselabel.Text + @"' WITH  FILE = 1,  MOVE N'ISMS_Data' TO N'" + BPath + BaseBox.Text + @".mdf',  MOVE N'ISMS_Log' TO N'" + BPath + BaseBox.Text + @".ldf',  NOUNLOAD,  STATS = 10""" + @" -u -o c:\osql.log";
 
                        }
                    }
                    else
                    {
                        if (WinID)
                        {
                            text = @" -S " + ServerComboValue + @" -E -Q ""RESTORE DATABASE [" +
                           BaseBox.Text + @"] FROM  DISK = N'" + ChooseBaselabel.Text + @"' WITH  FILE = 1,  MOVE N'ISMS_Data' TO N'" + BPath + BaseBox.Text + @".mdf',  MOVE N'ISMS_Log' TO N'" + BPath + BaseBox.Text + @".ldf',  NOUNLOAD,  STATS = 10""";
                        }
                        if (SQLID)
                        {
                            text = " -U " + UserBox.Text + " -P " + PassBox.Text + @" -S " + ServerComboValue + @" -Q ""RESTORE DATABASE [" +
                          BaseBox.Text + @"] FROM  DISK = N'" + ChooseBaselabel.Text + @"' WITH  FILE = 1,  MOVE N'ISMS_Data' TO N'" + BPath + BaseBox.Text + @".mdf',  MOVE N'ISMS_Log' TO N'" + BPath + BaseBox.Text + @".ldf',  NOUNLOAD,  STATS = 10""";
 
                        }
                    }

                   

                    // Стартуем osql
                    ProcessStartInfo info = new ProcessStartInfo("osql.exe", text);

                    info.UseShellExecute = false;
                    info.CreateNoWindow = true;
                    info.WindowStyle = ProcessWindowStyle.Hidden;
                    info.RedirectStandardOutput = true;
                    p = new Process();
                    p.StartInfo = info;


                    progressBar.Visible = true;
                    GetTimer().Tick += WorkingForm_Tick;

                    if (backgroundWorkerOSQL.IsBusy != true)
                    {
                        // Start the asynchronous operation.
                        backgroundWorkerOSQL.RunWorkerAsync();
                    }

                    
                                      
                }
            }
            catch (Exception ee)
            {
                StatusLabel.Text = ee.Message;
                FinishErr = true;
                //xDocument.Element("Setup").Element("STATUS").SetValue("0");
                //xDocument.Save(PATH_property + @"\param.xml");
            }

      
           
        }

       

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            SqlConnection Sqlconn = new SqlConnection();
            try
            {

                if (WinID)
                {
                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master; Integrated Security=True;";
                    
                }
                if (SQLID)
                {
                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master; User ID=" + UserBox.Text + ";Password=" + PassBox.Text;
                  
                }
                    Sqlconn.Open();
                    if (Sqlconn.State == System.Data.ConnectionState.Open)
                    {
                        SqlCommand cmd = Sqlconn.CreateCommand();
                        cmd.CommandText = @"select COUNT(name) from sys.databases WHERE name = '" + BaseBox.Text + "'";
                        cmd.CommandType = CommandType.Text;
                        cmd.ExecuteNonQuery();
                        SqlDataReader r = cmd.ExecuteReader();
                        if (r.HasRows)
                        {
                            while (r.Read())
                            {
                                if (r[0].ToString() != "0")
                                    StatusLabel.Text = "База данных " + BaseBox.Text + " существует! Введите другое имя базы.";
                                else
                                {
                                    StatusLabel.Text = @"Проверка параметров завершена успешно! Нажмите <<Продолжить>>...";
                                    // Выполняем предварительные операции
                                    Okbutton.Enabled = true;
                                }
                            }


                        }

                        r.Close();
                        Sqlconn.Close();
                        Sqlconn.Dispose();
                    }
                }
                catch (Exception ee)
                {
                    StatusLabel.Text = ee.Message;
                    FinishErr = true;
                   

                }
              









        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.Visible = false;
            GetTimer().Tick -= WorkingForm_Tick;
        }

        private void WinCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (WinCheckBox.Checked)
            {
                SQLCheckBox.CheckState = CheckState.Unchecked;
                UserLabel.Visible = false;
                Passlabel.Visible = false;
                UserBox.Visible = false;
                PassBox.Visible = false;
                CurrUserLabel.Visible = true;
                WinID = true;
                SQLID = false;
            }
        }

        private void SQLCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (SQLCheckBox.Checked)
            {
               WinCheckBox.CheckState = CheckState.Unchecked;
                UserLabel.Visible = true;
                Passlabel.Visible = true;
                UserBox.Visible = true;
                PassBox.Visible = true;
                CurrUserLabel.Visible = false;
                SQLID = true;
                WinID = false;
            }
        }

        private void OpenDiag1_Click(object sender, EventArgs e)
        {
            openFileDialog.ShowDialog();
        }

        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            ChooseBaselabel.Text = openFileDialog.FileName; 
        }

        private void checkLog_CheckedChanged(object sender, EventArgs e)
        {
            if (checkLog.Checked)
            {
                DoLogging = true;
                LogPath.Visible = true;
            }
            else
            {
                DoLogging = false;
                LogPath.Visible = false;
            }
        }

        private void backgroundWorkerOSQL_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                p.Start();
                StreamReader rr = p.StandardOutput;
                StatusLabel.Text = rr.ReadToEnd();
            }
            catch (Exception ee)
            {
                StatusLabel.Text = ee.Message;
                FinishErr = true;
            }

        }

        private void backgroundWorkerOSQL_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar.Visible = false;
            GetTimer().Tick -= WorkingForm_Tick;
            AfterInstallScripts(BaseBox.Text);
           StatusLabel.Text = "Установка БД прошла успешно!.. Прописываем ODBC";
             // Прописываем ODBC
            Thread.Sleep(3000);
            WriteODBC();
            Okbutton.Text = "Выход";
            ConnectButton.Enabled = false;
            this.Okbutton.Click -= new System.EventHandler(this.Okbutton_Click);
            this.Okbutton.Click += new System.EventHandler(this.Okbutton_Quit);
            if (!FinishErr)
            {
                StatusLabel.Text = "ODBC успешно прописан!.. Нажмите <<Выход>> для закрытия программы";
            }
            else
            {
                StatusLabel.Text = "В процессе установки возникли ошибки!.. Нажмите <<Выход>> для закрытия программы";
            }
           
        }

        private void Okbutton_Quit(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void LocalCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (LocalCheckBox.Checked)
            {
                NetCheckBox.CheckState = CheckState.Unchecked;
                ServerCombo.Visible = true;
                NetServerBox.Visible = false;
                
            }
        }

        private void NetCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (NetCheckBox.Checked)
            {
                LocalCheckBox.CheckState = CheckState.Unchecked;
                ServerCombo.Visible = false;
                NetServerBox.Visible = true;
            }

        }


        private void AfterInstallScripts(string BaseName)
        {
            SqlConnection Sqlconn = new SqlConnection();
            try
            {
                if (WinID)
                {

                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master; Integrated Security=True;";
                }
                if (SQLID)
                {
                    Sqlconn.ConnectionString = "Data Source=" + ServerComboValue + ";Initial Catalog=master; User ID=" + UserBox.Text + ";Password=" + PassBox.Text;

                }

                Sqlconn.Open();
                if (Sqlconn.State == System.Data.ConnectionState.Open)
                {
                    SqlCommand cmd1 = Sqlconn.CreateCommand();
//                    cmd1.CommandText = @"use " + BaseName + @" 
//                                         exec sp_changedbowner 'SYSADM'
//                                         ";
                    cmd1.CommandText = @"use " + BaseName;
                    cmd1.CommandType = CommandType.Text;
                    cmd1.ExecuteNonQuery();
                    cmd1.CommandText = @" exec sp_changedbowner 'SYSADM'"; 
                                       
                    cmd1.CommandType = CommandType.Text;
                    cmd1.ExecuteNonQuery();
                    cmd1.CommandText = @"ALTER DATABASE " + BaseName + @" SET TRUSTWORTHY ON";
                    cmd1.CommandType = CommandType.Text;
                    cmd1.ExecuteNonQuery();



                }



            }
            catch (Exception ee)
            {
                StatusLabel.Text = ee.Message;
                FinishErr = true;

            }
            finally
            {
                Sqlconn.Close();
                Sqlconn.Dispose();
            }

        }

        

        
    }
}
