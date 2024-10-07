using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Threading;
using System.Windows.Forms;


namespace ControlGUI
{
    public partial class Form1 : Form
    {
        //string[] ArrayComPortsNames = null;
        string[] portnames = null;
        string[] portnamesFilter = null;
        public string serialWriteData = "";
        public string redDuty = "0";
        public string pwnFrequency = "1000";
        public short OnOff = 0;
        public bool D_all = false;
        public bool D_all_Lame = false;
        public bool pwrScale = false;
        public bool freqScale = false;
        public bool FileIObool = false;
        public bool firstRun_150Log = true;
        public bool firstRun_LameLog = true;
        public DateTime rtcDateTime;
        public string rtcDateTimeString;

        SerialPort _serialPort;
        SerialPort _serialPortCtrl;
        private bool firstRun = true;
        private bool firstRun_Lame = true;
        private bool formOpen = true;

        // delegate is used to write to a UI control from a non-UI thread
        private delegate void SetTextDeleg(string text);

        // delegate is used to write to a UI control from a non-UI thread
        private delegate void SetDumpDeleg(string text);

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                this.ControlBox = false;
                int index = -1;
                btnPortState150Win.BackColor = Color.Coral;
                btnPortState1.BackColor = Color.Coral;

                //Com Ports C:\Users\diblu\source\repos\Serial Scanner

                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Caption like '%(COM%'"))
                {
                    portnames = SerialPort.GetPortNames();
                    var ports = searcher.Get().Cast<ManagementBaseObject>().ToList().Select(p => p["Caption"].ToString());

                    var portList = portnames.Select(n => n + " - " + ports.FirstOrDefault(s => s.Contains(n))).ToList();

                    foreach (string s in portList)
                    {
                        if ((s.Contains("Labs")))
                        {
                            index += 1;
                            cboPorts150WIN.Items.Add(s);
                            cboPorts1.Items.Add(s);
                            Console.WriteLine(s);
                        }
                    }
                    foreach (string s in ports)
                    {
                        Console.WriteLine(s);
                    }
                    foreach (string s in portList)
                    {
                        Console.WriteLine(s);
                    }
                    //cboPorts150WIN.Text = "COM";
                }
            }
            catch { }
        }

        public double MapValue(double a0, double a1, double b0, double b1, double a)
        {
            return b0 + (b1 - b0) * ((a - a0) / (a1 - a0));
        }


        private void BtnStart_Click(object sender, EventArgs e)
        {
            // Makes sure serial port is open before trying to write
            try
            {
                D_all = true;
                if (!_serialPort.IsOpen)
                    _serialPort.Open();
                string serialWriteData = textBox2.Text;
                serialWrite150WIN(serialWriteData + "\r\n");
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error Sending Data!");
            }
        }

        private void BtnSendCtrl_Click(object sender, EventArgs e)
        {
            // Makes sure serial port is open before trying to write
            try
            {
                if (!_serialPortCtrl.IsOpen)
                    _serialPortCtrl.Open();
                string serialWriteData = textBox1.Text;
                serialWriteLaserMeister(serialWriteData + "\r\n");
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error Sending Data!");
            }
        }


        void sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string data = _serialPort.ReadLine();
                this.BeginInvoke(new SetTextDeleg(si_DataReceived), new object[] { data });
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error opening/writing to ESP32 serial port :: " + ex.Message, "Error!");
            }
        }

        private void si_DataReceived(string data)
        {
            try
            {
                tb150WinRecieved.Text = data;
                tbComm.AppendText(data + "\r\n");
                tbComm.AppendText("\r\n");

                if (D_all == true)
                {
                    if (firstRun)
                    {
                        try
                        {
                            System.Data.OleDb.OleDbConnection MyConnection;
                            System.Data.DataSet DtSet;
                            System.Data.OleDb.OleDbDataAdapter MyCommand;
                            // This will get the current WORKING directory (i.e. \bin\Debug)
                            string workingDirectory = Environment.CurrentDirectory;
                            string[] paths = { workingDirectory, "150WIN command list.xls" };
                            string fullPath = Path.Combine(paths);
                            MyConnection = new System.Data.OleDb.OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + "; Extended Properties=Excel 8.0;");
                            MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
                            MyCommand.TableMappings.Add("Table", "Net-informations.com");
                            DtSet = new System.Data.DataSet();
                            MyCommand.Fill(DtSet);
                            dataGridView1.DataSource = DtSet.Tables[0];
                            MyConnection.Close();
                            D_all = false;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("150WIN table error");
                        }
                    }

                    string LASact = " ";
                    
                    string boardid = " ";
                    string dspvol = " ";
                    string dsptim = " ";
                    string Ddspgo = " ";
                    string Ddpstr = " ";
                    string Ddpgud = " ";
                    string Dvslsv = " ";
                    string Dsyrsv = " ";
                    string Dvslspv = " ";
                    string Dsyrpv = " ";
                    string Dlmpbr = " ";
                    string Dsyrch = " ";
                    string Dcbump = " ";
                    string LASset = " ";
                    string PWMrms = " ";
                    string keyClosed = " ";
                    string boxClosed = " ";
                    string WinAlarm = " ";
                    string State = " ";
                    int LD1offset = 0;
                    int LD2offset = 0;
                    int LD3offset = 0;
                    int waterCool = 0;
                    string SerNum;
                    string HdWver;
                    string SfWver;
                    string BldDay;

                    string dataSub;
                    int start = 0;
                    int at;
                    int at1;
                    int end = 0;
                    int count = 0;
                    int[] startIndex = new int[60];
                    int[] commaIndex = new int[60];
                    int i = 0;
                    end = data.Length;

                    at = 0;
                    at1 = 0;

                    while ((start < end) && (at > -1))
                    {
                        // start+count must be a position within -str-.
                        count = end - start;
                        at = data.IndexOf("F,", start, count);
                        start = at + 2;
                        count = end - start;
                        at1 = data.IndexOf(")", start, count);
                        if (at == -1) break;
                        if (at1 == -1) break;
                        Console.Write("{0} ", at);
                        Console.WriteLine();
                        Console.Write("{0} ", at1);

                        //startIndex[i] = at;
                        dataSub = data.Substring(at + 2, at1 - at - 2);
                        Console.Write(i);
                        Console.WriteLine("   {0}", dataSub);

                        switch (i)
                        {
                            case 0:
                                Console.WriteLine("Case 0");
                                boardid = dataSub;
                                break;
                            case 1:
                                Console.WriteLine("Case 1");
                                dspvol = dataSub;
                                break;
                            case 2:
                                Console.WriteLine("Case 2");
                                dsptim = dataSub;
                                break;
                            case 3:
                                Console.WriteLine("Case 3");
                                Ddspgo = dataSub;
                                break;
                            case 4:
                                Console.WriteLine("Case 4");
                                Ddpstr = dataSub;
                                break;
                            case 5:
                                Console.WriteLine("Case 5");
                                Ddpgud = dataSub;
                                break;
                            case 6:
                                Console.WriteLine("Case 6");
                                Dvslsv = dataSub;
                                break;
                            case 7:
                                Console.WriteLine("Case 7");
                                Dsyrsv = dataSub;
                                break;
                            case 8:
                                Console.WriteLine("Case 8");
                                Dvslspv = dataSub;
                                break;
                            case 12:
                                Console.WriteLine("Case 12");
                                Dsyrpv = dataSub;
                                break;
                            case 13:
                                Console.WriteLine("Case 13");
                                Dlmpbr = dataSub;
                                break;
                            case 14:
                                Console.WriteLine("Case 14");
                                Dsyrch = dataSub;
                                break;
                            case 15:
                                Console.WriteLine("Case 15");
                                Dcbump = dataSub;
                                break;
                            case 30:
                                Console.WriteLine("Case 30");
                                //WinAlarm = dataSub;
                                //if(WinAlarm == "off"){ btnAlarm.BackColor = Color.LightGreen; }
                                //else{ btnAlarm.BackColor = Color.Coral; }
                                break;
                            case 31:
                                Console.WriteLine("Case 31");
                                WinAlarm = dataSub;
                                if (WinAlarm == "off") { btnAlarm.BackColor = Color.Gray;
                                    btnAlarm.Enabled = false;
                                }
                                else { btnAlarm.BackColor = Color.Coral;
                                    btnAlarm.Enabled = true;
                                }
                                break;
                            case 35:
                                Console.WriteLine("Case 35");
                                PWMrms = dataSub;
                                break;
                            case 36:
                                Console.WriteLine("Case 36");
                                LASset = dataSub;
                                break;
                            case 39:
                                Console.WriteLine("Case 39");
                                LASact = dataSub;
                                break;
                            case 40:
                                Console.WriteLine("Case 40");
                                keyClosed = dataSub;
                                break;
                            case 41:
                                Console.WriteLine("Case 41");
                                boxClosed = dataSub;
                                break;
                            case 54:
                                Console.WriteLine("Case 54");
                                State = dataSub;
                                int result = Int32.Parse(dataSub);
                                break;
                            case 55:
                                Console.WriteLine("Case 55");
                                LD1offset = Int32.Parse(dataSub);
                                break;
                            case 56:
                                Console.WriteLine("Case 56");
                                LD2offset = Int32.Parse(dataSub);
                                break;
                            case 57:
                                Console.WriteLine("Case 57");
                                LD3offset = Int32.Parse(dataSub);
                                break;
                            case 67:
                                Console.WriteLine("Case 67");
                                waterCool = Int32.Parse(dataSub);
                                break;
                            case 68:
                                Console.WriteLine("Case 68");
                                tbSerial.Text = dataSub;
                                break;
                            case 69:
                                Console.WriteLine("Case 69");
                                tbHWver.Text = dataSub;
                                break;
                            case 70:
                                Console.WriteLine("Case 70");
                                tbSWver.Text = dataSub;
                                break;
                            case 71:
                                Console.WriteLine("Case 71");
                                tbBldDay.Text = dataSub;
                                break;
                            case 72:
                                Console.WriteLine("Case 72");
                                //tbBldDay.Text = dataSub;
                                if (dataSub == "0") { cbChillerPwr.Checked = false; }
                                else { cbChillerPwr.Checked = true; }
                                break;
                            case 73:
                                Console.WriteLine("Case 73");
                                break;
                            case 74:
                                Console.WriteLine("Case 74");
                                
                                break;
                            case 75:
                                Console.WriteLine("Case 75");
                                numChillSet.Value = (decimal)(double.Parse(dataSub) / 10);
                                break;
                            case 76:
                                Console.WriteLine("Case 76");
                                break;
                        }
                        this.dataGridView1.Rows[i].Cells[4].Value = dataSub;
                        i++;
                    }
                    Console.WriteLine();
                    tbboardid.Text = boardid;
                    tbdspvol.Text = dspvol;
                    tbDdpgud.Text = Ddpgud;
                    tbDvslsv.Text = Dvslsv;
                    tbKey.Text = keyClosed;
                    tbBox.Text = boxClosed;
                    tbState.Text = State;
                    //if (firstRun) cbActivePoll.Checked = true;
                    firstRun = false;
                    if (cbDatalog150.Checked == true)
                    {
                        FileIO150 dataLog = new FileIO150();
                        if (firstRun_150Log) { FileIO150.fullFileName = "150 datalog"; dataLog.fileCreate(); firstRun_150Log = false; }
                        FileIO150.Dsyrpv = Dsyrpv;
                        FileIO150.Dlmpbr = Dlmpbr;
                        FileIO150.Dsyrch = Dsyrch;
                        FileIO150.Dcbump = Dcbump;
                        FileIO150.PWMrms = PWMrms;
                        FileIO150.LASset = LASset;
                        //FileIO150.LASact = LASact;
                        FileIO150.boardid = boardid;
                        FileIO150.dspvol = dspvol;
                        FileIO150.dsptim = dsptim;
                        FileIO150.Ddspgo = Ddspgo;
                        FileIO150.Ddpstr = Ddpstr;
                        FileIO150.Ddpgud = Ddpgud;
                        FileIO150.Dvslsv = Dvslsv;
                        FileIO150.Dsyrsv = Dsyrsv;
                        FileIO150.Dvslspv = Dvslspv;
                        dataLog.fileWrite();
                    }
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error opening/writing to serial port :: " + ex.Message, "Error!");
            }
        }

        private void sp_DataReceivedCtrl(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                string data1 = _serialPortCtrl.ReadLine();
                this.BeginInvoke(new SetDumpDeleg(si_DataReceivedCtrl), new object[] { data1 });
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error opening/writing to TEST serial port :: " + ex.Message, "Error!");
            }
        }

        private void si_DataReceivedCtrl(string data1)
        {
            try
            {
                tbLaserMeisterRecieved.Text = data1;
                tbCommCtrl.AppendText(data1 + "\r\n");
                tbCommCtrl.AppendText("\r\n");

                if (D_all_Lame == true)
                {
                    if (firstRun_Lame)
                    {
                        try
                        {
                            System.Data.OleDb.OleDbConnection MyConnection;
                            System.Data.DataSet DtSet;
                            System.Data.OleDb.OleDbDataAdapter MyCommand;
                            // This will get the current WORKING directory (i.e. \bin\Debug)
                            string workingDirectory = Environment.CurrentDirectory;
                            string[] paths = { workingDirectory, "Lame command list.xls" };
                            // or: Directory.GetCurrentDirectory() gives the same result
                            // This will get the current PROJECT directory
                            //string projectDirectory = Application.StartupPath;
                            //string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
                            //string[] paths = { projectDirectory, "Lame command list.xls" };

                            string fullPath = Path.Combine(paths);
                            MyConnection = new System.Data.OleDb.OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fullPath + "; Extended Properties=Excel 8.0;");
                            MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection);
                            MyCommand.TableMappings.Add("Table", "Net-informations.com");
                            DtSet = new System.Data.DataSet();
                            MyCommand.Fill(DtSet);
                            dataGridView2.DataSource = DtSet.Tables[0];
                            MyConnection.Close();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("LAME table error");
                            Console.WriteLine("IOException source: {0}", ex.Source);
                            //MessageBox.Show("Error opening/writing to serial port :: " + ex.Message, "Error!");
                        }
                    }

                    string Ilokot = " ";
                    string AlrmIn = " ";
                    string Imode = " ";
                    string Pmode = " ";
                    string LASenb = " ";
                    string PWMdty = " ";
                    string PWMfrq = " ";
                    string LASset = " ";
                    string TEMPmn = " ";
                    string LASact = " ";
                    string DACrun = " ";
                    string DACstp = " ";
                    int PWMdtyInt = 0;
                    int PWMfrqInt = 0;
                    int DACrunInt = 0;
                    int DACstpInt = 0;
                    double LASsetDbl = 0;

                    string dataSub;
                    int start = 0;
                    int at;
                    int at1;
                    int end = 0;
                    int count = 0;
                    int[] startIndex = new int[60];
                    int[] commaIndex = new int[60];
                    int i = 0;
                    end = data1.Length;

                    at = 0;
                    at1 = 0;

                    while ((start < end) && (at > -1))
                    {
                        // start+count must be a position within -str-.
                        count = end - start;
                        at = data1.IndexOf("A:", start, count);
                        start = at + 2;
                        count = end - start;
                        at1 = data1.IndexOf("A:", start, count);
                        if (at == -1) break;
                        if (at1 == -1) break;
                        Console.Write("{0} ", at);
                        Console.WriteLine();
                        Console.Write("{0} ", at1);

                        //startIndex[i] = at;
                        dataSub = data1.Substring(at + 2, at1 - at - 2);
                        Console.Write(i);
                        Console.WriteLine("   {0}", dataSub);
                        switch (i)
                        {
                            case 0:
                                Console.WriteLine("Case 0");
                                Ilokot = dataSub;
                                if(Ilokot == "open") { cbInterlock.Checked = false; }
                                else { cbInterlock.Checked = true; }
                                break;
                            case 1:
                                Console.WriteLine("Case 1");
                                AlrmIn = dataSub;
                                if (AlrmIn == "off") { btnAlarmLame.BackColor = Color.Gray;
                                    btnAlarmLame.Enabled = false;
                                }
                                else { btnAlarmLame.BackColor = Color.Coral;
                                    btnAlarmLame.Enabled = true;
                                }
                                break;
                            case 2:
                                Console.WriteLine("Case 2");
                                Imode = dataSub;
                                if (Imode == "off") { cbIOmode.Checked = false; }
                                else { cbIOmode.Checked = true; }
                                break;
                            case 3:
                                Console.WriteLine("Case 3");
                                Pmode = dataSub;
                                if (Pmode == "off") { cbPWMmode.Checked = false; }
                                else { cbPWMmode.Checked = true; }
                                break;
                            case 4:
                                Console.WriteLine("Case 4");
                                LASenb = dataSub;
                                if (LASenb == "disabled") { btnLaserEn.Checked = false; }
                                else { btnLaserEn.Checked = true; }
                                break;
                            case 5:
                                Console.WriteLine("Case 5");
                                PWMdty = dataSub;
                                PWMdtyInt = Int32.Parse(PWMdty);
                                zDrpDnSpd.Value = PWMdtyInt;
                                break;
                            case 6:
                                Console.WriteLine("Case 6");
                                PWMfrq = dataSub;
                                PWMfrqInt = Int32.Parse(PWMfrq);
                                zDrpUpSpeed.Value = PWMfrqInt;
                                break;
                            case 7:
                                Console.WriteLine("Case 8");
                                LASset = dataSub;
                                LASsetDbl = double.Parse(LASset);
                                zDrpDist.Value = (decimal)(LASsetDbl/10);
                                break;
                            case 8:
                                Console.WriteLine("Case 7");
                                TEMPmn = dataSub;
                                tbTempMain.Text = TEMPmn;
                                break;
                            case 9:
                                Console.WriteLine("Case 9");
                                LASact = dataSub;
                                tbLasAct.Text = LASact;
                                break;
                            case 10:
                                Console.WriteLine("Case 10");
                                DACrun = dataSub;
                                if (DACrun == "off") { cbLaserSetPwrRamp.Checked = false; }
                                else { cbLaserSetPwrRamp.Checked = true; }
                                break;
                            case 11:
                                Console.WriteLine("Case 11");
                                DACstp = dataSub;
                                DACstpInt = Int32.Parse(DACstp);
                                numRampStep.Value = PWMdtyInt;
                                break;
                            default:
                                break;
                        }
                        this.dataGridView2.Rows[i].Cells[4].Value = dataSub;
                        i++;
                    }
                    if (firstRun_Lame) cbActivePollLame.Checked = true;
                    firstRun_Lame = false;

                    if (cbDatalogLame.Checked == true)
                    {
                        FileIOLame dataLog = new FileIOLame();
                        if (firstRun_LameLog) { FileIOLame.fullFileName = "LAME datalog"; dataLog.fileCreate(); firstRun_LameLog = false; }
                        FileIOLame.Ilokot = Ilokot;
                        FileIOLame.AlrmIn = AlrmIn;
                        FileIOLame.Imode = Imode;
                        FileIOLame.Pmode = Pmode;
                        FileIOLame.LASenb = LASenb;
                        FileIOLame.PWMdty = PWMdty;
                        FileIOLame.PWMfrq = PWMfrq;
                        FileIOLame.LASset = LASset;
                        FileIOLame.TEMPmn = TEMPmn;
                        FileIOLame.LASact = LASact;
                        dataLog.fileWrite();
                    }
                }
                    }
            catch (Exception ex)
            {
                //MessageBox.Show("Error opening/writing to serial port :: " + ex.Message, "Error!");
            }
        }

        private void d_allThread()
        {
            while (D_all == true && formOpen)
            {
                serialWrite150WIN("Ddspgo(S,1)");
                Thread.Sleep((int)d_allTime.Value*1000);
            }
        }


        private void CbActivePoll_CheckedChanged(object sender, EventArgs e)
        {
            if (cbActivePoll.Checked == true)
            {
                D_all = true;
                Thread thr = new Thread(d_allThread);
                thr.Start();
                numDisTime.Enabled = false;
                numDispenseVol.Enabled = false;
                numLD3Offset.Enabled = false;
            }
            else
            {
                D_all = false;
                numDisTime.Enabled = true;
                numDispenseVol.Enabled = true;
                numLD3Offset.Enabled = true;
            }
        }

        private void d_allThreadLame()
        {
            while (D_all_Lame == true && formOpen)
            {
                serialWriteLaserMeister("Zdrpgo(S,1)");
                Thread.Sleep((int)d_allTimeLame.Value * 1000);
            }
        }

        private void CbActivePollLame_CheckedChanged(object sender, EventArgs e)
        {
            if (cbActivePollLame.Checked == true)
            {
                D_all_Lame = true;
                Thread thr = new Thread(d_allThreadLame);
                thr.Start();
            }
            else
            {
                D_all_Lame = false;
            }
        }

        private void BtnPortState150Win_Click(object sender, EventArgs e)
        {
            try
            {
                if (btnPortState150Win.Text == "Connect")
                {
                    string port = cboPorts150WIN.Text;
                    port = port.Substring(0, 5);
                    _serialPort = new SerialPort(port, 115200, Parity.None, 8, StopBits.One);
                    _serialPort.Handshake = Handshake.None;
                    _serialPort.DataReceived += new SerialDataReceivedEventHandler(sp_DataReceived);
                    _serialPort.ReadTimeout = 2000;
                    _serialPort.WriteTimeout = 2000;

                    _serialPort.Open();
                    numDisTime.Enabled = true;
                    numDispenseVol.Enabled = true;
                    numLD3Offset.Enabled = true;
                    d_allTime.Enabled = true;
                    cbActivePoll.Enabled = true;
                    cbDatalog150.Enabled = true;
                    cbChillerPwr.Enabled = true;
                    numChillSet.Enabled = true;
                    btnChillErrorClear.Enabled = true;
                    D_all = true;
                    firstRun = true;
                            btnPortState150Win.Text = "Connected";
                    btnPortState150Win.BackColor = Color.LightGreen;
                    serialWrite150WIN("xxxxxx(A,0)");
                }
                else if (btnPortState150Win.Text == "Connected")
                {
                    Thread.Sleep(100);
                    _serialPort.Close();

                    numDisTime.Enabled = false;
                    numDispenseVol.Enabled = false;
                    numLD3Offset.Enabled = false;
                    cbActivePoll.Enabled = false;
                    d_allTime.Enabled = false;
                    cbDatalog150.Enabled = false;
                    cbChillerPwr.Enabled = false;
                    numChillSet.Enabled = false;
                    btnChillErrorClear.Enabled = false;
                    btnPortState150Win.Text = "Connect";
                    btnPortState150Win.BackColor = Color.Coral;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening D Board serial port!");
            }
        }

        private void BtnPortState1_Click(object sender, EventArgs e)
        {
            try
            {
                if (btnPortState1.Text == "Connect")
                {
                    string port = cboPorts1.Text;
                    port = port.Substring(0, 5);
                    _serialPortCtrl = new SerialPort(port, 115200, Parity.None, 8, StopBits.One);
                    _serialPortCtrl.Handshake = Handshake.None;
                    _serialPortCtrl.DataReceived += new SerialDataReceivedEventHandler(sp_DataReceivedCtrl);
                    _serialPortCtrl.ReadTimeout = 2000;
                    _serialPortCtrl.WriteTimeout = 2000;

                    _serialPortCtrl.Open();
                    zDrpDist.Enabled = true;
                    zDrpDnSpd.Enabled = true;
                    zDrpUpSpeed.Enabled = true;
                    numRampStep.Enabled = true;
                    numRampInterval.Enabled = true;

                    cbInterlock.Enabled = true;
                    cbIOmode.Enabled = true;
                    cbPWMmode.Enabled = true;
                    btnLaserEn.Enabled = true;
                    cbDatalogLame.Enabled = true;
                    cbLaserSetPwrRamp.Enabled = true;
                    Thread.Sleep(100);
                    //cbActivePollLame.Enabled = true;
                    d_allTimeLame.Enabled = true;
                    btnPortState1.Text = "Connected";
                    btnPortState1.BackColor = Color.LightGreen;
                    D_all_Lame = true;
                    firstRun_Lame = true;
                }
                else if (btnPortState1.Text == "Connected")
                {
                    _serialPortCtrl.Close();

                    zDrpDist.Enabled = false;
                    zDrpDnSpd.Enabled = false;
                    zDrpUpSpeed.Enabled = false;
                    numRampStep.Enabled = false;
                    numRampInterval.Enabled = false;
                    cbInterlock.Enabled = false;
                    cbIOmode.Enabled = false;
                    cbPWMmode.Enabled = false;
                    btnLaserEn.Enabled = false;
                    cbLaserSetPwrRamp.Enabled = false;
                    cbActivePollLame.Checked = false;
                    cbActivePollLame.Enabled = false;
                    cbDatalogLame.Enabled = false;
                    d_allTimeLame.Enabled = false;
                    btnPortState1.Text = "Connect";
                    btnPortState1.BackColor = Color.Coral;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening Laser Meister Emulator Board serial port!");
            }
        }


        private void Timer1_Tick(object sender, EventArgs e)
        {
            //statusBar1.Panels[0].AutoSize = StatusBarPanelAutoSize.Spring;
            statusBar1.Panels[0].ToolTipText = "Windows Time";
            statusBar1.Panels[0].Width = 350;
            statusBar1.Panels[0].Text = DateTime.Now.ToString("dddd, dd, MMMM, yyyy hh:mm:ss tt", CultureInfo.InvariantCulture);
        }

        private void serialWrite150WIN(string data)
        {
            // Makes sure serial port is open before trying to write
            try
            {
                if (!_serialPort.IsOpen)
                    _serialPort.Open();
                string serialWriteData = data;
                _serialPort.Write(serialWriteData + "\r\n");
                tbComm.AppendText("Sent: " + data + "\r\n");
                tbComm.AppendText("\r\n");
                tb150WinSent.Text = serialWriteData;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error opening/writing to serial port :: " + ex.Message, "Error!");
            }
        }

        private void serialWriteLaserMeister(string data)
        {
            // Makes sure serial port is open before trying to write
            try
            {
                if (!_serialPortCtrl.IsOpen) _serialPortCtrl.Open();
                string serialWriteData = data;
                _serialPortCtrl.Write(serialWriteData + "\r\n");
                tbCommCtrl.AppendText("Sent: " + data + "\r\n");
                tbCommCtrl.AppendText("\r\n");
                tbLaserMeisterSent.Text = serialWriteData;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error opening/writing to serial port :: " + ex.Message, "Error!");
            }
        }

        private void zDrpDist_ValueChanged(object sender, EventArgs e)
        {
            int val = (int)(zDrpDist.Value * 10);
            serialWriteLaserMeister("S:LASset;" + val.ToString());
            Thread.Sleep(200);
        }

        private void zDrpDnSpd_ValueChanged(object sender, EventArgs e)
        {
            serialWriteLaserMeister("S:PWMdty;" + zDrpDnSpd.Value.ToString());
            Thread.Sleep(200);
        }

        private void zDrpUpSpeed_ValueChanged(object sender, EventArgs e)
        {
            serialWriteLaserMeister("S:PWMfrq;" + (zDrpUpSpeed.Value).ToString());
            Thread.Sleep(200);
        }

        private void NumRampStep_ValueChanged(object sender, EventArgs e)
        {
            serialWriteLaserMeister("S:DACstp;" + numRampStep.Value.ToString());
            Thread.Sleep(200);
        }

        private void NumRampInterval_ValueChanged(object sender, EventArgs e)
        {
            serialWriteLaserMeister("S:DACtim;" + numRampInterval.Value.ToString());
            Thread.Sleep(200);
        }

        private void NumLD1Offset_ValueChanged(object sender, EventArgs e)
        {
            if (!D_all && !firstRun)
                serialWrite150WIN("S:L1oset;" + numDisTime.Value.ToString());
        }

        private void numDispenseVol_ValueChanged(object sender, EventArgs e)
        {
            if (!D_all && !firstRun)
                serialWrite150WIN("S:L2oset;" + numDispenseVol.Value.ToString());
        }

        private void NumLD3Offset_ValueChanged(object sender, EventArgs e)
        {
            if (!D_all && !firstRun)
                serialWrite150WIN("S:L3oset;" + numLD3Offset.Value.ToString());
        }
        private void CbInterlock_CheckedChanged(object sender, EventArgs e)
        {
            if (cbInterlock.Checked == true)
            {
                serialWriteLaserMeister("S:Ilokot;1");
            }
            else
            {
                serialWriteLaserMeister("S:Ilokot;0");
            }
            Thread.Sleep(200);
        }

        private void CbIOmode_CheckedChanged(object sender, EventArgs e)
        {
            if (cbIOmode.Checked == true)
            {
                serialWriteLaserMeister("S:Imode;1");
            }
            else
            {
                serialWriteLaserMeister("S:Imode;0");
            }
            Thread.Sleep(200);
        }

        private void CbPWMmode_CheckedChanged(object sender, EventArgs e)
        {
            if (cbPWMmode.Checked == true)
            {
                serialWriteLaserMeister("S:Pmode;1");
            }
            else
            {
                serialWriteLaserMeister("S:Pmode;0");
            }
            Thread.Sleep(200);
        }

        private void BtnLaserEn_CheckedChanged(object sender, EventArgs e)
        {
            if (btnLaserEn.Checked == true)
            {
                serialWriteLaserMeister("S:LASenb;1");
            }
            else
            {
                serialWriteLaserMeister("S:LASenb;0");
            }
            Thread.Sleep(200);
        }

        private void CbLaserSetPwrRamp_CheckedChanged(object sender, EventArgs e)
        {
            
            if (cbLaserSetPwrRamp.Checked == true)
            {
                serialWriteLaserMeister("S:DACrun;1");
            }
            else
            {
                serialWriteLaserMeister("S:DACrun;0");
            }
            Thread.Sleep(200);

        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Close Program?", "150WIN", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK) {
                try
                {
                    //serialWriteLaserMeister("S:Ilokot;0");
                    serialWriteLaserMeister("S:LASset;0");
                    Thread.Sleep(500);
                    serialWriteLaserMeister("S:Pmode;0");
                    Thread.Sleep(500);
                    serialWriteLaserMeister("S:LASenb;0");
                    Thread.Sleep(500);
                    formOpen = false;
                    Thread.Sleep(500);
                    this.Close();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Error opening/writing to serial port :: " + ex.Message, "Error!");
                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //cbLaserSetPwrRamp.Checked = false;
            Thread.Sleep(200);
            cbPWMmode.Checked = false;
            Thread.Sleep(200);
            btnLaserEn.Checked = false;
            Thread.Sleep(200);
            cbIOmode.Checked = false;
            Thread.Sleep(200);
            cbInterlock.Checked = false;
            Thread.Sleep(200);
            zDrpDist.Value = 0;
            //numPWMDuty.Value = 0;
        }

        private void BtnClear150_Click(object sender, EventArgs e)
        {
            tbComm.Clear();
        }

        private void BtnClearLame_Click(object sender, EventArgs e)
        {
            tbCommCtrl.Clear();
        }

        private void btn10x_Click(object sender, EventArgs e)
        {
            pwrScale = !pwrScale;
            if (pwrScale == true) { zDrpDist.Increment = 1; btn10x.BackColor = Color.DarkGray; ; }
            else { double pwrScaleDb = 0.1; zDrpDist.Increment = (decimal)pwrScaleDb; btn10x.BackColor = Color.LightGray; }
        }

        private void cbChillerPwr_CheckedChanged(object sender, EventArgs e)
        {
            if (cbChillerPwr.Checked == true)
            {
                serialWrite150WIN("S:WCpowr;1");
            }
            else
            {
                serialWrite150WIN("S:WCpowr;0");
            }
        }
        private void numChillSet_ValueChanged(object sender, EventArgs e)
        {
            int val = (int)(numChillSet.Value * 10);
            serialWrite150WIN("S:WCsetp;" + val.ToString());
            Thread.Sleep(200);
        }

        private void btnChillErrorClear_Click(object sender, EventArgs e)
        {
            serialWrite150WIN("S:WCrset;2");
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            serialWrite150WIN("Ddspgo(S,1)");
        }

        private void zDropGo_Click(object sender, EventArgs e)
        {
            serialWriteLaserMeister("Zdrpgo(S,1)");
        }

        private void zDropHalt_Click(object sender, EventArgs e)
        {
            serialWriteLaserMeister("Zhalts(S,1)");
        }
    }
}
