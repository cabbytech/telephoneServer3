////zzz

using CabbyTechOffice;
using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Net;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows.Forms;
using Traysoft.AddTapi;
//
namespace TelephoneServer3
{
    //
    public partial class Main : Form
    {
        private bool InitDone = false;
        private int RINGCOUNTER1;
        private TapiLine objLine1;
        private TapiLine objLine2;
        private int RINGCOUNTER2;
        private string connStr;

        int udpBroadCastPort = 11002;

        public TelRecordClass line1Record { get; set; }
        public TelRecordClass line2Record { get; set; }

        public Main()
        {
            InitializeComponent();
            FormClosing += Main_FormClosing;
            line1Record = new TelRecordClass();
            line2Record = new TelRecordClass();
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                TapiApp.Shutdown();
            }
            catch (Exception)
            {


            }
        }

        public async Task Init()
        {

            FlashLabel();



            ConfigCabbyTechOffice fig = ConfigUI.GetCabbyTechConfig();



            System.Data.SqlClient.SqlConnectionStringBuilder connBuilder = new System.Data.SqlClient.SqlConnectionStringBuilder
            {
                InitialCatalog = "TAXIDB",
                UserID = "sa",
                Password = fig.DbPassword,
                DataSource = fig.DatabaseComputerName + "\\" + fig.DatabaseInstanceName,
                ConnectTimeout = 8
            };

            udpBroadCastPort = fig.UDPTelephonePort;
            connStr = connBuilder.ConnectionString;






            try
            {
                await Task.Run(() =>
                {
                    TapiApp.SerialNumber = "7N8E3MN-XNVCPC8-XAVADBE-44PEM";
                    TapiApp.Initialize("CallerID");
                });
                InitDone = true;

                TapiApp.TapiError += OnTapiError;
                TapiApp.LineAdded += OnLineAdded;
                TapiApp.LineClosed += OnLineClosed;
                TapiApp.LineRemoved += OnLineRemoved;
                TapiApp.IncomingCall += OnIncommingCall;
                TapiApp.OutgoingCall += OnOutgoingCall;
                TapiApp.CallConnected += OnCallConnected;

                TapiApp.CallDisconnected += OnCallDisconnected;

                UpdateLinesCombobox1();
                UpdateLinesCombobox2();

            }
            catch (TapiException ex)
            {
                MessageBox.Show(ex.Message, "TapiException!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }


        }

        private async Task FlashLabel()
        {
            label1.ForeColor = Color.Red;
            while (!InitDone)
            {
                label1.Text = "WAIT FOR INITIALISATION";
                label1.Refresh();
                await Task.Delay(500);

            }
            label1.Text = "INITIALISATION DONE";
            label1.Refresh();
            await Task.Delay(3000);
            label1.ForeColor = Color.Black;
            label1.Text = "....";
            label1.Refresh();

        }

        private void UpdateLinesCombobox1()
        {
            try
            {
                comboBox1.Items.Clear();
                foreach (TapiLine item in TapiApp.Lines)
                {
                    comboBox1.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                AddToLog("Combo 1 " + ex.Message);
            }
        }

        private void UpdateLinesCombobox2()
        {
            try
            {
                comboBox2.Items.Clear();
                foreach (TapiLine item in TapiApp.Lines)
                {
                    comboBox2.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                AddToLog("Combo 2 " + ex.Message);
            }
        }

        public void OnCallDisconnected(object sender, TapiEventArgs args)
        {
            //called on worker thread so marshall on to ui thread

            MethodInvoker m = new MethodInvoker(() =>
            {
                int lineNumber = 0;
                try
                {
                    if (objLine1 != null)
                    {
                        if (args.Line.Name == objLine1.Name)
                        {
                            lineNumber = 1;
                            timer1.Stop();

                            label1.BackColor = Color.White;
                            label1.Text = "IDLE";

                        }
                    }
                    if (objLine2 != null)
                    {
                        if (args.Line.Name == objLine2.Name)
                        {
                            lineNumber = 2;
                            timer2.Stop();
                            label2.BackColor = Color.White;
                            label2.Text = "IDLE";


                        }
                    }
                    //AddToLog($"{args.Line.Name} is Disconnected");
                    //AddToLog($"");

                    SqlConnection conn = new SqlConnection(connStr);
                    SqlCommand cmd = new SqlCommand
                    {
                        CommandType = System.Data.CommandType.Text,
                        CommandText = "INSERT INTO TELLOG (Timemark,rings,answered,callended,line,TelNumber) VALUES (@timemark,@rings,@answered,@callended,@line,@telnumber)"
                    };

                    lineNumber = 1;
                    if (lineNumber == 1)
                    {
                        cmd.Parameters.AddWithValue("@line", 1);
                        cmd.Parameters.AddWithValue("@rings", RINGCOUNTER1);
                        cmd.Parameters.AddWithValue("@timemark", DateTime.Now);
                        cmd.Parameters.AddWithValue("@callended", DateTime.Now);
                        cmd.Parameters.AddWithValue("@answered", line1Record.answered);
                        cmd.Parameters.AddWithValue("@telnumber", line1Record.telNumber);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@line", 2);
                        cmd.Parameters.AddWithValue("@rings", RINGCOUNTER2);
                        cmd.Parameters.AddWithValue("@timemark", DateTime.Now);
                        cmd.Parameters.AddWithValue("@callended", DateTime.Now);
                        cmd.Parameters.AddWithValue("@answered", line2Record.answered);
                        cmd.Parameters.AddWithValue("@telnumber", line2Record.telNumber);
                    }
                    conn.Open();
                    cmd.Connection = conn;
                    cmd.ExecuteNonQuery();
                    label4.Text = $"LINE {lineNumber}  SAVED";

                }
                catch (Exception ex)
                {
                    label4.Text = ex.Message;

                }



            });
            this.Invoke(m);
        }
        private void OnCallConnected(object sender, TapiEventArgs args)
        {
            MethodInvoker m = new MethodInvoker(() =>
            {
                try
                {
                    string str = $"{Environment.MachineName}%4%{args.Call.Address.Address}%{args.Call.CallerID}";
                    if (args.Line.Name == objLine1?.Name)
                    {
                        label1.BackColor = Color.Yellow;
                        label1.Text = $"ANSWERED BY EXT { args.Call.Address.Address} AFTER  {RINGCOUNTER1} RINGS";
                        timer1.Stop();
                        line1Record.answered = true;
                    }
                    if (args.Line.Name == objLine2?.Name)
                    {
                        label2.BackColor = Color.Yellow;
                        label2.Text = $"ANSWERED BY EXT { args.Call.Address.Address} AFTER  {RINGCOUNTER2} RINGS";
                        timer2.Stop();
                        line2Record.answered = true;
                    }
                    AddToLog($"{args.Line.Name} CONNECTED");
                    SyncBroadcastOut(str, udpBroadCastPort);
                }
                catch (Exception ex)
                {
                }
            });
            this.Invoke(m);
        }
        private void SyncBroadcastOut(string msg, int port)
        {
            try
            {
                AddToLog($"SENDIND UDP-----{msg} ON PORT {port}");
                UdpClient udp = new UdpClient();
                byte[] bytes;
                IPEndPoint ep = new IPEndPoint(IPAddress.Parse("255.255.255.255"), port);
                bytes = System.Text.ASCIIEncoding.ASCII.GetBytes(msg);
                udp.Send(bytes, bytes.Length, ep);
                udp.Close();
                udp = null;

            }
            catch (Exception ex)
            {
            }
        }
        private void OnIncommingCall(object sender, TapiEventArgs args)
        {
            MethodInvoker m = new MethodInvoker(() =>
            {
                if (objLine1 != null)
                {
                    if (args.Line.Name == objLine1.Name)
                    {
                        line1Record = new TelRecordClass();
                        RINGCOUNTER1 = 1;
                        label1.BackColor = Color.Red;
                        label1.Text = $"RINGING {RINGCOUNTER1}";
                        timer1.Start();
                        line1Record.answered = false;
                        line1Record.CallbeginTime = DateTime.Now;
                        line1Record.telNumber = args.Call.CallerID;

                        AddToLog($"{Environment.NewLine} INCOMMING CALL {objLine1} {args.Call.CallerID}");
                    }
                }
                if (objLine2 != null)
                {
                    if (args.Line.Name == objLine2.Name)
                    {
                        line2Record = new TelRecordClass();
                        RINGCOUNTER2 = 1;
                        label2.BackColor = Color.Red;
                        label2.Text = $"RINGING {RINGCOUNTER2}";
                        timer2.Start();
                        line2Record.answered = false;
                        line2Record.CallbeginTime = DateTime.Now;
                        line1Record.telNumber = args.Call.CallerID;

                        AddToLog($"{Environment.NewLine} INCOMMING CALL {objLine2} {args.Call.CallerID}");
                    }
                }







            });

            this.Invoke(m);

        }

        private void OnOutgoingCall(object sender, TapiEventArgs args)
        {
            try
            {

            }
            catch (Exception)
            {


            }
        }

        private void OnLineRemoved(object sender, TapiEventArgs args)
        {

        }

        private void OnLineClosed(object sender, TapiEventArgs args)
        {

        }

        private void OnLineAdded(object sender, TapiEventArgs args)
        {

        }

        private void OnTapiError(object sender, TapiErrorEventArgs args)
        {
            string msg;
            MethodInvoker m = new MethodInvoker(() =>
            {
                try
                {
                    if (args.Line == null)
                    {
                        msg = $"TAPI ERROR LINE IS NULL {args.Message}";
                        return;
                    }
                    msg = $"TAPI ERROR {args.Line.Name} WAS FORCIBLY CLOSED";
                    AddToLog(msg);
                }
                catch (Exception)
                {


                }

            });

            this.Invoke(m);

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        //line 1
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                objLine1 = (TapiLine)comboBox1.SelectedItem;
                if (objLine1 == null)
                {
                    throw new NullReferenceException("LINE IS NULL");
                }
                if (objLine1.IsOpen)
                {
                    throw new ArgumentOutOfRangeException("Already Monitoring Line");
                }
                objLine1.RingsToAnswer = 0;
                objLine1.Open(true, null);
                AddToLog("Started monitoring line " + objLine1.Name);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        //line 2
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                objLine2 = (TapiLine)comboBox2.SelectedItem;
                if (objLine2 == null)
                {
                    throw new NullReferenceException("LINE IS NULL");
                }
                if (objLine2.IsOpen)
                {
                    throw new ArgumentOutOfRangeException("Already Monitoring Line");
                }
                objLine2.RingsToAnswer = 0;
                objLine2.Open(true, null);
                AddToLog("Started monitoring line " + objLine2.Name);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void AddToLog(string v)
        {
            log.Items.Add(v);
            if (log.Items.Count > 500)
            {
                for (int i = 0; i < 450; i++)
                {
                    log.Items.RemoveAt(0);
                }
            }
            log.SelectedIndex = log.Items.Count - 1;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            RINGCOUNTER1 += 1;
            label1.Text = $"RINGING {RINGCOUNTER1}";
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            RINGCOUNTER2 += 1;
            label2.Text = $"RINGING {RINGCOUNTER2}";
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }


        private void label3_Click(object sender, EventArgs e)
        {

        }

        private async void Main_Shown(object sender, EventArgs e)
        {
            await Init();
        }
    }
}
