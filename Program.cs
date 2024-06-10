// Decompiled with JetBrains decompiler
// Type: SqlToRs232.Program
// Assembly: SqlToRs232, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 492879F7-A08B-40B0-BCAB-9767FC0274F6
// Assembly location: C:\GitHub\RS232\SqlToRs232.exe

using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;

using Timer = System.Timers.Timer;

namespace SqlToRs232
{
    class DbToSerial
    {
        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        private Timer aTimer;

        private string DbConnectionString;

        private string TableName;

        private SerialPort ComPort;
        private SqlConnection DbConnection;
        public int contatore = 0;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private SerialPort GetSerialSettings()
        {
            string portName = ConfigurationManager.AppSettings["PortName"];

            string speed = ConfigurationManager.AppSettings["Speed"];
            int speedVal = Convert.ToInt32(speed);

            string parity = ConfigurationManager.AppSettings["Parity"];
            Parity parityVal;

            switch (parity.ToUpper())
            {
                case "NONE":
                    parityVal = Parity.None;
                    break;

                case "EVEN":
                    parityVal = Parity.Even;
                    break;

                case "MARK":
                    parityVal = Parity.Mark;
                    break;

                case "ODD":
                    parityVal = Parity.Odd;
                    break;

                case "SPACE":
                    parityVal = Parity.Space;
                    break;

                default:
                    parityVal = Parity.None;
                    break;
            }

            string dataBits = ConfigurationManager.AppSettings["DataBits"];
            int dataBitsVal = Convert.ToInt32(dataBits);

            string stopBits = ConfigurationManager.AppSettings["StopBit"];
            StopBits stopBitsVal;

            switch (stopBits.ToUpper())
            {
                case "NONE":
                    stopBitsVal = StopBits.None;
                    break;

                case "ONE":
                    stopBitsVal = StopBits.One;
                    break;

                case "ONEPOINTFIVE":
                    stopBitsVal = StopBits.OnePointFive;
                    break;

                case "TWO":
                    stopBitsVal = StopBits.Two;
                    break;

                default:
                    stopBitsVal = StopBits.None;
                    break;
            }

            SerialPort serialPort = new SerialPort(portName, speedVal, parityVal, dataBitsVal, stopBitsVal);

            return serialPort;

        }

        public void RunProgram()
        {
            try
            {
                if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
                {
                    Environment.Exit(1);
                }

                ShowWindow(GetConsoleWindow(), 0);
                bool flag = true;
                while (flag)
                {
                    try
                    {
                        ComPort.Open();
                        DbConnection.Open();
                        flag = false;
                        SetTimer();
                    }
                    catch
                    {
                        flag = true;
                        Thread.Sleep(5000);
                        ComPort.Dispose();
                        //comPort = new SerialPort("COM3", 9600, Parity.None, 8, StopBits.One);
                        ComPort = GetSerialSettings();
                        Thread.Sleep(1000);
                    }
                }

                Console.WriteLine("\nPress the Enter key to exit the application...\n");
                Console.WriteLine($"The application started at {(object)DateTime.Now:HH:mm:ss.fff}");
                Console.ReadLine();

                aTimer.Stop();
                aTimer.Dispose();

                Console.WriteLine("Terminating the application...");
                ComPort.Close();
                DbConnection.Close();
            }
            catch
            {
            }
        }


        public DbToSerial()
        {
            DbConnectionString = ConfigurationManager.ConnectionStrings["Connessione"].ConnectionString;

            DbConnection = new SqlConnection(DbConnectionString);

            TableName = ConfigurationManager.AppSettings["TableName"];

            //comPort = new SerialPort("COM3", 9600, Parity.None, 8, StopBits.One);
            ComPort = GetSerialSettings();
        }

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            aTimer.Enabled = false;

            SqlDataReader sqlDataReader = new SqlCommand($"SELECT TOP 1 Peso1 FROM {TableName}", DbConnection).ExecuteReader();

            bool hasRows = sqlDataReader.Read();

            if (hasRows)
            {
                int int32 = Convert.ToInt32(sqlDataReader.GetDouble(0));
                ComPort.Write($"P+{$"{(object)int32,6:D6}"}\r");
            }

            sqlDataReader.Close();

            aTimer.Enabled = true;
        }

        private void SetTimer()
        {
            aTimer = new Timer(500.0);
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
        }
    }

    class Program
    {

        public static void Main(string[] args)
        {
            DbToSerial instance = new DbToSerial();
            instance.RunProgram();
        }
    }
}
