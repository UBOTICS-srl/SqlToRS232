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

        private DateTime startTime;

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
                        Console.Clear();
                        Console.WriteLine("\nPremere INVIO per uscire dall'applicazione...\n");
                        Console.WriteLine($"Applicazione partita alle ore {startTime:HH:mm:ss.fff}");

                        try
                        {
                            ComPort.Open();
                        }
                        catch
                        {
                            Console.WriteLine($"Impossibile aprire la porta {ComPort.PortName}");
                            Console.WriteLine($"Ritento in 5 secondi...");
                            Thread.Sleep(4000);
                            ComPort.Dispose();
                            ComPort = GetSerialSettings();
                            Thread.Sleep(1000);
                            throw;
                        }

                        try
                        {
                            DbConnection.Open();
                        }
                        catch
                        {
                            Console.WriteLine($"Impossibile connettersi al DataBase");
                            Console.WriteLine($"Ritento in 5 secondi...");
                            Thread.Sleep(5000);
                            throw;
                        }

                        flag = false;
                        SetTimer();
                    }
                    catch
                    {
                        flag = true;
                    }
                }

                startTime = DateTime.Now;
                Console.WriteLine("\nPremere INVIO per uscire dall'applicazione...\n");
                Console.WriteLine($"Applicazione partita alle ore {startTime:HH:mm:ss.fff}");
                Console.ReadLine();

                aTimer.Stop();
                aTimer.Dispose();

                Console.WriteLine("Chiusura dell'applicazione in corso...");
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

            Console.Clear();

            Console.WriteLine("\nPremere INVIO per uscire dall'applicazione...\n");
            Console.WriteLine($"Applicazione partita alle ore {startTime:HH:mm:ss.fff}");

            Console.WriteLine($"Lettura della tabella {TableName}");
            SqlDataReader sqlDataReader = new SqlCommand($"SELECT TOP 1 Peso1 FROM {TableName}", DbConnection).ExecuteReader();

            bool hasRows = sqlDataReader.Read();

            if (hasRows)
            {
                int dbPeso = Convert.ToInt32(sqlDataReader.GetDouble(0));
                Console.WriteLine($"Letto peso {dbPeso}");

                try
                {
                    Console.WriteLine($"Invio peso {dbPeso} alla porta {ComPort.PortName}");

                    string scaleCommand = $"P+{$"{dbPeso,6:D6}"}\r";
                    string printScaleCommand = $"P+{$"{dbPeso,6:D6}"}\\r";

                    Console.WriteLine($"Invio comando {printScaleCommand} alla porta {ComPort.PortName}");

                    if (Debugger.IsAttached == false)
                    {
                        ComPort.Write(scaleCommand);
                    }
                }
                catch
                {
                    Console.WriteLine($"Impossibile inviare il peso alla porta {ComPort.PortName}");
                }
            }
            else
            {
                Console.WriteLine($"Nessuna riga nella tabella {TableName}");
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
