using System.Data.SqlClient;
using FluentModbus;
using System.IO.Ports;
using System.Timers;
using System.Diagnostics;
namespace DotNetADOProject
{
	class Program
	{
		private static System.Timers.Timer aTimer;
		public static ModbusRtuClient client = new ModbusRtuClient
				{
					BaudRate = 9600,
					Parity = Parity.None,
					StopBits = StopBits.One,
					ReadTimeout = 1000
				};
		public static short[] reg1Datas = new short[27];
		public static short[] reg2Datas = new short[2];
		public static string port;
		public static int id;
		public static float u12;
		public static float u23;
		public static float u31;
		public static float i1;
		public static float i2;
		public static float i3;

		public static float p1;
		public static float p2;
		public static float p3;

		public static float freq;
		public static float Avolt;

		public static float Acurrent;

		public static float Consumption;

		public static bool firstTimeRunning = true;
		public static bool stopExecuteSQL = false;
		public static void Main(string[] args)
		{


			Console.WriteLine("Terminating the application...");
			

			Console.WriteLine("=================================================================================");
			Console.WriteLine("========================= ME96SS SERIAL COMMUNICATION ===========================");
			Console.WriteLine("=================================================================================");

			Console.WriteLine("*********************** Initialize Modbus RTU Communication *********************");
			
			tryConnect();

			
			SetTimer();

			while (true)
			{
				Thread t1 = new Thread(ThreadReadModbus);
				Thread t2 = new Thread(ThreadPrintData);
				t1.Start();
				t2.Start();
				t1.Join();
				t2.Join();
				firstTimeRunning = false;
			}
		}
		public static void tryConnect()
		{

			int tryTimes = 5;
			while (true)
			{
				Console.Write(">Enter Port: ");
				port = Console.ReadLine();
				Console.Write(">Enter Id: ");
				string? idString = Console.ReadLine();
				id = int.Parse(idString);

				try
				{
					client.Connect($"COM{port}", ModbusEndianness.BigEndian);
					Console.WriteLine("> Connected");
					Console.WriteLine("****************************** Executing program ********************************");

					break;
				}
				catch (Exception ex)
				{
					Console.WriteLine($"------------------------Here's some errors, try again[{tryTimes--}]-------------------------");
					Console.WriteLine("" +
						"Make sure you enter correct [PORT] number and device [ID]. " +
						"\nIf not works, please check your connection, following:" +
						"\n--> 1. Windows driver for Serial RS485 USB Converter" +
						$"\n--> 2. Whether another program using this [PORT] COM{port}?" +
						"\n--> 3. Check physical connection between computer and device." +
						"\nIf done, press ENTER to continue!");
					if (tryTimes == 0)
					{
						Kill();
					}
					Console.ReadLine();
				}
			}			
		}

		public static void Kill()
		{
			Console.WriteLine("The program is closed due to some process could not be terminated!");
			Process.GetCurrentProcess().Kill();
		}

		public static void ExecuteSQl(Object source, ElapsedEventArgs e)
		{
			if (!firstTimeRunning && !stopExecuteSQL)
			{
				string sqlCommand = "INSERT INTO METTER_DATA VALUES (";

				float[] me96ssData = [u12, u23, u31, i1, i2, i3, p1, p2, p3, freq, Avolt, Acurrent, Consumption];

				string timeLogString = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

				for (int i = 0; i < me96ssData.Length; i++)
				{
					sqlCommand += me96ssData[i];
					sqlCommand += ',';
				}

				sqlCommand += "(CONVERT(datetime, '";
				sqlCommand += timeLogString;
				sqlCommand += "', 103)));";

				SqlConnection? connection = null;
				try
				{
					connection = new SqlConnection("data source=TRATO; database=aue-me96ss; integrated security=SSPI");
					SqlCommand command = new SqlCommand(sqlCommand, connection);
					connection.Open();
					command.ExecuteNonQuery();
					Console.WriteLine("SQL query executed!");
				}
				catch (Exception)
				{
					Console.WriteLine("Failed to execute SQL! Please check connection");
				}
				finally
				{
#pragma warning disable CS8602
					connection.Close();
#pragma warning restore CS8602
				}
			}
		}

		private static void ThreadReadModbus()
		{
			int errorCount = 0;
			while (true)
			{
				try
				{
					var Reg1 = client.ReadHoldingRegisters<short>(id, 768, 26);

					try
					{
						client.Connect($"COM{port}", ModbusEndianness.BigEndian);
					}
					catch
					{
						Console.WriteLine("Error at connection!");
					}

					var Reg2 = client.ReadHoldingRegisters<short>(id, 1280, 2);

					try
					{
						client.Connect($"COM{port}", ModbusEndianness.BigEndian);
					}
					catch (Exception e)
					{
						Console.WriteLine(".");
					}
					reg1Datas = Reg1.ToArray();
					reg2Datas = Reg2.ToArray();
					errorCount = 0;
					Console.WriteLine("OK!");
					break;
				}
				catch (Exception e) 
				{
					Console.Write(".");
					if (errorCount++ > 9)
					{
						Console.WriteLine("There's some error cause program to stopped");
						Console.WriteLine("Press ENTER to fix this problem!");
						Console.ReadLine();
						stopExecuteSQL = true;
						tryConnect();
					}
					Thread.Sleep(1000);
				}
			}
			stopExecuteSQL = false;
			Thread.Sleep(2000);
		}
		private static void ThreadPrintData() {
			if (!firstTimeRunning)
			{
				u12 = ToFloat(reg1Datas[14]);
				u23 = ToFloat(reg1Datas[15]);
				u31 = ToFloat(reg1Datas[16]);

				Avolt = ToFloat(reg1Datas[17]);

				i1 = ToFloat(reg1Datas[0]);
				i2 = ToFloat(reg1Datas[1]);
				i3 = ToFloat(reg1Datas[2]);

				Acurrent = ToFloat(reg1Datas[4]);

				p1 = ToFloat(reg1Datas[23]);
				p2 = ToFloat(reg1Datas[24]);
				p3 = ToFloat(reg1Datas[25]);

				freq = ToFloat(reg1Datas[22]);
				Consumption = ToFloat(reg2Datas[0] + reg2Datas[1]);

				Console.WriteLine("" +
				  "1. Voltage:	 {0, 5} V  {1, 5} V  {2, 5} V" +
				"\n2. Current:	 {3, 5} A  {4, 5} A  {5, 5} A" +
				"\n3. Power:	 {6, 5} kW {7, 5} kW {8, 5} kW" +
				"\n4. Average Voltage: {9, 5} V" +
				"\n5. Average Current: {10, 5} A" +
				"\n6. Frequency:  {11, 10} Hz" +
				"\n7. Consumption:{12, 10} kWh",
				u12, u23, u31, i1, i2, i3, p1, p2, p3, Avolt, Acurrent, freq, Consumption
				);
				Thread.Sleep(5000);
			}
		}

		private static float ToFloat (int wholeNumber)
		{
			int intNumber = wholeNumber / 10;
			int floatNumber = wholeNumber % 10;
			return (float)intNumber + (float)(floatNumber / 10f);
		}
		private static void SetTimer()
		{
			aTimer = new System.Timers.Timer(30000);
			aTimer.Elapsed += ExecuteSQl;
			aTimer.AutoReset = true;
			aTimer.Enabled = true;
		}
	}
}