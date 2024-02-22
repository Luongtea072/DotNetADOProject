using System.Data.SqlClient;
namespace DotNetADOProject
{
	class Program
	{
		static void Main(string[] args)
		{
	
			string sqlCommand = "INSERT INTO METTER_DATA VALUES (";

			double u12 = 383.81;
			double u23 = 384.33;
			double u31 = 383.64;

			double i1 = 558.9;
			double i2 = 384.9;
			double i3 = 461.54;

			double p1 = 117.39;
			double p2 = 83.82;
			double p3 = 97.33;

			double freq = 56;
			double Avolt = 383.81;
			double Acurrent = 466.23;
			double Consumption = 63.76;

			double[] me96ssData = [u12, u23, u31, i1, i2, i3, p1, p2, p3, freq, Avolt, Acurrent, Consumption];

			string timeLogString = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

			for (int i = 0; i < me96ssData.Length; i++)
			{
				sqlCommand += me96ssData[i];
				sqlCommand += ',';
			}

			sqlCommand += "(CONVERT(datetime, '";
			sqlCommand += timeLogString;
			sqlCommand += "', 103)));";

			Console.WriteLine(sqlCommand);

			new Program().ExecuteSQl(sqlCommand);
			
		}
		public void ExecuteSQl(string sqlCommand)
		{
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
}