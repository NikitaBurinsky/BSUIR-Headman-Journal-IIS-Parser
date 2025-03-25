using static System.Net.Mime.MediaTypeNames;
using System.Windows.Forms;

static class Program
{

	[STAThread]
	static public void Main()
	{
		try
		{
			UserDialog userDialog = new UserDialog();
			userDialog.UserHello();
			var login_password = userDialog.GetUserLoginPassword();
			var Dates = userDialog.UserGetFromDate_ToDate();
			DateOnly ToDate = Dates.to;
			DateOnly curDownloadMondayDate = Dates.from;
			curDownloadMondayDate = curDownloadMondayDate.AddDays(-1 * (curDownloadMondayDate.DayOfWeek - DayOfWeek.Monday));
			string filename = userDialog.GetUserFilepath();
			HLApi.ExcelTableWriter ex = new HLApi.ExcelTableWriter();
			ex.CreateAndUseBook(filename);

			Console.Clear();
			var dateParser = new HLApi.DateParser(login_password.login, login_password.password);
			List<HLApi.Day> testWeek = new List<HLApi.Day>();


			for (int week = 0; curDownloadMondayDate < ToDate; ++week)
			{
				for (int i = 0; i < 6; ++i)
				{
					var taskDay = dateParser.GetOmmisionsByDateRespone(curDownloadMondayDate.AddDays(i));
					while (!taskDay.IsCompleted) { };
					testWeek.Add(taskDay.Result);
				}
				ex.WriteWeek(testWeek, curDownloadMondayDate, week);
				curDownloadMondayDate = curDownloadMondayDate.AddDays(7);
				testWeek.Clear();
			}
			ex.SaveBook();
			userDialog.UserEndDownload();
		}
		catch (Exception ex)
		{
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine("Ошибка: " + ex.Message + '\n');
			Console.ResetColor();
			Console.ReadLine();
			Console.Clear();
			System.Diagnostics.Process.Start(System.Environment.ProcessPath);
			Environment.Exit(0);
		}
		Console.ReadLine();
	}
}