using static System.Net.Mime.MediaTypeNames;
using System.Windows.Forms;

static class Program
{
	internal class UserDialog
	{
		OpenFileDialog ofd = new OpenFileDialog();

		public void UserHello()
		{
			if (true)
			{
				Console.ForegroundColor = ConsoleColor.Blue;
				Console.WriteLine("Парсер журнала БГУИР ИИС в помощь себе и другим старостам");
				Console.WriteLine("Если будешь юзать, то не грех и спасибо сказать))");

				Console.WriteLine("Кто в гитхаб звезду поставит - в рай без очереди, там же будет код проги)");
				Console.WriteLine("\nGitHub : https://github.com/NikitaBurinsky?tab=repositories");
				Console.WriteLine("Instagram (Предложения, правки, предьявы) : @format_of_file");
				Console.WriteLine("LinkedIn : https://www.linkedin.com/in/nikita-burinsky-1a1313261/");
				Console.ResetColor();
			}
		}

		public (string login, string password) GetUserLoginPassword()
		{
			Console.WriteLine("Логин (Студак старосты) :");
			string log = Console.ReadLine();
			while (string.IsNullOrEmpty(log))
				log = Console.ReadLine();
			Console.WriteLine("Пароль (от ИИСа старосты):");
			string pass = Console.ReadLine();
			while (string.IsNullOrEmpty(pass))
				pass = Console.ReadLine();
			return (log, pass);
		}
		[STAThread]
		public string GetUserFilepath()
		{
			ofd.Title = "Journal book filepath";
			ofd.Filter = "Excel Files|*.xlsx;*.xlsm";
			ofd.CheckFileExists = false;

			Console.WriteLine("Выбери папку для книги журнала");
			ofd.ShowDialog();
			Console.WriteLine("Выбранный путь: " + ofd.FileName);
			return ofd.FileName;
		}

		public void UserEndDownload()
		{
			Console.Clear();
			Console.ForegroundColor = ConsoleColor.Yellow;
			Console.WriteLine("\nФайл готов ;)\n");
			Console.ResetColor();
			UserHello();
		}
	}
		[STAThread]
		static public void Main()
		{
			UserDialog userDialog = new UserDialog();
			userDialog.UserHello();
			var login_password = userDialog.GetUserLoginPassword();
			string filename = userDialog.GetUserFilepath();

			HLApi.ExcelTableWriter ex = new HLApi.ExcelTableWriter();
			ex.CreateAndUseBook(filename);

			Console.Clear();
			var dateParser = new HLApi.DateParser(login_password.login, login_password.password);
			List<HLApi.Day> testWeek = new List<HLApi.Day>();
			DateOnly rd = new DateOnly(2024, 9, 2);
			DateOnly ToDate = new DateOnly(2025, 3, 24);
			for (int week = 0; rd < ToDate; ++week)
			{
				for (int i = 0; i < 6; ++i)
				{
					var taskDay = dateParser.GetOmmisionsByDateRespone(rd.AddDays(i));
					while (!taskDay.IsCompleted) { };
					testWeek.Add(taskDay.Result);
				}
				ex.WriteWeek(testWeek, rd, week);
				rd = rd.AddDays(7);
				testWeek.Clear();
			}
			ex.SaveBook();
			userDialog.UserEndDownload();
		Console.ReadLine();
		}
	}