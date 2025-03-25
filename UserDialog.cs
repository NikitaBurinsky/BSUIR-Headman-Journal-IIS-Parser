using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
		string pass = ReadPassword();
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

		Console.WriteLine("Выбери файл excel или создай новый файл для книги журнала");
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

	public (DateOnly from, DateOnly to) UserGetFromDate_ToDate()
	{
		Console.WriteLine("Введи загружаемый диапазон дат (\"Enter\" чтобы выбрать весь год) \nПример : 12.09.2024 - 12.02.2025");
		DateOnly from = new DateOnly(DateTime.Today.Year, 9, 4), to = DateOnly.FromDateTime(DateTime.Today);
		bool scanSuccess = true;
		while (true)
		{
			scanSuccess = true;
			string ent = Console.ReadLine().Replace(" ", string.Empty);
			if (string.IsNullOrEmpty(ent))
			{
				if (from > to)
					from = from.AddYears(-1);
				return (from, to);
			}
			scanSuccess = scanSuccess && DateOnly.TryParse(ent.Substring(0, 10), out from);
			scanSuccess = scanSuccess && DateOnly.TryParse(ent.Substring(11, 10), out to);
			scanSuccess = scanSuccess && from <= to;
			scanSuccess = scanSuccess && to <= DateOnly.FromDateTime(DateTime.Today);
			if (!scanSuccess)
			{
				Console.ForegroundColor = ConsoleColor.DarkRed;
				Console.WriteLine("Неверный ввод диапазона, попробуй еще раз:");
				Console.ResetColor();
			}
			else
				break;
		}
		//Console.WriteLine($"Введеный диапазон : {from.ToString()} - {to.ToString()}");
		return (from, to);
	}
	public static string ReadPassword()
	{
		string password = "";
		ConsoleKeyInfo info = Console.ReadKey(true);
		while (info.Key != ConsoleKey.Enter)
		{
			if (info.Key != ConsoleKey.Backspace)
			{
				Console.Write("*");
				password += info.KeyChar;
			}
			else if (info.Key == ConsoleKey.Backspace)
			{
				if (!string.IsNullOrEmpty(password))
				{
					password = password.Substring(0, password.Length - 1);
					int pos = Console.CursorLeft;
					Console.SetCursorPosition(pos - 1, Console.CursorTop);
					Console.Write(" ");
					Console.SetCursorPosition(pos - 1, Console.CursorTop);
				}
			}
			info = Console.ReadKey(true);
		}
		Console.WriteLine();
		return password;
	}

}
