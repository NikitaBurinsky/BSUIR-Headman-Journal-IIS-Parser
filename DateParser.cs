using System.Text;
using System.Net.Http;
using System.Text.Json;


namespace HLApi
{
	internal class DateParser
	{
		HttpClient httpClient = new HttpClient();
		public DateParser(string login, string password)
		{
			Loggining(login, password);
		}
		private bool Loggining(string login, string password)
		{

			HttpRequestMessage logging = new HttpRequestMessage(HttpMethod.Post, "https://iis.bsuir.by/api/v1/auth/login");
			logging.Content = new StringContent(
				$"{{\"username\": \"{login}\", \"password\": \"{password}\"}}",
				Encoding.UTF8,
				"application/json");
			HttpResponseMessage response = httpClient.Send(logging);
			if (!response.IsSuccessStatusCode)
			{
				throw new Exception($"Logging failed : {response.StatusCode}");
			}
			return true;
		}

		public async Task<Day>? GetOmmisionsByDateRespone(DateOnly date)
		{
			HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Get,
				@$"https://iis.bsuir.by/api/v1/grade-book/by-date?date={date.Year}-{date.Month.ToString("D2")}-{date.Day.ToString("D2")}T06:42:06.64Z");
			Console.WriteLine($"Request to GET : {message.RequestUri}");

			var response = httpClient.Send(message);
			if (!response.IsSuccessStatusCode)
			{
				Console.WriteLine($"Request failed : Status  : {response.StatusCode}");
				throw new Exception($"Failed GET grade-book by date : {response.StatusCode}");
			}
			string jsonString = await response.Content.ReadAsStringAsync();

			return await DeParseResponse(jsonString); ;
		}

		private async Task<Day> DeParseResponse(string njsonString)
		{
			// Десериализация JSON в объект
			Day Day = JsonSerializer.Deserialize<Day>(njsonString);

			// Вывод данных для проверки
			/*
			foreach (var lesson in Day.lessons)
			{
				Console.WriteLine($"Lesson ID: {lesson.id}");
				Console.WriteLine($"Date: {lesson.dateString}");
				Console.WriteLine($"Subject: {lesson.nameAbbrev}");
				Console.WriteLine($"Type: {lesson.lessonTypeAbbrev}");
				Console.WriteLine($"Period: {lesson.lessonPeriod.startTime} - {lesson.lessonPeriod.endTime}");
				Console.WriteLine($"SubGroup: {lesson.subGroup}");
				Console.WriteLine("Students:");

				foreach (var student in lesson.students)
				{
					Console.WriteLine($"  Student ID: {student.id}");
					Console.WriteLine($"  Name: {student.fio}");
					if (student.omission != null)
					{
						Console.WriteLine($"  Omission: {student.omission.missedHours} hours, Respectful: {student.omission.respectfulOmission}");
					}
					else
					{
						Console.WriteLine("  Omission: None");
					}
				}
				Console.WriteLine(new string('-', 40));
			}*/
			return Day;
		}
	}
}
