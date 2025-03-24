using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;

namespace HLApi
{
	internal class ExcelTableWriter
	{
		ExcelWorkbook workbook;
		ExcelPackage package;
		string filepath;
		public void CreateAndUseBook(string bookName)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			package = new ExcelPackage("Journal");
			filepath = bookName;
			workbook = package.Workbook;
		}

		public void WriteWeek(List<Day> days, DateOnly dateOnly, int weekInd)
		{
			if (days.Count != 6)
				throw new Exception("Week has to include 7 days");
			string[] week =
			{
				"Понедельник",
				"Вторник",
				"Среда",
				"Четверг",
				"Пятница",
				"Суббота"
			};


			for (int i = 0; i < 6; ++i)
			{
				WriteDay(days[i].lessons, dateOnly, week[i], weekInd, i, 2);
				WriteWeekLeftHeader(days[i], weekInd);
				dateOnly = dateOnly.AddDays(1);
			}

		}
		private void WriteWeekLeftHeader(Day day, int sheetInd) { 
		var sheet = workbook.Worksheets[sheetInd];
			//1
			sheet.Cells[1, 1, (int)RowType.OmmisionsStart - 1, 1].Merge = true;
			int student = 0;
			for(var i = RowType.OmmisionsStart; i < RowType.OmmisionsEnd; ++i)
			{
				sheet.Cells[(int)i, 1].Value = i - RowType.OmmisionsStart + 1;
				if (day.lessons.Count != 0 && student < day.lessons[0].students.Count)
					sheet.Cells[(int)i, 2].Value = day.lessons[0].students[student++].fio;
			}
			sheet.Columns[1].Width = 2.7;
			//2
			sheet.Cells[1, 2].Value = "Дни недели";
			sheet.Cells[2, 2].Value = "Дата занятий";
			sheet.Cells[3, 2].Value = "Форма занятий";
			sheet.Cells[4, 2].Value = "Количество акад. часов";
			sheet.Cells[5, 2].Value = "Название дисциплины\n/\nФИО ";
			sheet.Cells[5, 2].Style.TextRotation = 0;
			sheet.Cells[5, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
			sheet.Cells[5, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			sheet.Columns[2].Width = 35;
			sheet.Columns[2].Style.Font.Name = "Times New Roman";
			sheet.Columns[1].Style.Font.Name = "Times New Roman";
			sheet.Columns[1].Style.Font.Bold = true;
			sheet.Columns[2].Style.Font.Bold = true;
		}
		private void WriteDay(List<Lesson> lessons, DateOnly dateOnly, string weekDayName, int weekInd, int dayOffset, int tableOffset = 0)
		{
			while (workbook.Worksheets.Count < weekInd + 1)
				workbook.Worksheets.Add(dateOnly.ToString());
			var sheet = workbook.Worksheets[weekInd];
			int startCol = dayOffset * 5 + 1 + tableOffset;
			int currentColumn;
			int lessonsLength = lessons.Count < 5 ? lessons.Count : 5;
			//Day name
			currentColumn = startCol + 4;
			sheet.Cells[(int)RowType.Date, startCol, (int)RowType.Date, currentColumn].Merge = true;
			sheet.Cells[(int)RowType.Date, startCol, (int)RowType.Date, currentColumn].Value = weekDayName;
			//Date
			sheet.Cells[(int)RowType.DayName, startCol, (int)RowType.DayName, currentColumn].Merge = true;
			sheet.Cells[(int)RowType.DayName, startCol, (int)RowType.DayName, currentColumn].Value = dateOnly.ToString();
			//Lessons 
			for (int i = 0; i < lessonsLength; ++i)
			{
				currentColumn = startCol + i;
				//Type
				sheet.Cells[(int)RowType.LessonsTypes, currentColumn].Value = lessons[i].lessonTypeAbbrev;
				//Time Length
				sheet.Cells[(int)RowType.LessonsLength, currentColumn].Value = lessons[i].lessonPeriod.lessonPeriodHours;
				//Name
				sheet.Cells[(int)RowType.LessonsNames, currentColumn].Value = lessons[i].nameAbbrev;
				for (int o = 0; o < (RowType.OmmisionsEnd - RowType.OmmisionsStart) && o < lessons[i].students.Count; ++o)
				{
					if (lessons[i].students[o].omission == null)
						sheet.Cells[o + (int)RowType.OmmisionsStart, currentColumn].Value = "";
					else
					{
						sheet.Cells[o + (int)RowType.OmmisionsStart, currentColumn].Value = "2";
						if (lessons[i].students[o].omission.respectfulOmission == true)
							sheet.Cells[o + (int)RowType.OmmisionsStart, currentColumn].Style.Font.Color.SetColor(System.Drawing.Color.Green);
						else
							sheet.Cells[o + (int)RowType.OmmisionsStart, currentColumn].Style.Font.Color.SetColor(System.Drawing.Color.Red);
					}
				}
			}


			//Style
			//Collumns
			for (int i = 0; i < 5; ++i)
			{
				var collumn = sheet.Columns[startCol + i];
				collumn.Style.Font.Bold = true;
				collumn.Style.Font.Name = "Times New Roman";
				collumn.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
				collumn.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
				collumn.Width = 3.35;
				collumn.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
			}
			//Rows
			var namesRow = sheet.Rows[(int)RowType.LessonsNames];
			namesRow.Style.TextRotation = 90;
			namesRow.Height = 36;
		}
		public void SaveBook()
		{
			File.WriteAllBytes(filepath, package.GetAsByteArray());
		}

		private enum RowType : int
		{
			DayName = 1,
			Date = 2,
			LessonsTypes,
			LessonsLength,
			LessonsNames,
			OmmisionsStart,
			OmmisionsEnd = 37
		}
	}


}
