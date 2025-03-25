using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;
using System.Windows.Forms;

namespace HLApi
{
	internal class ExcelTableWriter
	{
		ExcelWorkbook workbook;
		ExcelPackage package;
		string filepath;
		int studentsCount;
		Dictionary<string, int> studentsIndByName = new Dictionary<string, int>();//!!Подумать об оптимизации
		TotalWeekOmissionsResulter TotalWeekOmissionsResulter;
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
				throw new Exception($"Dev:Week has to include 7 days:1042:days.Count={days.Count}");
			string[] week =
			{
				"Понедельник",
				"Вторник",
				"Среда",
				"Четверг",
				"Пятница",
				"Суббота"
			};
			InitStudentsListsAndDictionaries(days);
			TotalWeekOmissionsResulter = new TotalWeekOmissionsResulter(days, studentsCount);
			for (int i = 0; i < 6; ++i)
			{
				WriteDay(days[i].lessons, dateOnly, week[i], weekInd, i, 2);
				WriteWeekLeftHeader(days[i], weekInd);
				WriteWeekRightTrailer(days, weekInd);
				dateOnly = dateOnly.AddDays(1);
			}

			studentsIndByName.Clear();
		}
		private void WriteWeekRightTrailer(List<Day> week, int sheetInd)
		{
			//33 - lk
			var sheet = workbook.Worksheets[sheetInd];
			var Cells = sheet.Cells;
			Cells["AG1:AN1"].Merge = true;
			Cells["AG1"].Value = "Всего пропущено часов";
			Cells["AG2:AJ2"].Merge = true;
			Cells["AK2:AN2"].Merge = true;
			Cells["AG2"].Value = "по уваж пр.";
			Cells["AK2"].Value = "по неуваж пр.";

			Cells[(int)RowType.LessonsTypes, 33].Value = "ЛК";
			Cells[(int)RowType.LessonsTypes, 34].Value = "ЛР";
			Cells[(int)RowType.LessonsTypes, 35].Value = "ПЗ";
			Cells[(int)RowType.LessonsTypes, 36].Value = "ВСЕГО";
			Cells[(int)RowType.LessonsTypes, 37].Value = "ЛК";
			Cells[(int)RowType.LessonsTypes, 38].Value = "ЛР";
			Cells[(int)RowType.LessonsTypes, 39].Value = "ПЗ";
			Cells[(int)RowType.LessonsTypes, 40].Value = "ВСЕГО";

			for(int s = 0; s < studentsCount; ++s)
			{
				Cells[(int)RowType.OmmisionsStart + s, 33].Value =
					TotalWeekOmissionsResulter.studentWeekOmissions[s]
					.StudentOmissionsOnLessonType["ЛК"].respectful;
				Cells[(int)RowType.OmmisionsStart + s, 34].Value =
					TotalWeekOmissionsResulter.studentWeekOmissions[s]
					.StudentOmissionsOnLessonType["ЛР"].unrespectful;
				Cells[(int)RowType.OmmisionsStart + s, 35].Value =
					TotalWeekOmissionsResulter.studentWeekOmissions[s]
					.StudentOmissionsOnLessonType["ПЗ"].respectful;
				Cells[(int)RowType.OmmisionsStart + s, 36].Value = TotalWeekOmissionsResulter.r_StudentTotalOmmisionsOnType(s);
				Cells[(int)RowType.OmmisionsStart + s, 37].Value =
					TotalWeekOmissionsResulter.studentWeekOmissions[s]
					.StudentOmissionsOnLessonType["ЛК"].unrespectful;
				Cells[(int)RowType.OmmisionsStart + s, 38].Value =
					TotalWeekOmissionsResulter.studentWeekOmissions[s]
					.StudentOmissionsOnLessonType["ЛР"].respectful;
				Cells[(int)RowType.OmmisionsStart + s, 39].Value =
					TotalWeekOmissionsResulter.studentWeekOmissions[s]
					.StudentOmissionsOnLessonType["ПЗ"].unrespectful;
				Cells[(int)RowType.OmmisionsStart + s, 40].Value = TotalWeekOmissionsResulter.unr_StudentTotalOmmisionsOnType(s);
			}

			var collumns = sheet.Columns[33, 40];
			collumns.Style.Font.Bold = true;
			collumns.Style.Font.Name = "Times New Roman";
			collumns.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
			collumns.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
			collumns.Width = 5;
			collumns.Style.Border.BorderAround(ExcelBorderStyle.Double, System.Drawing.Color.Black);
			sheet.Columns[36].Width = 7.5;
			sheet.Columns[40].Width = 7.5;
		}

		private int InitStudentsListsAndDictionaries(List<Day> days)
		{
			studentsCount = -1;
			for (int i = 0; i < days.Count; ++i)
				for (int j = 0; j < days[i].lessons.Count; ++j)
					if (days[i].lessons[j].subGroup == 0)
					{
						studentsCount = days[i].lessons[j].students.Count;
						CreateStudentsIndexListByName(days[i].lessons[j].students);
						i = days.Count;
						break;
					}
			if (studentsCount == -1)
				throw new Exception("List of student was not found. Please connect with developer");
			return studentsCount;
		}
		private void CreateStudentsIndexListByName(List<StudentOmission> stunts)
		{
			for (int i = 0; i < stunts.Count; ++i)
			{
				studentsIndByName.Add(stunts[i].fio, i);
			}
		}

		private void WriteWeekLeftHeader(Day day, int sheetInd)
		{
			var sheet = workbook.Worksheets[sheetInd];
			//1
			sheet.Cells[1, 1, (int)RowType.OmmisionsStart - 1, 1].Merge = true;
			int student = 0;
			for (var i = RowType.OmmisionsStart; i < RowType.OmmisionsEnd; ++i)
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
			sheet.Columns[1,2].Style.Font.Name = "Times New Roman";
			sheet.Columns[1,2].Style.Font.Bold = true;
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
				if (lessons[i].lessonTypeAbbrev == "ЛР" && lessons[i].subGroup != 0 && lessons[i].students.Count < studentsCount)
					foreach (StudentOmission student in lessons[i].students)
					{
						int currentRow = studentsIndByName[student.fio] + (int)RowType.OmmisionsStart;
						if(student.omission == null)
							sheet.Cells[currentRow, currentColumn].Value = "";
						else
						{
							sheet.Cells[currentRow, currentColumn].Value = student.omission.missedHours;
							if (student.omission.respectfulOmission == true)
							{
								sheet.Cells[currentRow, currentColumn].Style.Font.Color.SetColor(System.Drawing.Color.Green);
							}
							else
							{
								sheet.Cells[currentRow, currentColumn].Style.Font.Color.SetColor(System.Drawing.Color.Red);
							}
						}

					}//КОСТЫЛЬ
				else
					for (int o = 0; o < (RowType.OmmisionsEnd - RowType.OmmisionsStart) && o < lessons[i].students.Count; ++o)
					{
						if (lessons[i].students[o].omission == null)
							sheet.Cells[o + (int)RowType.OmmisionsStart, currentColumn].Value = "";
						else
						{
							sheet.Cells[o + (int)RowType.OmmisionsStart, currentColumn].Value = lessons[i].students[o].omission.missedHours;
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
