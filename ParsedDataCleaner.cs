using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HLApi
{
	internal class ParsedDataCleaner
	{
		public static void CombineDoublingLessons(HLApi.Day day)
		{
			for (int l = 0; l < day.lessons.Count - 1; ++l)
			{
				Lesson lesson = day.lessons[l];
				Lesson nextLesson = day.lessons[l + 1];//deleting
				if (lesson.lessonPeriod.startTime == nextLesson.lessonPeriod.startTime &&
					lesson.lessonPeriod.endTime == nextLesson.lessonPeriod.endTime &&
					lesson.nameAbbrev == nextLesson.nameAbbrev &&
					lesson.subGroup == nextLesson.subGroup)
				{
					List<StudentOmission> resultMerge = new List<StudentOmission> ();
					for (int i = 0; i < lesson.students.Count; ++i)
					{
						StudentOmission c = lesson.students[i], d = nextLesson.students[i];
						if (d.fio == c.fio)
							if (d.omission == null)
							{
								resultMerge.Add(c);
							}
							else
							{
								resultMerge.Add(d);
							}
						else
							throw new Exception("Dev:CombineDoublingLessins:1532");
					}
					day.lessons.RemoveAt(l+1);
					lesson.students.Clear();
					lesson.students = resultMerge;
				}
			}
		}

	}
}
