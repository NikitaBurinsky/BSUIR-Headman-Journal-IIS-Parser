using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HLApi
{
	class TotalWeekOmissionsResulter
	{
		public List<StudentWeekOmissions> studentWeekOmissions = new List<StudentWeekOmissions>(); 
		public TotalWeekOmissionsResulter(List<Day> Week, int StudentsCount, Dictionary<string, int> studIndexByNames)
		{
			for (int i = 0; i < StudentsCount; ++i)
				studentWeekOmissions.Add(new StudentWeekOmissions());

			foreach (Day day in Week)
				foreach (Lesson lesson in day.lessons)
					for (int s = 0; s < lesson.students.Count; ++s)
					{
						var omission = lesson.students[s];
						int so = studIndexByNames[omission.fio];
						if (omission.omission != null && lesson.nameAbbrev != "К.Ч.")
						{
							float skipedHours = omission.omission.missedHours;
							(float r, float ur) temp = (studentWeekOmissions[so].StudentOmissionsOnLessonType[lesson.lessonTypeAbbrev].respectful
									 , studentWeekOmissions[so].StudentOmissionsOnLessonType[lesson.lessonTypeAbbrev].unrespectful);
							if (omission.omission.respectfulOmission)
								temp.r += skipedHours;
							else
								temp.ur += skipedHours;
							studentWeekOmissions[so].StudentOmissionsOnLessonType[lesson.lessonTypeAbbrev] = temp;
						}
					}
		}
		public void Clear()
		{
			studentWeekOmissions.Clear();
		}

		public float r_StudentTotalOmmisionsOnType(int studentInd) => studentWeekOmissions[studentInd].TotalOmissions.Item1;
		public float unr_StudentTotalOmmisionsOnType(int studentInd) => studentWeekOmissions[studentInd].TotalOmissions.Item2;



		public record StudentWeekOmissions()
		{
			public Dictionary<string, (float respectful, float unrespectful)> StudentOmissionsOnLessonType { get; set; }
				= new Dictionary<string, (float respectful, float unrespectful)>()
				{
				{ "ЛК", (0,0)},
				{ "ЛР", (0,0)},
				{ "ПЗ", (0,0)},
				};

		public (float, float) TotalOmissions
			{
				get
				{
					(float r, float ur) res = StudentOmissionsOnLessonType["ЛК"];
					res.r += StudentOmissionsOnLessonType["ЛР"].respectful;
					res.ur += StudentOmissionsOnLessonType["ЛР"].unrespectful;
					res.r += StudentOmissionsOnLessonType["ПЗ"].respectful;
					res.ur += StudentOmissionsOnLessonType["ПЗ"].unrespectful;
					return res;
				}
			}
		}
	}
}
