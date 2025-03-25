using System.Text.Json;

namespace HLApi
{
	public record Day
	{
		public List<Lesson> lessons { get; set; }
	}
	public record LessonPeriod
	{
		public float lessonPeriodHours { get; set; }
		public string startTime { get; set; }
		public string endTime { get; set; }
	}

	public record StudentOmission
	{
		public int id { get; set; }
		public string fio { get; set; }
		public OmissionDetails omission { get; set; }
	}

	public record OmissionDetails
	{
		public int id { get; set; }
		public int missedHours { get; set; }
		public bool respectfulOmission { get; set; }
	}

	public record Lesson
	{
		public int id { get; set; }
		public string dateString { get; set; }
		public string nameAbbrev { get; set; }
		public string lessonTypeAbbrev { get; set; }
		public LessonPeriod? lessonPeriod { get; set; }
		public int subGroup { get; set; }
		public List<StudentOmission>? students { get; set; }
	}
}
