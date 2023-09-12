using CsvHelper.Configuration;

public class TeacherMapper : ClassMap<Teacher>
{
    public TeacherMapper()
    {
        Map(m => m.Name).Name("Nom");
        Map(m => m.Type).Convert(row => row.Row.GetField<string>("Type") switch
            {
                "s" => TeacherType.Specialist,
                "t" => TeacherType.Teacher,
                _ => TeacherType.Teacher,
            }
        );
        Map(m => m.Specialty).Name("Specialité");
        Map(m => m.AssignedClass).Name("Classe Assignée");
        Map(m => m.Schedule).Convert(map =>
        {
            var columns = new string[] { "Jour 1", "Jour 2", "Jour 3", "Jour 4", "Jour 5" };
            var schedule = new List<List<int>>();
            for (int i = 0; i < columns.Length; i++)
            {
                var columnName = columns[i];
                schedule.Add(map.Row.GetField<string>(columnName)?
                    .Split(",", StringSplitOptions.RemoveEmptyEntries)
                    .Select(day => int.Parse(day))
                    .ToList() ?? new List<int>());
            }

            return schedule;
        });
        Map(m => m.Liberation).Name("Liberation");
    }
}