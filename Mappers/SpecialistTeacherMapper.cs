using CsvHelper.Configuration;

public class SpecialistTeacherMapper : ClassMap<SpecialistTeacher>
{
    public SpecialistTeacherMapper()
    {
        Map(m => m.Name).Name("Nom");
        Map(m => m.Specialty).Name("Specialité");
        Map(m => m.Schedule).Convert(map =>
        {
            var columns = new string[] 
            {
                "Disponibilité Jour 1",
                "Disponibilité Jour 2",
                "Disponibilité Jour 3",
                "Disponibilité Jour 4",
                "Disponibilité Jour 5",
            };
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
        Map(m => m.Notes);
    }
}