using CsvHelper.Configuration;

public class ClassMapper : ClassMap<Class>
{
    public ClassMapper()
    {
        Map(m => m.TeacherName).Name("Titulaire");
        Map(m => m.ClassNumber).Name("Numéro Classe");
        Map(m => m.Liberations).Convert(map =>
            map.Row.GetField<string>("Libérations")
                ?.Split(",", StringSplitOptions.RemoveEmptyEntries)
                .Select(day => int.Parse(day))
                .ToList() ?? new List<int>()
        );
    }
}