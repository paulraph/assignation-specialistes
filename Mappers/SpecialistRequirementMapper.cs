using CsvHelper.Configuration;

public class SpecialistRequirementMapper : ClassMap<SpecialistRequirement>
{
    public SpecialistRequirementMapper()
    {
        Map(m => m.ClassNumber).Name("Classe");
        Map(m => m.Specialty).Name("Specialité");
        Map(m => m.WeeklyRequirement).Name("Requis par semaine");
    }
}