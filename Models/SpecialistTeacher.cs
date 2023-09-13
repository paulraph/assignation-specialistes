public class SpecialistTeacher
{
    public string Name { get; set; } = string.Empty;

    public string Specialty { get; set; } = string.Empty;

    public List<List<int>> Schedule { get; set; } = new();

    public string? Notes { get; set; }
}