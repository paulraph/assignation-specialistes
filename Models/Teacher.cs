public class Teacher
{
    public string Name { get; set; } = "";

    public TeacherType Type { get; set; } = TeacherType.Teacher;

    public string Specialty { get; set; } = "";

    public string? AssignedClass { get; set; }

    public List<List<int>> Schedule { get; set; } = new List<List<int>>();

    public int? Liberation { get; set; }
}