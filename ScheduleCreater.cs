using System.Globalization;
using CsvHelper;
using OfficeOpenXml;

public class ScheduleCreater
{
    const int MAX_NUMBER_OF_PERIODS = 5;
    const int MAX_NUMBER_OF_DAYS = 5;

    public void CreateSchedule()
    {
        // Read CSV files
        var teachers = GetTeachers();
        var specialistRequirements = GetSpecialistRequirements();

        // Initialize the schedule
        var schedules = new Dictionary<string, List<List<SpecialistAvailability?>>>();
        // List all assigned teachers
        var assignedTeachers = teachers.Where(t => t.Type == TeacherType.Teacher).ToList();
        // Get all priority assignments
        var priorityRequirements = GetPriorityRequirements(teachers);

        // Get a List of classes
        // TODO: Maybe we can review this and use an explicit list of class?
        var classes = specialistRequirements
            .GroupBy(sr => sr.ClassNumber)
            .Select(grp => grp.Key)
            .ToList();

        // Generate empty schedule
        foreach (var classNumber in classes)
        {
            schedules[classNumber] = Enumerable.Range(0, MAX_NUMBER_OF_DAYS).Select(day =>
            {
                return Enumerable.Range(0, MAX_NUMBER_OF_PERIODS).Select(period =>
                {
                    return null as SpecialistAvailability;
                }).ToList();
            }).ToList();
        }

        // var expandedSpecialistRequirements = new List<SpecialistRequirement>();
        // specialistRequirements.ForEach(sr =>
        // {
        //     for (int i = 0; i < sr.WeeklyRequirement; i++)
        //     {
        //         expandedSpecialistRequirements.Add(new SpecialistRequirement
        //         {
        //             ClassNumber = sr.ClassNumber,
        //             Specialty = sr.Specialty,
        //             WeeklyRequirement = 1,
        //         });
        //     }
        // });

        // Create a flat list of requirements by duplicating it the number of time a week it is needed.
        var flattenedSpecialistRequirements = FlattenSpecialistRequirements(specialistRequirements);

        var specialistAvailabilities = new List<SpecialistAvailability>();
        teachers.Where(t => t.Type == TeacherType.Specialist).ToList().ForEach(t =>
        {
            for (int i = 0; i < t.Schedule.Count; i++)
            {
                var daySchedule = t.Schedule[i];
                for (int j = 0; j < daySchedule.Count; j++)
                {
                    var day = i + 1;
                    var period = daySchedule[j];
                    specialistAvailabilities.Add(new SpecialistAvailability(t.Name, t.Specialty, day, period));
                }
            }
        });

        if (specialistAvailabilities.Count < flattenedSpecialistRequirements.Count)
        {
            // TODO: Be more explicit on what we are missing.
            // TODO: Advanced: let it do its job and return a special kind of result that indicates the assignment is complete, but it is missing resources.
            throw new Exception("There are not enough availabilities to fill all the requirements. Please Add availabilities");
        }

        // Start backtracking
        if (Backtrack(schedules, specialistAvailabilities, flattenedSpecialistRequirements, priorityRequirements, 0))
        {
            Console.WriteLine("Solution found!");
            OutputToExcel(schedules);
            // TODO: Output one sheet per teacher with classNumbers
        }
        else
        {
            Console.WriteLine("No solution found.");
        }
    }

    List<Teacher> GetTeachers()
    {
        var filePath = "/home/paulraph/projects/specialist-class-scheduling/MockData/teachers.csv";
        Console.WriteLine($"Reading Teachers at {filePath}");

        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Context.RegisterClassMap<TeacherMapper>();
        var teachers = csv.GetRecords<Teacher>().ToList();

        Console.WriteLine($"Found {teachers.Count} teacher(s)");
        return teachers;
    }

    List<SpecialistRequirement> GetSpecialistRequirements()
    {
        var filePath = "/home/paulraph/projects/specialist-class-scheduling/MockData/specialist_requirements.csv";
        Console.WriteLine($"Reading Specialist Requirements at {filePath}");

        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Context.RegisterClassMap<SpecialistRequirementMapper>();

        var specialistRequirements = csv.GetRecords<SpecialistRequirement>().ToList();

        Console.WriteLine($"Found {specialistRequirements.Count} specialist requirement(s)");

        return specialistRequirements;
    }

    private List<SpecialistRequirement> FlattenSpecialistRequirements(List<SpecialistRequirement> specialistRequirements) {
        return specialistRequirements
            .SelectMany(requirement => Enumerable.Range(0, requirement.WeeklyRequirement).Select(_ => new SpecialistRequirement{
                ClassNumber = requirement.ClassNumber,
                Specialty = requirement.Specialty,
                WeeklyRequirement = 1,
            }))
            .ToList();
    }

    private List<PriorityRequirement> GetPriorityRequirements(List<Teacher> teachers)
    {
        return teachers
            .Where(t => t.Type == TeacherType.Teacher)
            .Where(t => t.Liberation != null)
            .Where(t => t.AssignedClass != null)
            .Select(t => new PriorityRequirement(t.AssignedClass!, t.Liberation!.Value))
            .ToList();
    }

    public bool Backtrack(Dictionary<string, List<List<SpecialistAvailability?>>> schedules,
                          List<SpecialistAvailability> specialistAvailabilities,
                          List<SpecialistRequirement> expandedSpecialistRequirements,
                          List<PriorityRequirement> priorityRequirements,
                          int currentRequirementIndex)
    {
        // Base case: If all requirements are filled, return true
        if (currentRequirementIndex >= expandedSpecialistRequirements.Count)
        {
            return true;
        }

        var currentRequirement = expandedSpecialistRequirements[currentRequirementIndex];
        var classNumber = currentRequirement.ClassNumber;
        var specialty = currentRequirement.Specialty;

        // First, try to fill liberation days
        foreach (var priorityRequirement in priorityRequirements)
        {
            var priorityDay = priorityRequirement.Day;
            var availabilityForLiberation = specialistAvailabilities.FirstOrDefault(a => a.Day == priorityDay && a.Specialty == specialty);

            if (availabilityForLiberation != null)
            {
                // Make the move
                schedules[classNumber][priorityDay - 1][availabilityForLiberation.Period - 1] = availabilityForLiberation;

                // Remove the used availability
                var newSpecialistAvailabilities = new List<SpecialistAvailability>(specialistAvailabilities);
                newSpecialistAvailabilities.Remove(availabilityForLiberation);

                // Remove the priority requirement
                var newPriorityRequirements = new List<PriorityRequirement>(priorityRequirements);
                newPriorityRequirements.Remove(priorityRequirement);

                // Recur to the next requirement
                if (Backtrack(schedules, newSpecialistAvailabilities, expandedSpecialistRequirements, newPriorityRequirements, currentRequirementIndex + 1))
                {
                    return true;
                }

                // If we reach here, we need to backtrack
                schedules[classNumber][priorityDay - 1][availabilityForLiberation.Period - 1] = null;
            }
        }

        // Then, proceed with the general backtracking logic
        foreach (var availability in specialistAvailabilities.Where(a => a.Specialty == specialty).ToList())
        {
            if (IsValid(schedules, availability, classNumber))
            {
                // Make the move
                schedules[classNumber][availability.Day - 1][availability.Period - 1] = availability;

                // Remove the used availability
                var newSpecialistAvailabilities = new List<SpecialistAvailability>(specialistAvailabilities);
                newSpecialistAvailabilities.Remove(availability);

                // Recur to the next requirement
                var newPriorityRequirements = new List<PriorityRequirement>();
                if (Backtrack(schedules, newSpecialistAvailabilities, expandedSpecialistRequirements, newPriorityRequirements, currentRequirementIndex + 1))
                {
                    return true;
                }

                // If we reach here, we need to backtrack
                schedules[classNumber][availability.Day - 1][availability.Period - 1] = null;
            }
        }

        return false;
    }

    public bool IsValid(Dictionary<string, List<List<SpecialistAvailability?>>> schedules,
                        SpecialistAvailability availability,
                        string classNumber)
    {
        var currentSchedule = schedules[classNumber];

        // Check if there is already a specialist scheduled that day
        if (currentSchedule[availability.Day - 1].Any(period => period != null))
        {
            return false;
        }

        // Check if this slot is already filled
        if (schedules[classNumber][availability.Day - 1][availability.Period - 1] != null)
        {
            return false;
        }

        return true;
    }

    public void OutputToExcel(Dictionary<string, List<List<SpecialistAvailability?>>> schedules)
    {
        var outputPath = "/home/paulraph/projects/specialist-class-scheduling/Results/";
        using (ExcelPackage excel = new ExcelPackage())
        {
            foreach (var classSchedule in schedules)
            {
                var classNumber = classSchedule.Key;
                var schedule = classSchedule.Value;

                // Create a new worksheet for each class
                var worksheet = excel.Workbook.Worksheets.Add(classNumber);

                // Add headers for days and periods
                string[] days = { "Day1", "Day2", "Day3", "Day4", "Day5" };
                string[] periods = { "Period1", "Period2", "Period3", "Period4", "Period5" };

                for (int i = 0; i < days.Length; i++)
                {
                    worksheet.Cells[1, i + 2].Value = days[i];
                }

                for (int i = 0; i < periods.Length; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = periods[i];
                }

                // Fill in the schedule
                for (int day = 0; day < schedule.Count; day++)
                {
                    for (int period = 0; period < schedule[day].Count; period++)
                    {
                        var cellValue = schedule[day][period]?.Name ?? "None";
                        worksheet.Cells[period + 2, day + 2].Value = cellValue;

                        // Optional: Add color coding or other formatting here
                    }
                }
            }

            // Save the Excel file
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = Path.Combine(outputPath, $"schedule_{timestamp}.xlsx");
            FileInfo excelFile = new FileInfo(path);
            excel.SaveAs(excelFile);
        }
    }

}
