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
        var classes = GetClasses();
        var specialistRequirements = GetSpecialistRequirements();
        var specialistTeachers = GetSpecialistTeachers();

        // Initialize the schedule
        var schedules = new Dictionary<string, List<List<SpecialistAvailability?>>>();
        // Get all priority assignments
        var priorityRequirements = GetPriorityRequirements(classes);

        // Generate empty schedule for each class
        foreach (var classNumber in classes.Select(c => c.ClassNumber))
        {
            schedules[classNumber] = Enumerable.Range(0, MAX_NUMBER_OF_DAYS).Select(day =>
            {
                return Enumerable.Range(0, MAX_NUMBER_OF_PERIODS).Select(period =>
                {
                    return null as SpecialistAvailability;
                }).ToList();
            }).ToList();
        }

        // Create a flat list of all the requirements for easier processing
        var flattenedSpecialistRequirements = FlattenSpecialistRequirements(specialistRequirements);

        var specialistAvailabilities = new List<SpecialistAvailability>();
        specialistTeachers.ForEach(t =>
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

    List<Class> GetClasses()
    {
        var filePath = "/home/paulraph/projects/specialist-class-scheduling/MockData/classes.csv";
        Console.WriteLine($"Reading Classes at {filePath}");

        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Context.RegisterClassMap<ClassMapper>();
        var classes = csv.GetRecords<Class>().ToList();

        Console.WriteLine($"Found {classes.Count} class(es)");
        return classes;
    }

    List<SpecialistRequirement> GetSpecialistRequirements()
    {
        var filePath = "/home/paulraph/projects/specialist-class-scheduling/MockData/requis_specialistes.csv";
        Console.WriteLine($"Reading Specialist Requirements at {filePath}");

        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Context.RegisterClassMap<SpecialistRequirementMapper>();

        var specialistRequirements = csv.GetRecords<SpecialistRequirement>().ToList();

        Console.WriteLine($"Found {specialistRequirements.Count} specialist requirement(s)");

        return specialistRequirements;
    }

    List<SpecialistTeacher> GetSpecialistTeachers()
    {
        var filePath = "/home/paulraph/projects/specialist-class-scheduling/MockData/specialistes.csv";
        Console.WriteLine($"Reading Specialists at {filePath}");

        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        csv.Context.RegisterClassMap<SpecialistTeacherMapper>();

        var specialistTeachers = csv.GetRecords<SpecialistTeacher>().ToList();

        Console.WriteLine($"Found {specialistTeachers.Count} specialist requirement(s)");

        return specialistTeachers;
    }

    private List<SpecialistRequirement> FlattenSpecialistRequirements(List<SpecialistRequirement> specialistRequirements)
    {
        return specialistRequirements
            .SelectMany(requirement => Enumerable.Range(0, requirement.WeeklyRequirement).Select(_ => new SpecialistRequirement
            {
                ClassNumber = requirement.ClassNumber,
                Specialty = requirement.Specialty,
                WeeklyRequirement = 1,
            }))
            .ToList();
    }

    private List<PriorityRequirement> GetPriorityRequirements(List<Class> classes)
    {
        return classes
            .Where(c => c.Liberations.Any())
            .SelectMany(c => c.Liberations.Select(liberation => new PriorityRequirement(c.ClassNumber, liberation)))
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
                string[] days = { "Jour 1", "Jour 2", "Jour 3", "Jour 4", "Jour 5" };
                string[] periods = { "Période 1", "Période 2", "Période 3", "Période 4", "Période 5" };

                // HEADERS
                worksheet.Cells[1, 1].Value = classNumber;
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(1, 204, 204, 204);

                for (int i = 0; i < days.Length; i++)
                {
                    worksheet.Cells[1, i + 2].Value = days[i];
                    worksheet.Cells[1, i + 2].Style.Font.Bold = true;
                    worksheet.Cells[1, i + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 2].Style.Fill.BackgroundColor.SetColor(1, 204, 204, 204);
                }

                for (int i = 0; i < periods.Length; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = periods[i];
                    worksheet.Cells[i + 2, 1].Style.Font.Bold = true;
                    worksheet.Cells[i + 2, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[i + 2, 1].Style.Fill.BackgroundColor.SetColor(1, 204, 204, 204);
                }

                // Fill in the schedule
                for (int day = 0; day < schedule.Count; day++)
                {
                    for (int period = 0; period < schedule[day].Count; period++)
                    {
                        var cellValue = schedule[day][period]?.Name ?? "";
                        worksheet.Cells[period + 2, day + 2].Value = cellValue;

                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            worksheet.Cells[period + 2, day + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[period + 2, day + 2].Style.Fill.BackgroundColor.SetColor(1, 230, 242, 255);
                        }

                        // Optional: Add color coding or other formatting here
                    }
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            }



            // Save the Excel file
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string path = Path.Combine(outputPath, $"schedule_{timestamp}.xlsx");
            FileInfo excelFile = new FileInfo(path);
            excel.SaveAs(excelFile);
        }
    }

}
