using System.Globalization;
using CsvHelper;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var scheduleCreater = new ScheduleCreater();
scheduleCreater.CreateSchedule();
Console.WriteLine("Press Enter to exit");
var _ = Console.ReadLine();

