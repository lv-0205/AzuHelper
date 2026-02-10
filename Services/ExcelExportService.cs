using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using AzuHelper_v2.Models;

namespace AzuHelper_v2.Services;

public static class ExcelExportService
{
    public static string? CreateTimesheet(AppConfig config, IEnumerable<DayEntry> days)
    {
        try
        {
            var templatePath = PathService.ResourcePath(Path.Combine("src", "tmp.xlsx"));
            if (!File.Exists(templatePath))
            {
                return null;
            }

            var week = ISOWeek.GetWeekOfYear(DateTime.Today);
            var tokens = BuildTokens(config, week);
            var filename = TemplateService.Apply(config.FileNameTemplate, tokens);
            filename = EnsureXlsxExtension(SanitizeFileName(filename));
            var outputPath = Path.Combine(PathService.SavesDirectory(), filename);

            if (File.Exists(outputPath))
            {
                try
                {
                    File.Delete(outputPath);
                }
                catch
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
                    filename = EnsureXlsxExtension(SanitizeFileName($"{filename}_{timestamp}"));
                    outputPath = Path.Combine(PathService.SavesDirectory(), filename);
                }
            }

            File.Copy(templatePath, outputPath, overwrite: true);

            using (var workbook = new XLWorkbook(outputPath))
            {
                var sheet = workbook.Worksheets.First();

                sheet.Cell("D7").Value = config.Name;
                sheet.Cell("E7").Value = config.Vorname;
                sheet.Cell("F7").Value = config.Berufsgruppe;
                sheet.Cell("G7").Value = week;
                sheet.Cell("H7").Value = config.Stammnummer;
                sheet.Cell("I7").Value = config.Ausbilder;

                foreach (var day in days)
                {
                    if (!day.Enabled)
                    {
                        continue;
                    }

                    var col = DayOfWeekToColumn(day.Date.DayOfWeek);
                    if (col is null)
                    {
                        continue;
                    }

                    sheet.Cell(10, col.Value).Value = day.StartTime;
                    sheet.Cell(11, col.Value).Value = day.EndTime;
                    sheet.Cell(12, col.Value).Value = day.Duration;
                    sheet.Cell(13, col.Value).Value = "HO";
                }

                workbook.Save();
            }

            return outputPath;
        }
        catch
        {
            return null;
        }
    }

    private static Dictionary<string, string> BuildTokens(AppConfig config, int week)
    {
        var today = DateTime.Today;
        var fullname = $"{config.Vorname} {config.Name}".Trim();
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["name"] = config.Name,
            ["vorname"] = config.Vorname,
            ["fullname"] = fullname,
            ["week"] = week.ToString(CultureInfo.InvariantCulture),
            ["kw"] = week.ToString(CultureInfo.InvariantCulture),
            ["year"] = today.Year.ToString(CultureInfo.InvariantCulture),
            ["date"] = today.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture)
        };
    }

    private static string EnsureXlsxExtension(string fileName)
        => fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? fileName
            : fileName + ".xlsx";

    private static string SanitizeFileName(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sanitized = fileName;
        foreach (var ch in invalid)
        {
            sanitized = sanitized.Replace(ch, '_');
        }

        return string.IsNullOrWhiteSpace(sanitized) ? "timesheet.xlsx" : sanitized;
    }

    private static int? DayOfWeekToColumn(DayOfWeek dayOfWeek)
    {
        return dayOfWeek switch
        {
            DayOfWeek.Monday => 4,
            DayOfWeek.Tuesday => 5,
            DayOfWeek.Wednesday => 6,
            DayOfWeek.Thursday => 7,
            DayOfWeek.Friday => 8,
            DayOfWeek.Saturday => 9,
            DayOfWeek.Sunday => 10,
            _ => null
        };
    }
}
