using System.IO;
using ClosedXML.Excel;
using Xunit;
using ToggJirazer.Models;
using ToggJirazer.Services;
using ToggJirazer.Tests.Fakes;

namespace ToggJirazer.Tests;

public class ReportXlsxTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "ToggJirazerTests", Guid.NewGuid().ToString("N"));

    public ReportXlsxTests() => Directory.CreateDirectory(_dir);

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); } catch { /* best effort */ }
    }

    private static void Write(string path, List<(string, List<ReportRow>)> sheets)
    {
        var service = new ReportService(new JiraService(new JiraConfig(), new FakeJiraTools((_, _) => default)));
        service.WriteXlsx(sheets, [], [], [], [], [], path);
    }

    [Fact]
    public void User_sheets_survive_regeneration_and_generated_sheets_are_replaced()
    {
        var path = Path.Combine(_dir, "report.xlsx");
        Write(path, [("March 2026", []), ("April 2026", [])]);

        using (var wb = new XLWorkbook(path))
        {
            wb.Worksheets.Add("Noter").Cell(1, 1).Value = "Mine noter";
            wb.Worksheets.Add("Budget 2027").Cell("B2").FormulaA1 = "1+1";
            wb.Save();
        }

        Write(path, [("Rapport", [])]);

        using var result = new XLWorkbook(path);
        var names = result.Worksheets.Select(ws => ws.Name).ToList();
        Assert.Equal(["Rapport", "Versionsrapport", "Opsummering", "Plan", "Noter", "Budget 2027"], names);
        Assert.Equal("Mine noter", result.Worksheet("Noter").Cell(1, 1).GetString());
        Assert.Equal("1+1", result.Worksheet("Budget 2027").Cell("B2").FormulaA1);
    }

    [Theory]
    [InlineData("Rapport", true)]
    [InlineData("plan", true)]
    [InlineData("Version Report", true)]
    [InlineData("Intet budget", true)]
    [InlineData("January 2026", true)]
    [InlineData("Noter", false)]
    [InlineData("Plan 2027", false)]
    public void Generated_sheet_names_are_recognised(string name, bool generated) =>
        Assert.Equal(generated, ReportService.IsGeneratedSheetName(name));
}
