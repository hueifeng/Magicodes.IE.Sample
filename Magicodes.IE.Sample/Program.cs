using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

var builder = WebApplication.CreateBuilder(args);

var app = builder.Build();

app.MapGet("/download", async () =>
{
    IExporter exporter = new ExcelExporter();
    var result = await exporter.ExportAsByteArray(new List<StudentDto>()
                {
                    new StudentDto
                    {
                        Name = "MR.A",
                        Age = 18,
                        Remarks = "�ҽ�MR.A,����18��",
                        Birthday=DateTime.Now
                    },
                    new StudentDto
                    {
                        Name = "MR.B",
                        Age = 19,
                        Remarks = "�ҽ�MR.B,����19��",
                        Birthday=DateTime.Now
                    }
                });
    return Results.File(result, "application/octet-stream", "abc.xlsx");
});

app.MapGet("/download/excelattr", async () =>
{
    IExporter exporter = new ExcelExporter();
    var result = await exporter.ExportAsByteArray(new List<StudentAttrDto>()
                {
                    new StudentAttrDto
                    {
                        Name = "MR.A",
                        Age = 18,
                        Remarks = "�ҽ�MR.A,����18��",
                        Birthday=DateTime.Now
                    },
                    new StudentAttrDto
                    {
                        Name = "MR.B",
                        Age = 19,
                        Remarks = "�ҽ�MR.B,����19��",
                        Birthday=DateTime.Now
                    },
                    new StudentAttrDto
                    {
                        Name = "MR.C",
                        Age = 20,
                        Remarks = "�ҽ�MR.C,����20��",
                        Birthday=DateTime.Now
                    }
                });
    return Results.File(result, "application/octet-stream", "abc.xlsx");
});

app.MapGet("/getstudent", async () =>
{
    var importer = new ExcelImporter();
    using (var stream = new FileStream(Path.Combine("excel.xlsx"), FileMode.Open))
    {
        var result = await importer.Import<StudentAttrDto>(stream);
        return Results.Ok(result.Data);
    }
});

app.Run();

public class StudentDto
{
    public string Name { get; set; }
    public int Age { get; set; }
    public string Remarks { get; set; }
    public DateTime Birthday { get; set; }
}

public class StudentAttrDto
{

    [ExporterHeader("����")]
    [ImporterHeader(Name ="����")]
    public string Name { get; set; }

    [ExporterHeader("����")]
    [ImporterHeader(Name = "����")]
    public int Age { get; set; }
    public string Remarks { get; set; }

    [ExporterHeader(DisplayName = "��������", Format = "yyyy-mm-dd")]
    [ImporterHeader(Name = "��������")]
    public DateTime Birthday { get; set; }
}