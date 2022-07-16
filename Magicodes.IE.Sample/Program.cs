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
                        Remarks = "我叫MR.A,今年18岁",
                        Birthday=DateTime.Now
                    },
                    new StudentDto
                    {
                        Name = "MR.B",
                        Age = 19,
                        Remarks = "我叫MR.B,今年19岁",
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
                        Remarks = "我叫MR.A,今年18岁",
                        Birthday=DateTime.Now
                    },
                    new StudentAttrDto
                    {
                        Name = "MR.B",
                        Age = 19,
                        Remarks = "我叫MR.B,今年19岁",
                        Birthday=DateTime.Now
                    },
                    new StudentAttrDto
                    {
                        Name = "MR.C",
                        Age = 20,
                        Remarks = "我叫MR.C,今年20岁",
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

    [ExporterHeader("姓名")]
    [ImporterHeader(Name ="姓名")]
    public string Name { get; set; }

    [ExporterHeader("年龄")]
    [ImporterHeader(Name = "年龄")]
    public int Age { get; set; }
    public string Remarks { get; set; }

    [ExporterHeader(DisplayName = "出生日期", Format = "yyyy-mm-dd")]
    [ImporterHeader(Name = "出生日期")]
    public DateTime Birthday { get; set; }
}