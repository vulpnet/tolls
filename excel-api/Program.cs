using ClosedXML.Excel;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.Caching.Memory;

var builder = WebApplication.CreateBuilder(args);

// Render (và nhiều PaaS khác) set biến PORT để chỉ định cổng lắng nghe — không cố định port trong image
var port = Environment.GetEnvironmentVariable("PORT") ?? "5199";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

builder.Services.AddOpenApi();
builder.Services.AddMemoryCache();
builder.Services.Configure<FormOptions>(opts =>
{
    // cho phép upload tới 500MB — vượt xa giới hạn 150MB của bản xử lý client-side (SheetJS trong trình duyệt)
    opts.MultipartBodyLengthLimit = 500L * 1024 * 1024;
});
builder.WebHost.ConfigureKestrel(opts =>
{
    // Kestrel mặc định giới hạn request body ở mức thấp hơn nhiều so với MultipartBodyLengthLimit — phải nới riêng
    opts.Limits.MaxRequestBodySize = 500L * 1024 * 1024;
});

builder.Services.AddCors(opts =>
{
    opts.AddDefaultPolicy(policy =>
    {
        // API công khai phục vụ các bản static site (Firebase/Vercel/Render) khác domain — không có cookie/credential nhạy cảm
        policy.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader();
    });
});

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseCors();

const long MaxUploadBytes = 500L * 1024 * 1024;
var sessions = app.Services.GetRequiredService<IMemoryCache>();

app.MapPost("/api/excel/upload", async (HttpRequest request) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Cần gửi dạng multipart/form-data với field 'file'." });

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    if (file is null || file.Length == 0)
        return Results.BadRequest(new { error = "Không tìm thấy file, hoặc file rỗng." });

    if (file.Length > MaxUploadBytes)
        return Results.BadRequest(new { error = $"File quá lớn ({file.Length / 1024 / 1024}MB). Giới hạn server: {MaxUploadBytes / 1024 / 1024}MB." });

    var sessionId = Guid.NewGuid().ToString("N");

    try
    {
        using var stream = file.OpenReadStream();
        using var workbook = new XLWorkbook(stream);

        var sheetInfos = new List<object>();
        var parsedSheets = new Dictionary<string, ParsedSheet>();

        foreach (var ws in workbook.Worksheets)
        {
            var usedRange = ws.RangeUsed();
            if (usedRange is null)
            {
                sheetInfos.Add(new { name = ws.Name, rowCount = 0, colCount = 0 });
                parsedSheets[ws.Name] = new ParsedSheet(new List<string[]>());
                continue;
            }

            var rows = new List<string[]>(usedRange.RowCount());
            foreach (var row in usedRange.Rows())
            {
                var cells = row.Cells(1, usedRange.ColumnCount())
                    .Select(c => c.GetFormattedString())
                    .ToArray();
                rows.Add(cells);
            }

            parsedSheets[ws.Name] = new ParsedSheet(rows);
            sheetInfos.Add(new { name = ws.Name, rowCount = rows.Count, colCount = usedRange.ColumnCount() });
        }

        // giữ dữ liệu đã parse trong memory cache 30 phút — đủ để user phân trang/lọc mà không phải upload lại
        sessions.Set(sessionId, parsedSheets, TimeSpan.FromMinutes(30));

        return Results.Ok(new { sessionId, sheets = sheetInfos });
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = "Không đọc được file: " + ex.Message });
    }
})
.DisableAntiforgery();

app.MapGet("/api/excel/{sessionId}/sheet/{sheetName}", (string sessionId, string sheetName, int page, int pageSize) =>
{
    if (!sessions.TryGetValue<Dictionary<string, ParsedSheet>>(sessionId, out var sheets) || sheets is null)
        return Results.NotFound(new { error = "Phiên làm việc đã hết hạn hoặc không tồn tại — hãy upload lại file." });

    if (!sheets.TryGetValue(sheetName, out var sheet))
        return Results.NotFound(new { error = $"Không tìm thấy sheet '{sheetName}'." });

    page = Math.Max(page, 1);
    pageSize = Math.Clamp(pageSize <= 0 ? 2000 : pageSize, 1, 20000);

    var totalRows = sheet.Rows.Count;
    var skip = (page - 1) * pageSize;
    var pageRows = sheet.Rows.Skip(skip).Take(pageSize).ToArray();

    return Results.Ok(new
    {
        sheetName,
        page,
        pageSize,
        totalRows,
        totalPages = (int)Math.Ceiling(totalRows / (double)pageSize),
        rows = pageRows
    });
});

app.MapDelete("/api/excel/{sessionId}", (string sessionId) =>
{
    sessions.Remove(sessionId);
    return Results.Ok();
});

app.MapGet("/api/health", () => Results.Ok(new { status = "ok" }));

app.Run();

record ParsedSheet(List<string[]> Rows);
