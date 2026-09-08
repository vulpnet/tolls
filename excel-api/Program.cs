using ExcelDataReader;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.Caching.Memory;

// ExcelDataReader cần provider này để decode 1 số encoding (đặc biệt file .xls cũ hoặc ký tự đặc biệt trong .xlsx)
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

var builder = WebApplication.CreateBuilder(args);

// Render (và nhiều PaaS khác) set biến PORT để chỉ định cổng lắng nghe — không cố định port trong image
var port = Environment.GetEnvironmentVariable("PORT") ?? "5199";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

builder.Services.AddOpenApi();
builder.Services.AddMemoryCache();
// Đo thực tế trên Render free tier (512MB RAM, KHÔNG chỉ đo local vì overhead cố định của container
// (OS + .NET runtime + ASP.NET Core hosting) chiếm phần lớn RAM khả dụng): 13.2MB (45k dòng) OK,
// 17.5MB (60k dòng) → OutOfMemoryException. Đặt 15MB — nằm giữa vùng đã xác nhận an toàn/lỗi, có margin.
const long MaxUploadBytes = 15L * 1024 * 1024;
builder.Services.Configure<FormOptions>(opts =>
{
    opts.MultipartBodyLengthLimit = MaxUploadBytes;
});
builder.WebHost.ConfigureKestrel(opts =>
{
    // Kestrel mặc định giới hạn request body ở mức thấp hơn nhiều so với MultipartBodyLengthLimit — phải nới riêng
    opts.Limits.MaxRequestBodySize = MaxUploadBytes;
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
        var sheetInfos = new List<object>();
        var parsedSheets = new Dictionary<string, ParsedSheet>();

        using (var stream = file.OpenReadStream())
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            do
            {
                var rows = new List<string[]>();
                int colCount = 0;
                // reader.Read() trả về từng dòng một (SAX-style) — không dựng object model của cả workbook,
                // nên RAM chỉ tỉ lệ với dữ liệu đang giữ (rows) chứ không cộng thêm styles/formulas/formatting
                while (reader.Read())
                {
                    var cells = new string[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        cells[i] = reader.IsDBNull(i) ? "" : Convert.ToString(reader.GetValue(i)) ?? "";
                    }
                    rows.Add(cells);
                    colCount = Math.Max(colCount, reader.FieldCount);
                }

                var sheetName = reader.Name;
                parsedSheets[sheetName] = new ParsedSheet(rows);
                sheetInfos.Add(new { name = sheetName, rowCount = rows.Count, colCount });
            } while (reader.NextResult()); // chuyển sang sheet tiếp theo trong cùng workbook
        }

        // giữ dữ liệu đã parse trong memory cache 30 phút — đủ để user phân trang/lọc mà không phải upload lại
        sessions.Set(sessionId, parsedSheets, TimeSpan.FromMinutes(30));

        return Results.Ok(new { sessionId, sheets = sheetInfos });
    }
    catch (OutOfMemoryException)
    {
        return Results.BadRequest(new { error =
            $"Server không đủ bộ nhớ để xử lý file này ({file.Length / 1024 / 1024}MB). " +
            "Hãy thử tách bớt sheet/dòng dữ liệu thành file nhỏ hơn." });
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
