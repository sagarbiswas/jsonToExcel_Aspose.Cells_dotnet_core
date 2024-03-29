using Aspose.Cells.Utility;
using Aspose.Cells;
using Microsoft.AspNetCore.Mvc;

namespace JsonToExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUtilityController : ControllerBase
    {

    [HttpPost("DownloadExcel")]
    public Task<IActionResult> DownloadExcel(ExcelUtilModel excelUtil)
    {

        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        CellsFactory factory = new CellsFactory();
        Style style = factory.CreateStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;

        // Set JsonLayoutOptions
        JsonLayoutOptions options = new JsonLayoutOptions();
        options.TitleStyle = style;
        options.ArrayAsTable = true;

            // Import JSON Data
            JsonUtility.ImportData(excelUtil.JsonInput, worksheet.Cells, 0, 0, options);

            MemoryStream ms = new MemoryStream();
        workbook.Save(ms, SaveFormat.Excel97To2003);
        ms.Seek(0, SeekOrigin.Begin);

        byte[] buffer = new byte[(int)ms.Length];
        buffer = ms.ToArray();
        if (buffer == null)
            return Task.FromResult<IActionResult>(NotFound()); // returns a NotFoundResult with Status404NotFound response.

        return Task.FromResult<IActionResult>(File(buffer, "application/octet-stream", "xyz.xls")); // returns a FileStreamResult
    }

    }
}

