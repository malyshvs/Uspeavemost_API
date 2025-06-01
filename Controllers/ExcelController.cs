using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Uspevaemost_API.Models;

namespace Uspevaemost_API.Controllers
{
    [ApiController]
    [Route("[controller]")]
    [Authorize]
    public class ExcelController : Controller
    {
    
        private readonly Services.ReportService _reportService;
        public ExcelController(Services.ReportService reportService)
        {
            _reportService = reportService;
        }

        [HttpPost("download")]
        public async Task<IActionResult> DownloadReport([FromBody] Models.ReportRequest request)
        {
            
            var username = User.Identity?.Name;
            username = username.Split("\\")[1];
            for (var i = 0; i < 100; i++)
            {
                System.Diagnostics.Debug.WriteLine(username);

            }
            if (username!=null)
            {
                string uchps = Requests.uchp(username);
                System.Diagnostics.Debug.WriteLine(uchps);
                if (uchps != "")
                {
                    var fileContent = await _reportService.GenerateExcelReportAsync(request, uchps);

                    return File(fileContent,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                $"Отчет_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                }
                else return Forbid();
            }else return NotFound();


        }
    }
}
