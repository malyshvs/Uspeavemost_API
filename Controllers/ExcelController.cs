using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Uspevaemost_API.Models;

namespace Uspevaemost_API.Controllers
{
    [ApiController]
    [Route("[controller]")]

    public class ExcelController : Controller
    {
    
        private readonly Services.ReportService _reportService;
        private readonly string conn;
        public ExcelController(Services.ReportService reportService, IConfiguration configuration)
        {
            _reportService = reportService;
            conn = configuration.GetConnectionString("DefaultConnection"); 
        }

        [HttpPost("download")]
        public async Task<IActionResult> DownloadReport([FromBody] Models.ReportRequest request)
        {

            var token = request.token;
            string username = Requests.checkToken(token, conn);

            if (username!=null)
            {
                
                var fileContent = await _reportService.GenerateExcelReportAsync(request, username);

                    return File(fileContent,
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                $"Отчет_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                
                
            }else return Forbid();


        }
    }
}
