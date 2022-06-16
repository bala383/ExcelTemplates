using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace API
{
    public class Download 
    {
        public async Task<ActionResult> DownloadFIle(List<T> result)
        {
               
            if (result.Any())
            {
                MemoryStream mem = new();
                var mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return File(_excelTemplateReader.ExportListToExcel(mem, result, "{SheetName}", excelTemplate.Columns), mimeType, "FileName.xlsx");
            }
            return new NotFoundObjectResult(new StatusCodeProblemDetails(404,
                                                                              instance,
                                                                             $"result list is empty",
                                                                              traceId));
        }
    }
}
