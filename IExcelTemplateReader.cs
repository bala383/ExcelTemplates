using DocumentFormat.OpenXml.Packaging;

namespace API.Services.Excel
{
    public interface IExcelTemplateReader
    {
        IEnumerable<T> ReadExcelFile(IFormFile file, int productId, CancellationToken token);

        byte[] ExportListToExcel<T>(MemoryStream mem, List<T> lstBookingTemplates, string SheetName, string Columns);
    }
}
