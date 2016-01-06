using System.IO;
using System.Web;
using System.Web.Mvc;

namespace LFEG.Mvc
{
    public abstract class ExcelFileResultBase : FileResult
    {
        private readonly ExcelFileGeneratorSettings _settings;
        protected string FileName;

        protected ExcelFileResultBase(ExcelFileGeneratorSettings settings, string fileName = "Output")
            : base("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            _settings = settings;
            FileName = fileName;
        }

        protected override void WriteFile(HttpResponseBase response)
        {
            response.AddHeader("content-disposition", $"attachment;  filename={FileName}.xlsx");

            var generator = _settings.CreateGenerator();
            GenerateExcelFile(generator, response.OutputStream);
        }

        protected abstract void GenerateExcelFile(ExcelFileGenerator generator, Stream outputStream);
    }
}