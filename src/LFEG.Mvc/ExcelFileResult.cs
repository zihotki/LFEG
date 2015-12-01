using System.Collections.Generic;
using System.Data;
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

    public class ExcelFileResult<T> : ExcelFileResultBase
    {
        private readonly IEnumerable<T> _data;

        public ExcelFileResult(ExcelFileGeneratorSettings settings, IEnumerable<T> data, string fileName = "Output")
            : base(settings, fileName)
        {
            _data = data;
        }

        protected override void GenerateExcelFile(ExcelFileGenerator generator, Stream outputStream)
        {
            generator.GenerateFile(_data, outputStream);
        }
    }

    public class ExcelFileResult : ExcelFileResultBase
    {
        private readonly DataTable _table;

        public ExcelFileResult(ExcelFileGeneratorSettings settings, DataTable table, string fileName = "Output")
            : base(settings, fileName)
        {
            _table = table;
        }

        protected override void GenerateExcelFile(ExcelFileGenerator generator, Stream outputStream)
        {
            generator.GenerateFile(_table, outputStream);
        }
    }
}