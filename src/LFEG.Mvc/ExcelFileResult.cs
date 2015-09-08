using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace LFEG.Mvc
{
    public abstract class ExcelFileResultBase : FileResult
    {
        protected string FileName;

        protected ExcelFileResultBase(string fileName = "Output")
            : base("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        {
            FileName = fileName;
        }

        protected override void WriteFile(HttpResponseBase response)
        {
            response.AddHeader("content-disposition", string.Format("attachment;  filename={0}.xlsx", FileName));

            var generator = new ExcelFileGenerator();
            GenerateExcelFile(generator, response.OutputStream);
        }

        protected abstract void GenerateExcelFile(ExcelFileGenerator generator, Stream outputStream);
    }

    public class ExcelFileResult<T> : ExcelFileResultBase
    {
        private readonly IEnumerable<T> _data;

        public ExcelFileResult(IEnumerable<T> data, string fileName = "Output")
            : base(fileName)
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

        public ExcelFileResult(DataTable table, string fileName = "Output")
            : base(fileName)
        {
            _table = table;
        }

        protected override void GenerateExcelFile(ExcelFileGenerator generator, Stream outputStream)
        {
            generator.GenerateFile(_table, outputStream);
        }
    }
}