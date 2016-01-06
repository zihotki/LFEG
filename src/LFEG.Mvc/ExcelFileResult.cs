using System.Collections.Generic;
using System.Data;
using System.IO;

namespace LFEG.Mvc
{
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