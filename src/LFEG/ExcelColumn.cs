using System;

namespace LFEG
{
    public class ExcelColumn
    {
        public string Caption { get; set; }

        public string ExcelDataType { get; set; }

        public Func<object, string> DataProvider { get; set; }
    }
}