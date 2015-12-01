using System;

namespace LFEG
{
    public class ExcelColumn<T>
    {
        public string Caption { get; set; }

        public string ExcelDataType { get; set; }

        public Func<T, string> DataProvider { get; set; }
    }
}