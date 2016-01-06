using System;

namespace LFEG.Infrastructure
{
    public class ExcelColumn<T>
    {
        public string Caption { get; set; }
        public string ExcelDataType { get; set; }
        public Func<T, string> DataProvider { get; set; }
        public int? Style { get; set; } = null;
        public int Width { get; set; } = 25;
    }
}