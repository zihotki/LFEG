using System;

namespace LFEG
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelExportAttribute : Attribute
    {
        public int Width { get; set; } = 25;
        public string Caption { get; set; }
        public bool? Encode { get; set; }
        public bool Intern { get; set; } = false;
        public string DataFormat { get; set; }
    }
}