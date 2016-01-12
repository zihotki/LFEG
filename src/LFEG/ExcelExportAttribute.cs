using System;

namespace LFEG
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelExportAttribute : Attribute
    {
        /// <summary>
        /// Width of the cell in characters. Default width is 25 characters.
        /// </summary>
        public int Width { get; set; } = 25;

        /// <summary>
        /// Column caption for the property
        /// </summary>
        public string Caption { get; set; }

        /// <summary>
        /// Whether to apply IXmlEncoder.Encode to the data or not. Don't set it to false if you can not 
        /// guarantee that the data won't have any xml special characters.
        /// </summary>
        public bool? Encode { get; set; }

        /// <summary>
        /// Whether to intern data or not. Used only for string data. Interning means that the value will be 
        /// put into Shared Strings table and a reference to that value will be used for a cell instead
        /// of inlining the string data. Interning allows to reduce the size of a result file if it's applied
        /// to data containing a lot of repetitive data. On the other hand, interning unique data will 
        /// increase size of a result file. By default only enums are interned.
        /// </summary>
        public bool Intern { get; set; } = false;

        /// <summary>
        /// Data format for the column. See https://support.office.com/en-my/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 
        /// for details.
        /// </summary>
        public string DataFormat { get; set; }
    }
}