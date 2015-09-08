using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace LFEG
{
    public class ExcelWriter
    {
        private readonly IXmlEncoder _xmlEncoder;
        private readonly Dictionary<int, string> _columnNamesCache = new Dictionary<int, string>();

        public ExcelWriter( IXmlEncoder xmlEncoder = null)
        {
            _xmlEncoder = xmlEncoder ?? new XmlEncoder();
        }

        public virtual void WriteXmlHeader(StreamWriter writer)
        {
            writer.Write(@"<?xml version=""1.0"" encoding=""utf-8"" ?>");
            writer.Write(
                @"<worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
        }

        public virtual void WriteXmlFooter(StreamWriter writer)
        {
            writer.Write("</worksheet>");
        }

        public virtual void WriteColumnsDefinitions(StreamWriter writer, ExcelColumn[] columns)
        {
            const string colStrFormat = "<col min=\"{0}\" max=\"{0}\" width=\"25\" customWidth=\"1\" />";
            writer.Write("<cols>");

            for (var i = 1; i <= columns.Length; i++)
            {
                writer.Write(colStrFormat, i);
            }

            writer.Write("</cols>");
        }

        public virtual void WriteRows(StreamWriter writer, IEnumerable items, ExcelColumn[] columns)
        {
            var row = 2;
            var cols = columns.Length;

            foreach (var item in items)
            {
                writer.Write("<row r=\"{0}\" >", row);

                for (var colNumber = 1; colNumber <= cols; colNumber++)
                {
                    var column = columns[colNumber - 1];
                    var value = _xmlEncoder.Encode(column.DataProvider(item));
                    if (column.ExcelDataType == ExcelDataTypes.String)
                    {
                        writer.Write("<c r=\"{0}{1}\" s=\"0\" t=\"{2}\"><is><t>{3}</t></is></c>",
                        GetColumnName(colNumber), row, column.ExcelDataType, value);
                    }
                    else
                    {
                        writer.Write("<c r=\"{0}{1}\" s=\"0\" t=\"{2}\"><v>{3}</v></c>",
                        GetColumnName(colNumber), row, column.ExcelDataType, value);
                    }
                }
                writer.Write("</row>");

                row++;
            }
        }

        public virtual void WriteTitleRow(StreamWriter writer, ExcelColumn[] columns)
        {
            writer.Write("<row r=\"{0}\" >", 1);

            var cols = columns.Length;
            for (var i = 1; i <= cols; i++)
            {
                writer.Write("<c r=\"{0}{1}\" t=\"str\"><v>{2}</v></c>", GetColumnName(i), 1,
                    _xmlEncoder.Encode(columns[i - 1].Caption));
            }
            writer.Write("</row>");
        }

        public virtual void WriteSheetStart(StreamWriter writer)
        {
            writer.Write("<sheetData>");
        }

        public virtual void WriteSheetEnd(StreamWriter writer)
        {
            writer.Write("</sheetData>");
        }

        protected virtual string GetColumnName(int column)
        {
            string name;
            if (_columnNamesCache.TryGetValue(column, out name))
            {
                return name;
            }

            name = ExcelColumnFromNumber(column);
            _columnNamesCache[column] = name;

            return name;
        }

        public virtual string ExcelColumnFromNumber(int column)
        {
            var columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                var currentLetterNumber = (columnNumber - 1) % 26;
                var currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }
    }
}