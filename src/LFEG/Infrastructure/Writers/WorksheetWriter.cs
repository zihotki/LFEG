using System.Collections.Generic;
using System.IO;
using System.Text;
using Ionic.Zip;

namespace LFEG.Infrastructure.Writers
{
    /*

        For more information about OOXML please see http://officeopenxml.com/ 
        or http://www.ecma-international.org/publications/standards/Ecma-376.htm $3.8.31

    */
    public class WorksheetWriter : IWorksheetWriter
    {
        private readonly IXmlEncoder _xmlEncoder;
        private readonly ISharedStringsInterner _interner;
        private readonly Dictionary<int, string> _columnNamesCache = new Dictionary<int, string>();

        public WorksheetWriter(IXmlEncoder xmlEncoder, ISharedStringsInterner interner)
        {
            _xmlEncoder = xmlEncoder;
            _interner = interner;
        }

        public virtual void Write<T>(ZipFile zip, IEnumerable<T> items, ExcelColumn<T>[] columns)
        {
            zip.AddEntry("xl\\worksheets\\sheet.xml", (name, entryStream) =>
            {
                using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
                {
                    WriteContent(writer, items, columns);
                }
            });
        }

        public virtual void WriteContent<T>(StreamWriter writer, IEnumerable<T> items, ExcelColumn<T>[] columns)
        {
            WriteXmlHeader(writer);
            WriteColumnsDefinitions(writer, columns);
            WriteSheetStart(writer);

            WriteTitleRow(writer, columns);
            WriteRows(writer, items, columns);

            WriteSheetEnd(writer);
            WriteXmlFooter(writer);
        }

        public virtual void WriteXmlHeader(StreamWriter writer)
        {
            writer.Write(@"<?xml version=""1.0"" encoding=""utf-8"" ?>");
            writer.Write(@"<worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
        }

        public virtual void WriteXmlFooter(StreamWriter writer)
        {
            writer.Write("</worksheet>");
        }

        public virtual void WriteColumnsDefinitions<T>(StreamWriter writer, ExcelColumn<T>[] columns)
        {
            const string colStrFormat = "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" customWidth=\"1\" />";
            writer.Write("<cols>");

            for (var i = 1; i <= columns.Length; i++)
            {
                writer.Write(colStrFormat, i, columns[i-1].Width);
            }

            writer.Write("</cols>");
        }

        public virtual void WriteRows<T>(StreamWriter writer, IEnumerable<T> items, ExcelColumn<T>[] columns)
        {
            var row = 2;
            var cols = columns.Length;

            foreach (var item in items)
            {
                writer.Write("<row r=\"{0}\" >", row);

                for (var colNumber = 1; colNumber <= cols; colNumber++)
                {
                    var column = columns[colNumber - 1];
                    var value = column.DataProvider(item);

                    writer.Write("<c r=\"");
                    writer.Write(GetColumnName(colNumber));
                    writer.Write(row);
                    writer.Write("\"");
                    
                    if (column.Style != null)
                    {
                        writer.Write(" s=\"");
                        writer.Write(column.Style.Value);
                        writer.Write("\"");
                    }

                    writer.Write(" t=\"");
                    writer.Write(column.ExcelDataType);
                    writer.Write("\">");


                    if (column.ExcelDataType == ExcelDataTypes.InlineString)
                    {
                        value = _xmlEncoder.Encode(value);

                        writer.Write("<is><t>");
                        writer.Write(value);
                        writer.Write("</t></is>");
                    }
                    else
                    {
                        if (column.ExcelDataType == ExcelDataTypes.String)
                        {
                            value = _interner.Intern(value).ToString();
                        }

                        writer.Write("<v>");
                        writer.Write(value);
                        writer.Write("</v>");
                    }

                    writer.Write("</c>");
                }

                writer.Write("</row>");
                row++;
            }
        }

        public virtual void WriteTitleRow<T>(StreamWriter writer, ExcelColumn<T>[] columns)
        {
            writer.Write("<row r=\"{0}\" >", 1);

            var cols = columns.Length;
            for (var i = 1; i <= cols; i++)
            {
                writer.Write("<c r=\"{0}{1}\" t=\"inlineStr\"><is><t>{2}</t></is></c>", GetColumnName(i), 1,
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