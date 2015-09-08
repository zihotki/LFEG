using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Ionic.Zip;

namespace LFEG
{
    public class ExcelFileGenerator
    {
        private readonly ExcelWriter _writer;
        private readonly Stream _templateStream;
        private readonly ExcelColumnFactory _formatter;
        private const string ResourceName = "LFEG.Template.Output.zip";

        public ExcelFileGenerator(ExcelColumnFactory formatter = null, ExcelWriter writer = null, Stream templateStream = null)
        {
            _formatter = formatter ?? new ExcelColumnFactory();
            _writer = writer ?? new ExcelWriter();
            _templateStream = templateStream;
        }

        public void GenerateFile(DataTable dataTable, Stream outputStream)
        {
            var cols = dataTable.Columns.Cast<DataColumn>()
                .Select((column, index) => _formatter.CreateDataTableColumn(column, index))
                .ToArray();

            Generate(dataTable.Rows.Cast<DataRow>(), cols, outputStream);
        }

        public void GenerateFile<T>(IEnumerable<T> items, Stream outputStream)
        {
            var cols = GetColumns(typeof (T));

            Generate(items, cols, outputStream);
        }

        public void GenerateFile(IQueryable items, Stream outputStream)
        {
            var cols = GetColumns(items.ElementType);

            Generate(items, cols, outputStream);
        }

        public void GenerateFile(IDataReader dataReader, Stream outputStream)
        {
            var schemaTable = dataReader.GetSchemaTable();
            if (schemaTable == null)
            {
                throw new InvalidOperationException("Schema Table for data reader is null.");
            }

            var cols = schemaTable.Columns.Cast<DataColumn>()
                .Select((c, i) => _formatter.CreateDataReaderColumn(c, i))
                .ToArray();
            
            Generate(dataReader.Enumerate(), cols.ToArray(), outputStream);
        }

        private void Generate(IEnumerable items, ExcelColumn[] columns, Stream outputStream)
        {
            var assembly = typeof(ExcelFileGenerator).Assembly;

            var stream = _templateStream ?? assembly.GetManifestResourceStream(ResourceName);

            try
            {
                using (var zip = ZipFile.Read(stream))
                {
                    zip.AddEntry("xl\\worksheets\\sheet.xml", (name, entryStream) =>
                    {
                        using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
                        {
                            _writer.WriteXmlHeader(writer);
                            _writer.WriteColumnsDefinitions(writer, columns);
                            _writer.WriteSheetStart(writer);
                            
                            _writer.WriteTitleRow(writer, columns);
                            _writer.WriteRows(writer, items, columns);
                            
                            _writer.WriteSheetEnd(writer);
                            _writer.WriteXmlFooter(writer);
                        }
                    });

                    zip.Save(outputStream);
                }
            }
            catch (Exception e)
            {
                if (_templateStream == null && stream != null)
                {
                    stream.Dispose();
                }
            }
        }

        protected virtual ExcelColumn[] GetColumns(Type type)
        {
            var cols = type.GetProperties()
                .Where(x => x.GetCustomAttribute<IgnoreExcelExportAttribute>(false) == null)
                .Select(x => _formatter.CreatePropertyColumn(x))
                .ToArray();

            return cols;
        }
    }
}