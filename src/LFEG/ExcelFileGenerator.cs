using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Ionic.Zip;
using Ionic.Zlib;

namespace LFEG
{
    public class ExcelFileGenerator
    {
        private readonly ExcelWriter _writer;
        private readonly Stream _templateStream;
        private readonly CompressionStrategy _compressionStrategy;
        private readonly CompressionLevel _compressionLevel;
        private readonly ExcelColumnFactory _factory;

        public ExcelFileGenerator(ExcelColumnFactory factory, ExcelWriter writer, Stream templateStream, 
            CompressionStrategy compressionStrategy, CompressionLevel compressionLevel)
        {
            _factory = factory;
            _writer = writer;
            _templateStream = templateStream;
            _compressionStrategy = compressionStrategy;
            _compressionLevel = compressionLevel;
        }

        
        public void GenerateFile(DataTable dataTable, Stream outputStream)
        {
            var cols = dataTable.Columns.Cast<DataColumn>()
                .Select((column, index) => _factory.CreateDataTableColumn(column, index))
                .ToArray();

            Generate(dataTable.Rows.Cast<DataRow>(), cols, outputStream);
        }

        public void GenerateFile(IDataReader dataReader, Stream outputStream)
        {
            var schemaTable = dataReader.GetSchemaTable();
            if (schemaTable == null)
            {
                throw new InvalidOperationException("Schema Table for data reader is null.");
            }

            var cols = schemaTable.Columns.Cast<DataColumn>()
                .Select((c, i) => _factory.CreateDataReaderColumn(c, i))
                .ToArray();

            Generate(dataReader.Enumerate(), cols.ToArray(), outputStream);
        }

        public void GenerateFile<T>(IEnumerable<T> items, Stream outputStream)
        {
            var cols = GetColumns<T>();

            Generate(items, cols, outputStream);
        }

        public void GenerateFile<T>(IQueryable<T> items, Stream outputStream)
        {
            var cols = GetColumns<T>();

            Generate(items, cols, outputStream);
        }

        private void Generate<T>(IEnumerable<T> items, ExcelColumn<T>[] columns, Stream outputStream)
        {
            var stream = _templateStream;

            try
            {
                using (var zip = ZipFile.Read(stream))
                {
                    zip.CompressionLevel = _compressionLevel;
                    zip.Strategy = _compressionStrategy;

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
            catch (Exception)
            {
                if (_templateStream == null)
                {
                    stream?.Dispose();
                }
            }
        }

        protected virtual ExcelColumn<T>[] GetColumns<T>()
        {
            var cols = typeof(T).GetProperties()
                .Where(x => Attribute.GetCustomAttribute(x, typeof(IgnoreExcelExportAttribute), false) == null)
                .Select(x => _factory.CreatePropertyColumn<T>(x))
                .ToArray();

            return cols;
        }
    }
}