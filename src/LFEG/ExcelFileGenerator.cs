using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Ionic.Zip;
using Ionic.Zlib;
using LFEG.Infrastructure;
using LFEG.Infrastructure.Writers;

namespace LFEG
{
    public class ExcelFileGenerator
    {
        private readonly IWorksheetWriter _worksheetWriter;
        private readonly ISharedStringsWriter _sharedStringsWriter;
        private readonly IStylesWriter _stylesWriter;
        private readonly Stream _templateStream;
        private readonly CompressionStrategy _compressionStrategy;
        private readonly CompressionLevel _compressionLevel;
        private readonly ExcelColumnFactory _factory;

        public ExcelFileGenerator(ExcelColumnFactory factory, 
            IWorksheetWriter worksheetWriter,
            ISharedStringsWriter sharedStringsWriter,
            IStylesWriter stylesWriter,
            Stream templateStream, 
            CompressionStrategy compressionStrategy, 
            CompressionLevel compressionLevel)
        {
            _factory = factory;
            _worksheetWriter = worksheetWriter;
            _sharedStringsWriter = sharedStringsWriter;
            _stylesWriter = stylesWriter;
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

                    _worksheetWriter.Write(zip, items, columns);
                    _sharedStringsWriter.Write(zip);
                    _stylesWriter.Write(zip);

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