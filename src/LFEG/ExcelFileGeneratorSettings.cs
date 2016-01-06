using System;
using System.IO;
using System.Reflection;
using Ionic.Zlib;
using LFEG.Infrastructure;
using LFEG.Infrastructure.DataProviderVisitors;
using LFEG.Infrastructure.Styling;
using LFEG.Infrastructure.Writers;

namespace LFEG
{
    public class ExcelFileGeneratorSettings
    {
        private CompressionLevel _compressionLevel = CompressionLevel.Default;
        private CompressionStrategy _compressionStrategy = CompressionStrategy.Default;
        private IXmlEncoder _encoder = new XmlEncoder();
        private string _resourceName = "LFEG.Template.Output.zip";
        private Assembly _resourceAssembly = typeof (ExcelFileGeneratorSettings).Assembly;
        private Func<Stream> _templateStreamFactory;
        private int _defaultCapacity = 100;
        private string _numberFormat;
        private string _booleanFormat;
        private string _dateFormat = "dd-mmm-yyyy";

        public ExcelFileGeneratorSettings()
        {
            _templateStreamFactory = () => _resourceAssembly.GetManifestResourceStream(_resourceName);
        }

        public static ExcelFileGeneratorSettings Create()
        {
            return new ExcelFileGeneratorSettings();
        }

        public ExcelFileGeneratorSettings SetCompressionLevel(CompressionLevel compressionLevel)
        {
            _compressionLevel = compressionLevel;
            return this;
        }

        public ExcelFileGeneratorSettings SetCompressionStrategy(CompressionStrategy compressionStrategy)
        {
            _compressionStrategy = compressionStrategy;
            return this;
        }

        public ExcelFileGeneratorSettings SetXmlEncoder(IXmlEncoder encoder)
        {
            if (encoder == null)
                throw new ArgumentNullException(nameof(encoder));

            _encoder = encoder;
            return this;
        }

        public ExcelFileGeneratorSettings SetTemplate(Assembly assembly, string resourceName)
        {
            if (assembly == null)
                throw new ArgumentNullException(nameof(assembly));

            if (resourceName == null)
                throw new ArgumentNullException(nameof(resourceName));

            _resourceAssembly = assembly;
            _resourceName = resourceName;
            return this;
        }

        public ExcelFileGeneratorSettings SetTemplate(Func<Stream> templateStreamFactory)
        {
            if (templateStreamFactory == null)
                throw new ArgumentNullException(nameof(templateStreamFactory));

            _templateStreamFactory = templateStreamFactory;
            return this;
        }

        public ExcelFileGeneratorSettings SetSharedStringsInternerDefaultCapacity(int capacity)
        {
            _defaultCapacity = capacity;
            return this;
        }

        public ExcelFileGeneratorSettings SetDefaultNumberFormat(string format)
        {
            _numberFormat = format;
            return this;
        }

        public ExcelFileGeneratorSettings SetDefaultBooleanFormat(string format)
        {
            _booleanFormat = format;
            return this;
        }

        public ExcelFileGeneratorSettings SetDefaultDateFormat(string format)
        {
            _dateFormat = format;
            return this;
        }

        public ExcelFileGenerator CreateGenerator()
        {
            var sharedStringsInterner = new SharedStringsInterner(_defaultCapacity);
            var styleProvider = new StyleProvider();
            
            var visitors = new IDataProviderVisitor[]
            {
                new EnumDataProviderVisitor(),
                new NumericDataProviderVisitor(string.IsNullOrEmpty(_numberFormat)
                    ? (int?) null
                    : styleProvider.GetColumnStyle(_numberFormat)),
                new DateTimeDataProviderVisitor(string.IsNullOrEmpty(_dateFormat)
                    ? (int?) null
                    : styleProvider.GetColumnStyle(_dateFormat)),
                new BooleanDataProviderVisitor(string.IsNullOrEmpty(_booleanFormat)
                    ? (int?) null
                    : styleProvider.GetColumnStyle(_booleanFormat)),
                new DefaultStringDataProviderVisitor()
            };
            
            var columnFactory = new ExcelColumnFactory(visitors, styleProvider);

            return new ExcelFileGenerator(columnFactory,
                new WorksheetWriter(_encoder, sharedStringsInterner),
                new SharedStringsWriter(_encoder, sharedStringsInterner),
                new StylesWriter(styleProvider),
                _templateStreamFactory(),
                _compressionStrategy,
                _compressionLevel);
        }
    }
}