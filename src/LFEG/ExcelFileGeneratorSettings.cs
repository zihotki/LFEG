using System;
using System.Collections;
using System.IO;
using System.Reflection;
using Ionic.Zlib;
using LFEG.Infrastructure;
using LFEG.Infrastructure.DataProviderVisitors;
using LFEG.Infrastructure.Styling;
using LFEG.Infrastructure.Writers;
using System.Collections.Generic;
using System.Linq;

namespace LFEG
{
    public class ExcelFileGeneratorSettings
    {
        private CompressionLevel _compressionLevel = CompressionLevel.Default;
        private CompressionStrategy _compressionStrategy = CompressionStrategy.Default;
        private IXmlEncoder _encoder = new XmlEncoder();
        private string _resourceName = "LFEG.Template.Output.zip";
        private Assembly _resourceAssembly = typeof (ExcelFileGeneratorSettings).Assembly;
        private Func<IEnumerable<PropertyInfo>, IEnumerable<PropertyInfo>> _propertiesFilter = null;
        private Func<Stream> _templateStreamFactory;
        private int _defaultCapacity = 100;
        private string _numberFormat;
        private string _booleanFormat;
        private string _dateFormat = "dd-mmm-yyyy";
        private bool _throwOnLimit = false;

        public ExcelFileGeneratorSettings()
        {
            _templateStreamFactory = () => _resourceAssembly.GetManifestResourceStream(_resourceName);
        }

        public static ExcelFileGeneratorSettings Create()
        {
            return new ExcelFileGeneratorSettings();
        }

        /// <summary>
        /// Sets the compression level. By default CompressionLevel.Default is used
        /// </summary>
        /// <param name="compressionLevel">Compression level</param> 
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetCompressionLevel(CompressionLevel compressionLevel)
        {
            _compressionLevel = compressionLevel;
            return this;
        }

        /// <summary>
        /// Sets the compression strategy. By default CompressionStrategy.Default is used
        /// </summary>
        /// <param name="compressionStrategy">Compression strategy</param> 
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetCompressionStrategy(CompressionStrategy compressionStrategy)
        {
            _compressionStrategy = compressionStrategy;
            return this;
        }

        /// <summary>
        /// Sets XML encoder. By default LFEG.Infrastructure.XmlEncoder is used which encodes strings using 
        /// System.Security.SecurityElement.Escape method
        /// </summary>
        /// <param name="encoder">Encoder</param>
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetXmlEncoder(IXmlEncoder encoder)
        {
            if (encoder == null)
                throw new ArgumentNullException(nameof(encoder));

            _encoder = encoder;
            return this;
        }

        /// <summary>
        /// Sets excel template. By default LFEG's template is used.
        /// </summary>
        /// <param name="assembly">Assembly which contains the template</param>
        /// <param name="resourceName">Name of the excel template resource</param>
        /// <returns>Settings</returns>
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

        /// <summary>
        /// Sets excel template. By default LFEG's template is used.
        /// </summary>
        /// <param name="templateStreamFactory">Template stream factory</param>
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetTemplate(Func<Stream> templateStreamFactory)
        {
            if (templateStreamFactory == null)
                throw new ArgumentNullException(nameof(templateStreamFactory));

            _templateStreamFactory = templateStreamFactory;
            return this;
        }

        /// <summary>
        /// Sets default capacity of interner. If your data contains a lot of repetitive strings, interner puts them into 
        /// shared strings and that allows to reduce size of a rezult file. Providing correct capacity will allow to increase 
        /// performance by a little bit. By default interning is enabled for enum properties and default capacity is 100.
        /// </summary>
        /// <param name="capacity">The initial number of elements that the interner can contain.</param>
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetSharedStringsInternerDefaultCapacity(int capacity)
        {
            _defaultCapacity = capacity;
            return this;
        }

        /// <summary>
        /// Sets default number format. See https://support.office.com/en-my/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 
        /// for format details.
        /// </summary>
        /// <param name="format">Number format</param>
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetDefaultNumberFormat(string format)
        {
            _numberFormat = format;
            return this;
        }

        /// <summary>
        /// Sets default boolean format. See https://support.office.com/en-my/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 
        /// for format details. Excel doesn't allow to specify format for boolean data but if the format is set then the LFEG
        /// will make boolean cells to be numbers (0 for FALSE, 1 for TRUE) and then formatting can be applied.
        /// </summary>
        /// <param name="format">Boolean format</param>
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetDefaultBooleanFormat(string format)
        {
            _booleanFormat = format;
            return this;
        }

        /// <summary>
        /// Sets default date format. See https://support.office.com/en-my/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4 
        /// for format details.
        /// </summary>
        /// <param name="format"></param>
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetDefaultDateFormat(string format)
        {
            _dateFormat = format;
            return this;
        }

        /// <summary>
        /// Sets additional DTO properties filter. Using this setting you can remove from export some additional properties, 
        /// for example - the ones marked as [ScaffoldColumn(false)]
        /// </summary>
        /// <param name="filter">Properties filter</param>
        /// <returns>Settings</returns>
        public ExcelFileGeneratorSettings SetDtoPropertiesFilter(Func<IEnumerable<PropertyInfo>, IEnumerable<PropertyInfo>> filter)
        {
            _propertiesFilter = filter;
            return this;
        }

        public ExcelFileGeneratorSettings ThrowOnValueLengthExceedsExcelLimit()
        {
            _throwOnLimit = true;
            return this;
        }


        /// <summary>
        /// Creates a new instance of a generator.
        /// </summary>
        /// <returns>Generator</returns>
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
                new DefaultStringDataProviderVisitor(_throwOnLimit)
            };
            
            var columnFactory = new ExcelColumnFactory(visitors, styleProvider);

            return new ExcelFileGenerator(columnFactory,
                new WorksheetWriter(_encoder, sharedStringsInterner),
                new SharedStringsWriter(_encoder, sharedStringsInterner),
                new StylesWriter(styleProvider),
                _templateStreamFactory(),
                _compressionStrategy,
                _compressionLevel,
                _propertiesFilter);
        }
    }
}