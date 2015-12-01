using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Ionic.Zlib;
using LFEG.ExcelColumnDataInitializerVisitors;

namespace LFEG
{
    public class ExcelFileGeneratorSettings
    {
        private readonly List<Func<IDataProviderVisitor>> _visitorFactories =
            new List<Func<IDataProviderVisitor>>();

        private CompressionLevel _compressionLevel = CompressionLevel.Default;
        private CompressionStrategy _compressionStrategy = CompressionStrategy.Default;
        private IXmlEncoder _encoder = new XmlEncoder();
        private string _resourceName = "LFEG.Template.Output.zip";
        private Assembly _resourceAssembly = typeof (ExcelFileGeneratorSettings).Assembly;
        private Func<Stream> _templateStreamFactory;

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

        public ExcelFileGeneratorSettings SetResource(Assembly assembly, string resourceName)
        {
            if (assembly == null)
                throw new ArgumentNullException(nameof(assembly));

            if (resourceName == null)
                throw new ArgumentNullException(nameof(resourceName));

            _resourceAssembly = assembly;
            _resourceName = resourceName;
            return this;
        }

        public ExcelFileGeneratorSettings SetTemplateStreamFactory(Func<Stream> templateStreamFactory)
        {
            if (templateStreamFactory == null)
                throw new ArgumentNullException(nameof(templateStreamFactory));

            _templateStreamFactory = templateStreamFactory;
            return this;
        }

        public ExcelFileGeneratorSettings AddColumnVisitor<T>()
            where T : IDataProviderVisitor, new()
        {
            _visitorFactories.Add(() => new T());
            return this;
        }

        public ExcelFileGeneratorSettings AddDefaultColumnVisitors()
        {
            _visitorFactories.Add(() => new EnumDataProviderVisitor());
            _visitorFactories.Add(() => new NumericDataProviderVisitor());
            _visitorFactories.Add(() => new DefaultStringDataProviderVisitor());

            return this;
        }

        public ExcelFileGenerator CreateGenerator()
        {
            return new ExcelFileGenerator(CreateColumnFactory(), CreateWriter(),
                _templateStreamFactory(), _compressionStrategy, _compressionLevel);
        }

        private ExcelWriter CreateWriter()
        {
            return new ExcelWriter(_encoder);
        }

        private ExcelColumnFactory CreateColumnFactory()
        {
            var visitors = _visitorFactories.Select(x => x()).ToArray();
            return new ExcelColumnFactory(visitors);
        }
    }
}