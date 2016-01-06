using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Ionic.Zip;

namespace LFEG.Infrastructure.Writers
{
    public class SharedStringsWriter : ISharedStringsWriter
    {
        private readonly ISharedStringsInterner _interner;
        private readonly IXmlEncoder _xmlEncoder;

        public SharedStringsWriter(IXmlEncoder xmlEncoder, ISharedStringsInterner interner)
        {
            _interner = interner;
            _xmlEncoder = xmlEncoder;
        }

        public void Write(ZipFile zip)
        {
            zip.AddEntry("xl\\sharedStrings.xml", (name, entryStream) =>
            {
                using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
                {
                    WriteHeader(writer, _interner.Cache.Count);
                    WriteContent(writer, _interner.Cache);
                    WriteFooter(writer);
                }
            });
        }

        private void WriteContent(StreamWriter writer, Dictionary<string, int> cache)
        {
            // order of items in dictionary may be different from insert order
            foreach (var item in cache.OrderBy(x => x.Value))
            {
                writer.Write("<si><t>");
                writer.Write(_xmlEncoder.Encode(item.Key));
                writer.Write("</t></si>");
            }
        }

        private void WriteHeader(StreamWriter writer, int count)
        {
            writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""{0}"" uniqueCount=""{0}"">", count);
        }

        private void WriteFooter(StreamWriter writer)
        {
            writer.Write("</sst>");           
        }
    }
}