using System.IO;
using System.Text;
using Ionic.Zip;
using LFEG.Infrastructure.Styling;

namespace LFEG.Infrastructure.Writers
{
    public class StylesWriter : IStylesWriter
    {
        private readonly IStyleProvider _styleProvider;

        public StylesWriter(IStyleProvider styleProvider)
        {
            _styleProvider = styleProvider;
        }

        public void Write(ZipFile zip)
        {
            zip.AddEntry("xl\\styles.xml", (name, entryStream) =>
            {
                using (var writer = new StreamWriter(entryStream, Encoding.UTF8))
                {
                    WriteHeader(writer);

                    WriteNumberFormats(writer);

                    WriteDefaultFont(writer);
                    WriteDefaultFill(writer);
                    WriteDefaultBorder(writer);
                    WriteDefaultCellStyleFormattingRecords(writer);

                    WriteCellFormats(writer);

                    WriteDefaultCellStyle(writer);
                    //WriteTableStyles(writer);
                    
                    WriteFooter(writer);
                }
            });
        }

        private void WriteCellFormats(StreamWriter writer)
        {
            writer.Write(@"<cellXfs count=""{0}"">", _styleProvider.CellFormats.Count);

            foreach (var cellXf in _styleProvider.CellFormats)
            {
                writer.Write(@"<xf numFmtId=""{0}"" fontId=""0"" fillId=""0"" borderId=""0"" xfId=""0"" applyAlignment=""1"" applyFont=""1"" applyNumberFormat=""1""/>",
                    cellXf.FormatId);
            }

            writer.Write("</cellXfs>");
        }

        private void WriteNumberFormats(StreamWriter writer)
        {
            writer.Write(@"<numFmts count=""{0}"">", _styleProvider.NumberFormats.Count);

            foreach (var numberFormat in _styleProvider.NumberFormats)
            {
                writer.Write(@"<numFmt numFmtId=""{0}"" formatCode=""{1}""/>", 
                    numberFormat.Id, 
                    numberFormat.FormatCode.Replace("\"", "&quot;"));
            }

            writer.Write("</numFmts>");
        }

        /*private void WriteTableStyles(StreamWriter writer)
        {
            writer.Write(@"<tableStyles count=""0"" defaultTableStyle=""{0}"" />", _styleProvider.TableStyleName);
        }*/

        
        #region Default stuff

        private void WriteDefaultBorder(StreamWriter writer)
        {
            writer.Write(@"<borders count=""1""><border><left /><right /><top /><bottom /></border></borders>");
        }

        private void WriteDefaultFill(StreamWriter writer)
        {
            // default fill
            writer.Write(@"<fills count=""1""><fill><patternFill patternType=""none"" /></fill></fills>");
        }

        private void WriteDefaultFont(StreamWriter writer)
        {
            // default font
            writer.Write(@"<fonts count=""1""><font /></fonts>");
        }

        private void WriteDefaultCellStyle(StreamWriter writer)
        {
            writer.Write(@"<cellStyles count=""1""><cellStyle xfId=""0"" name=""Normal"" builtinId=""0"" /></cellStyles>");
        }

        private void WriteDefaultCellStyleFormattingRecords(StreamWriter writer)
        {
            writer.Write(@"<cellStyleXfs count=""1"">");
            writer.Write(@"<xf borderId=""0"" fillId=""0"" fontId=""0"" numFmtId=""0"" applyAlignment=""1"" applyFont=""1"" /> ");
            writer.Write("</cellStyleXfs>");
        }

        #endregion

        private void WriteFooter(StreamWriter writer)
        {
            writer.Write("</styleSheet>");
        }

        private void WriteHeader(StreamWriter writer)
        {
            writer.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"">");
        }
    }
}