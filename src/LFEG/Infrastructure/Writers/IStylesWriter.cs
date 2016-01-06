using Ionic.Zip;

namespace LFEG.Infrastructure.Writers
{
    public interface IStylesWriter
    {
        void Write(ZipFile zip);
    }
}