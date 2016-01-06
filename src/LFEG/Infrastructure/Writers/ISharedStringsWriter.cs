using Ionic.Zip;

namespace LFEG.Infrastructure.Writers
{
    public interface ISharedStringsWriter
    {
        void Write(ZipFile zip);
    }
}