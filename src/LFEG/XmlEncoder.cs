using System.Security;

namespace LFEG
{
    public class XmlEncoder : IXmlEncoder
    {
        public string Encode(string str)
        {
            return SecurityElement.Escape(str);
        }
    }

    public interface IXmlEncoder
    {
        string Encode(string str);
    }
}