using System.Collections.Generic;

namespace LFEG.Infrastructure
{
    public interface ISharedStringsInterner
    {
        Dictionary<string, int> Cache { get; } 

        int Intern(string value);
    }
}