using System.Collections.Generic;

namespace LFEG.Infrastructure
{
    public class SharedStringsInterner : ISharedStringsInterner
    {
        public Dictionary<string, int> Cache { get; }

        public SharedStringsInterner(int capacity)
        {
            Cache = new Dictionary<string, int>(capacity);
        }

        public int Intern(string value)
        {
            int index;
            if (Cache.TryGetValue(value, out index))
            {
                return index;
            }

            index = Cache.Count;
            Cache.Add(value, Cache.Count);

            return index;
        }
    }
}