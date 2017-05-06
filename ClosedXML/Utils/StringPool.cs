using System.Collections.Generic;

namespace ClosedXML.Utils
{
    public class StringPool
    {
        private readonly Dictionary<string, string> _pool = new Dictionary<string, string>();

        public string Get(string entry)
        {
            string result;
            if (!_pool.TryGetValue(entry, out result))
            {
                _pool[entry] = entry;
                result = entry;
            }
            return result;
        }
    }
}