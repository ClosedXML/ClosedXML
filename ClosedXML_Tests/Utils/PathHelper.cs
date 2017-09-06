using System;
using System.IO;
using System.Linq;

namespace ClosedXML_Tests.Utils
{
    public static class PathHelper
    {
        public static string Combine(params string[] paths)
        {
#if _NET35_
            if (paths == null)
            {
                throw new ArgumentNullException("paths");
            }
            return paths.Aggregate(Path.Combine);
#else
            return Path.Combine(paths);
#endif
        }
    }
}
