using System.IO;

namespace ClosedXML_Examples
{
    public static class ExampleHelper
    {
        public static string GetTempFilePath()
        {
            return Path.GetTempFileName();
        }
    }
}