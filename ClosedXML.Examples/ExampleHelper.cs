using System.IO;

namespace ClosedXML.Examples
{
    public static class ExampleHelper
    {
        public static string GetTempFilePath()
        {
            return Path.GetTempFileName();
        }

        public static string GetTempFilePath(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            var tempFilePath = GetTempFilePath();
            return Path.ChangeExtension(tempFilePath, extension);
        }
    }
}
