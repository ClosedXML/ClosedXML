using System;
using System.IO;

namespace ClosedXML.Examples
{
    public class Program
    {
        public static string BaseCreatedDirectory
        {
            get
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Created");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        }

        public static string BaseModifiedDirectory
        {
            get
            {
                var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Modified");
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                return path;
            }
        }

        private static void Main(string[] args)
        {
            CreateFiles.CreateAllFiles();
            LoadFiles.LoadAllFiles();
        }
    }
}