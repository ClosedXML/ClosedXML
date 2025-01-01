using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using ClosedXML.Excel.IO;

namespace ClosedXML.Sandbox
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            long len = 0;
            var sw = new Stopwatch();
            var commonCrawlDir = @"d:\temp\cc";
            var files = Directory.EnumerateFiles(commonCrawlDir, "*.zip").ToList();//.Where(x => x.Contains("136c58acbbc5fc46031f2d7d370fb6934d44c28019ef8fd7aebd9c0817dc81c9"));
            var count = 0;
            var errorCount = 0;
            var zipErrorCount = 0;
            var styleLessCount = 0;
            foreach (var filePath in files.Take(100000))
            {
                count++;
                if (count % 100 == 0)
                    Console.Write('.');

                if (count % 5000 == 0)
                {
                    Console.WriteLine("\n{0:N} ms", sw.ElapsedMilliseconds);
                    Console.WriteLine("Total len {0:N}", len);
                }

                using var f = File.OpenRead(filePath);
                ZipArchive archive;
                try
                {
                    archive = new ZipArchive(f, ZipArchiveMode.Read);
                }
                catch
                {
                    zipErrorCount++;
                    continue;
                }
                
                ZipArchiveEntry styles;
                try
                {
                    styles = archive.GetEntry("xl/styles.xml");
                    if (styles is null)
                    {
                        styleLessCount++;
                        continue;
                    }
                }
                catch
                {
                    zipErrorCount++;
                    continue;
                }

                using var styleStream = styles.Open();
                len += styles.Length;
                try
                {
                    sw.Start();
                    new StyleSheetReader().Load(styleStream);
                }
                catch (Exception ex)
                {
                    errorCount++;
                    Console.WriteLine($"\n{Path.GetFileName(filePath)} ERROR: " + ex.Message);
                    Console.WriteLine(ex.StackTrace);
                }
                finally
                {
                    sw.Stop();
                }

                archive.Dispose();
                
            }


            Console.WriteLine("Total: {0}", count);
            Console.WriteLine("Errors: {0}", errorCount);
            Console.WriteLine("ZipErrors: {0}", zipErrorCount);
            Console.WriteLine("StyleLessCount: {0}", styleLessCount);
            Console.WriteLine("{0:N} ms", sw.ElapsedMilliseconds);
            Console.WriteLine("Total len {0:N}", len);
            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
