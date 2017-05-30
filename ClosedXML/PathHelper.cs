using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ClosedXML
{
    /// <summary>
    ///   Help methods for work with files
    /// </summary>
    internal static class PathHelper
    {
        private static readonly Regex ms_checkValidAbsolutePathRegEx = new Regex(
                "^(([a-zA-Z]\\:)|\\\\)(\\\\|(\\\\[^\\n\\r\\t/:*?<>\\\"|]*)+)$",
                RegexOptions.IgnoreCase
                | RegexOptions.Singleline
                | RegexOptions.CultureInvariant
                | RegexOptions.Compiled
                );
        /// <summary>
        ///   Can check only .\dfdfd\dfdf\dfdf or ..\..\gfhfgh\fghfh
        /// </summary>
        private static readonly Regex ms_checkValidRelativePathRegEx = new Regex(
                "^(\\.\\.|\\.)(\\\\|(\\\\[^\\n\\r\\t/:*?<>\\\"|]*)+)$",
                RegexOptions.IgnoreCase
                | RegexOptions.Multiline
                | RegexOptions.CultureInvariant
                | RegexOptions.Compiled
                );

        private static readonly object ms_syncAbsolutePathObj = new object();

        /// <summary>
        ///   Gets data and time string stamp for file name
        /// </summary>
        /// <returns></returns>
        public static string GetTimeStamp()
        {
            return GetTimeStamp(DateTime.Now);
        }
        /// <summary>
        ///   Gets data and time string stamp for file name
        /// </summary>
        /// <returns></returns>
        public static string GetTimeStamp(DateTime dateTime)
        {
            return dateTime.ToString("ddMMMyyyy_HHmmss", DateTimeFormatInfo.InvariantInfo);
        }
        /// <summary>
        ///   Safety delete file(with try block)
        /// </summary>
        /// <param name = "fileName">file name</param>
        public static void SafetyDeleteFile(string fileName)
        {
            try
            {
                if (!string.IsNullOrEmpty(fileName) && File.Exists(fileName))
                {
                    File.Delete(fileName);
                }
            }
            catch (Exception ex)
            {
                Debug.Fail("Error during file deleting. Error message:" + ex.Message);
            }
        }
        /// <summary>
        ///   Replace all not implemented symbols to '_'
        /// </summary>
        /// <param name = "fileName">input filename</param>
        /// <returns></returns>
        public static string NormalizeFileName(string fileName)
        {
            char[] invalidCharacters = Path.GetInvalidFileNameChars();
            #region Replace begin space
            string result = fileName.TrimStart(' ');
            if (result.Length < fileName.Length)
            {
                result = new string('_', fileName.Length - result.Length) + result;
            }
            #endregion
            foreach (char curChar in invalidCharacters)
            {
                result = result.Replace(curChar, '_');
            }
            return result;
        }

        public static string NormalizePathName(string fileName)
        {
            char[] invalidCharacters = Path.GetInvalidPathChars();
            string result = fileName.TrimStart(' ');
            foreach (char curChar in invalidCharacters)
            {
                result = result.Replace(curChar, '_');
            }
            return result;
        }
        /// <summary>
        ///   ValidatePath file or diretory path
        /// </summary>
        /// <param name = "path"></param>
        /// <param name = "type">path type</param>
        /// <returns></returns>
        public static bool ValidatePath(string path, PathTypes type)
        {
            bool _result = false;
            if ((type & PathTypes.Absolute) == PathTypes.Absolute)
            {
                _result |= ms_checkValidAbsolutePathRegEx.IsMatch(path);
            }
            if ((type & PathTypes.Relative) == PathTypes.Relative)
            {
                _result |= ms_checkValidRelativePathRegEx.IsMatch(path);
            }
            return _result;
        }

        /// <summary>
        ///   ValidatePath file or diretory path
        /// </summary>
        /// <param name = "fileName"></param>
        /// <returns></returns>
        public static bool ValidateFileName(string fileName)
        {
            return fileName.LastIndexOfAny(Path.GetInvalidFileNameChars()) < 0;
        }

        public static string EvaluateRelativePath(string mainDirPath, string absoluteFilePath)
        {
            string[] firstPathParts = mainDirPath.Trim(Path.DirectorySeparatorChar).Split(Path.DirectorySeparatorChar);
            string[] secondPathParts = absoluteFilePath.Trim(Path.DirectorySeparatorChar).Split(Path.DirectorySeparatorChar);

            int sameCounter = 0;
            int partsCount = Math.Min(firstPathParts.Length, secondPathParts.Length);
            for (int i = 0; i < partsCount; i++)
            {
                if (String.Compare(firstPathParts[i], secondPathParts[i], true) != 0)
                {
                    break;
                }
                sameCounter++;
            }

            if (sameCounter == 0)
            {
                return absoluteFilePath;
            }

            string newPath = string.Empty;
            for (int i = sameCounter; i < firstPathParts.Length; i++)
            {
                if (i > sameCounter)
                {
                    newPath += Path.DirectorySeparatorChar;
                }
                newPath += "..";
            }
            if (newPath.Length == 0)
            {
                newPath = ".";
            }
            for (int i = sameCounter; i < secondPathParts.Length; i++)
            {
                newPath += Path.DirectorySeparatorChar;
                newPath += secondPathParts[i];
            }
            return newPath;
        }

        public static string EvaluateAbsolutePath(string rootPath, string relativePath)
        {
            lock (ms_syncAbsolutePathObj)
            {
                string _temp = Environment.CurrentDirectory;
                Environment.CurrentDirectory = rootPath;
                string _result = Path.GetFullPath(relativePath);
                Environment.CurrentDirectory = _temp;
                return _result;
            }
        }

        public static bool TryCreateFile(string filePath, out string message)
        {
            #region Check
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentNullException("filePath");
            }
            #endregion
            try
            {
                if (File.Exists(filePath))
                {
                    message = null;
                    return true;
                }

                using (FileStream _stream = new FileStream(filePath,
                                                           FileMode.Create,
                                                           FileAccess.ReadWrite,
                                                           FileShare.Read,
                                                           1024,
                                                           FileOptions.DeleteOnClose))
                {
                }
            }
            catch (Exception exc)
            {
                message = exc.Message;
                return false;
            }

            message = null;
            return true;
        }

        public static string CreateDirectory(string path)
        {
            try
            {
                var info = new DirectoryInfo(path);
                if (!info.Exists)
                {
                    info.Create();
                }
            }
            catch
            {
            }
            return path;
        }

        [Flags]
        public enum PathTypes
        {
            None = 0,
            Absolute = 1,
            Relative = 2
        }
    }
}