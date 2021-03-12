// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Utils
{
    internal static class CryptographicAlgorithms
    {
        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(plainTextBytes);
        }

        public static String GenerateNewSalt(Algorithm algorithm)
        {
            if (RequiresSalt(algorithm))
                return GetSalt();
            else
                return String.Empty;
        }

        public static String GetPasswordHash(Algorithm algorithm, String password, String salt = "", UInt32 spinCount = 0)
        {
            if (password == null)
                throw new ArgumentNullException(nameof(password));

            if (salt == null)
                throw new ArgumentNullException(nameof(salt));

            if ("" == password) return "";

            switch (algorithm)
            {
                case Algorithm.SimpleHash:
                    return GetDefaultPasswordHash(password);

                case Algorithm.SHA512:
                    return GetSha512PasswordHash(password, salt, spinCount);

                default:
                    return string.Empty;
            }
        }

        public static string GetSalt(int length = 32)
        {
            using (var random = new RNGCryptoServiceProvider())
            {
                var salt = new byte[length];
                random.GetNonZeroBytes(salt);
                return Convert.ToBase64String(salt);
            }
        }

        public static Boolean RequiresSalt(Algorithm algorithm)
        {
            switch (algorithm)
            {
                case Algorithm.SimpleHash:
                    return false;

                case Algorithm.SHA512:
                    return true;

                default:
                    return false;
            }
        }

        private static String GetDefaultPasswordHash(String password)
        {
            if (password == null)
                throw new ArgumentNullException(nameof(password));

            // http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
            // http://sc.openoffice.org/excelfileformat.pdf - 4.18.4
            // http://web.archive.org/web/20080906232341/http://blogs.infosupport.com/wouterv/archive/2006/11/21/Hashing-password-for-use-in-SpreadsheetML.aspx
            byte[] passwordCharacters = Encoding.ASCII.GetBytes(password);
            int hash = 0;
            if (passwordCharacters.Length > 0)
            {
                int charIndex = passwordCharacters.Length;

                while (charIndex-- > 0)
                {
                    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                    hash ^= passwordCharacters[charIndex];
                }
                // Main difference from spec, also hash with charcount
                hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7fff);
                hash ^= passwordCharacters.Length;
                hash ^= (0x8000 | ('N' << 8) | 'K');
            }

            return Convert.ToString(hash, 16).ToUpperInvariant();
        }

        private static String GetSha512PasswordHash(String password, String salt, UInt32 spinCount)
        {
            if (password == null)
                throw new ArgumentNullException(nameof(password));

            if (salt == null)
                throw new ArgumentNullException(nameof(salt));

            var saltBytes = Convert.FromBase64String(salt);
            var passwordBytes = Encoding.Unicode.GetBytes(password);
            var bytes = saltBytes.Concat(passwordBytes).ToArray();

            byte[] hashedBytes;
            using (var hash = new SHA512Managed())
            {
                hashedBytes = hash.ComputeHash(bytes);

                bytes = new byte[hashedBytes.Length + sizeof(uint)];
                for (uint i = 0; i < spinCount; i++)
                {
                    var le = BitConverter.GetBytes(i);
                    Array.Copy(hashedBytes, bytes, hashedBytes.Length);
                    Array.Copy(le, 0, bytes, hashedBytes.Length, le.Length);
                    hashedBytes = hash.ComputeHash(bytes);
                }
            }

            return Convert.ToBase64String(hashedBytes);
        }
    }
}
