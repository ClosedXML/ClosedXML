// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    internal class XLSheetProtection : IXLSheetProtection
    {
        public XLSheetProtection()
        {
            AllowedElements = XLSheetProtectionElements.SelectEverything;
        }

        public XLSheetProtectionElements AllowedElements { get; set; }
        public Boolean IsProtected { get; set; }
        internal String PasswordHash { get; set; }

        public IXLSheetProtection AllowElement(XLSheetProtectionElements element, bool allowed = true)
        {
            if (!allowed)
                return DisallowElement(element);

            AllowedElements |= element;
            return this;
        }

        public IXLSheetProtection AllowEverything()
        {
            return AllowElement(XLSheetProtectionElements.Everything);
        }

        public IXLSheetProtection AllowNone()
        {
            AllowedElements = XLSheetProtectionElements.None;
            return this;
        }

        public object Clone()
        {
            return new XLSheetProtection()
            {
                IsProtected = this.IsProtected,
                PasswordHash = this.PasswordHash,
                AllowedElements = this.AllowedElements
            };
        }

        public IXLSheetProtection CopyFrom(IXLSheetProtection sheetProtection)
        {
            this.IsProtected = sheetProtection.IsProtected;
            this.PasswordHash = (sheetProtection as XLSheetProtection).PasswordHash;
            this.AllowedElements = sheetProtection.AllowedElements;
            return this;
        }

        public IXLSheetProtection DisallowElement(XLSheetProtectionElements element)
        {
            AllowedElements &= ~element;
            return this;
        }

        public IXLSheetProtection Protect()
        {
            return Protect(String.Empty);
        }

        public IXLSheetProtection Protect(String password)
        {
            if (IsProtected)
            {
                throw new InvalidOperationException("The worksheet is already protected");
            }
            else
            {
                IsProtected = true;
                PasswordHash = GetPasswordHash(password);
            }
            return this;
        }

        public IXLSheetProtection SetPassword(String value)
        {
            PasswordHash = GetPasswordHash(value);
            return this;
        }

        public IXLSheetProtection Unprotect()
        {
            return Unprotect(String.Empty);
        }

        public IXLSheetProtection Unprotect(String password)
        {
            if (IsProtected)
            {
                String hash = GetPasswordHash(password);
                if (hash != PasswordHash)
                    throw new ArgumentException("Invalid password");
                else
                {
                    IsProtected = false;
                    PasswordHash = String.Empty;
                }
            }

            return this;
        }

        private String GetPasswordHash(String password)
        {
            Int32 pLength = password.Length;
            Int32 hash = 0;
            if (pLength == 0) return String.Empty;

            for (Int32 i = pLength - 1; i >= 0; i--)
            {
                hash ^= password[i];
                hash = hash >> 14 & 0x01 | hash << 1 & 0x7fff;
            }
            hash ^= 0x8000 | 'N' << 8 | 'K';
            hash ^= pLength;
            return hash.ToString("X");
        }
    }
}
