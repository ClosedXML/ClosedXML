// Keep this file CodeMaid organised and cleaned
using System;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel
{
    internal class XLSheetProtection : IXLSheetProtection
    {
        public XLSheetProtection(Algorithm algorithm)
        {
            this.Algorithm = algorithm;
            AllowedElements = XLSheetProtectionElements.SelectEverything;
        }

        public Algorithm Algorithm { get; internal set; }
        public XLSheetProtectionElements AllowedElements { get; set; }

        public Boolean IsPasswordProtected => this.IsProtected && !String.IsNullOrEmpty(PasswordHash);
        public Boolean IsProtected { get; internal set; }


        internal String Base64EncodedSalt { get; set; }
        internal String PasswordHash { get; set; }
        internal UInt32 SpinCount { get; set; } = 100000;

        public IXLSheetProtection AllowElement(XLSheetProtectionElements element, Boolean allowed = true)
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
            return new XLSheetProtection(this.Algorithm)
            {
                IsProtected = this.IsProtected,
                PasswordHash = this.PasswordHash,
                SpinCount = this.SpinCount,
                Base64EncodedSalt = this.Base64EncodedSalt,
                AllowedElements = this.AllowedElements
            };
        }

        public IXLSheetProtection CopyFrom(IXLSheetProtection sheetProtection)
        {
            if (sheetProtection is XLSheetProtection xlSheetProtection)
            {
                this.IsProtected = xlSheetProtection.IsProtected;
                this.Algorithm = xlSheetProtection.Algorithm;
                this.PasswordHash = xlSheetProtection.PasswordHash;
                this.SpinCount = xlSheetProtection.SpinCount;
                this.Base64EncodedSalt = xlSheetProtection.Base64EncodedSalt;
                this.AllowedElements = xlSheetProtection.AllowedElements;
            }
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

        public IXLSheetProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm)
        {
            if (IsProtected)
            {
                throw new InvalidOperationException("The worksheet is already protected");
            }
            else
            {
                IsProtected = true;

                password = password ?? "";

                this.Algorithm = algorithm;
                this.Base64EncodedSalt = Utils.CryptographicAlgorithms.GenerateNewSalt(this.Algorithm);
                this.PasswordHash = Utils.CryptographicAlgorithms.GetPasswordHash(this.Algorithm, password, this.Base64EncodedSalt, this.SpinCount);
            }
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
                password = password ?? "";

                if ("" != PasswordHash && "" == password)
                    throw new InvalidOperationException("The worksheet is password protected");

                var hash = Utils.CryptographicAlgorithms.GetPasswordHash(this.Algorithm, password, this.Base64EncodedSalt, this.SpinCount);
                if (hash != PasswordHash)
                    throw new ArgumentException("Invalid password");
                else
                {
                    IsProtected = false;
                    PasswordHash = String.Empty;
                    this.Base64EncodedSalt = String.Empty;
                }
            }

            return this;
        }
    }
}
