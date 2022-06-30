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

        public XLSheetProtection CopyFrom(IXLElementProtection<XLSheetProtectionElements> sheetProtection)
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

        public IXLSheetProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
        {
            return Protect(String.Empty, algorithm);
        }

        public IXLSheetProtection Protect(XLSheetProtectionElements allowedElements)
            => Protect(string.Empty, DefaultProtectionAlgorithm, allowedElements);

        public IXLSheetProtection Protect(Algorithm algorithm, XLSheetProtectionElements allowedElements)
            => Protect(string.Empty, algorithm, allowedElements);

        public IXLSheetProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm, XLSheetProtectionElements allowedElements = XLSheetProtectionElements.SelectEverything)
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

            this.AllowedElements = allowedElements;

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

        #region IXLProtectable interface

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowElement(XLSheetProtectionElements element, Boolean allowed) => AllowElement(element, allowed);

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowEverything() => AllowEverything();

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowNone() => AllowNone();

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.CopyFrom(IXLElementProtection<XLSheetProtectionElements> protectable) => CopyFrom(protectable);

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.DisallowElement(XLSheetProtectionElements element) => DisallowElement(element);

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Protect(Algorithm algorithm) => Protect(algorithm);

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Protect(String password, Algorithm algorithm) => Protect(password, algorithm);

        IXLSheetProtection IXLSheetProtection.Protect(XLSheetProtectionElements allowedElements) => Protect(allowedElements);

        IXLSheetProtection IXLSheetProtection.Protect(Algorithm algorithm, XLSheetProtectionElements allowedElements) => Protect(algorithm, allowedElements);

        IXLSheetProtection IXLSheetProtection.Protect(String password, Algorithm algorithm, XLSheetProtectionElements allowedElements) => Protect(password, algorithm, allowedElements);

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Unprotect() => Unprotect();

        IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Unprotect(String password) => Unprotect(password);

        #endregion IXLProtectable interface
    }
}
