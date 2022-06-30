// Keep this file CodeMaid organised and cleaned
using System;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel
{
    internal class XLWorkbookProtection : IXLWorkbookProtection
    {
        public XLWorkbookProtection(Algorithm algorithm)
            : this(algorithm, XLWorkbookProtectionElements.Windows)
        {
        }

        public XLWorkbookProtection(Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
        {
            this.Algorithm = algorithm;
            this.AllowedElements = allowedElements;
        }

        public Algorithm Algorithm { get; internal set; }
        public XLWorkbookProtectionElements AllowedElements { get; set; }
        public bool IsPasswordProtected => this.IsProtected && !String.IsNullOrEmpty(PasswordHash);
        public bool IsProtected { get; internal set; }

        internal String Base64EncodedSalt { get; set; }
        internal String PasswordHash { get; set; }
        internal UInt32 SpinCount { get; set; } = 100000;

        public IXLWorkbookProtection AllowElement(XLWorkbookProtectionElements element, Boolean allowed = true)
        {
            if (allowed)
                AllowedElements |= element;
            else
                AllowedElements &= ~element;

            return this;
        }

        public IXLWorkbookProtection AllowEverything()
        {
            AllowedElements = XLWorkbookProtectionElements.Everything;
            return this;
        }

        public IXLWorkbookProtection AllowNone()
        {
            AllowedElements = XLWorkbookProtectionElements.None;
            return this;
        }

        public object Clone()
        {
            return new XLWorkbookProtection(this.Algorithm, this.AllowedElements)
            {
                IsProtected = this.IsProtected,
                PasswordHash = this.PasswordHash,
                SpinCount = this.SpinCount,
                Base64EncodedSalt = this.Base64EncodedSalt
            };
        }

        public IXLWorkbookProtection CopyFrom(IXLElementProtection<XLWorkbookProtectionElements> workbookProtection)
        {
            if (workbookProtection is XLWorkbookProtection xlWorkbookProtection)
            {
                this.IsProtected = xlWorkbookProtection.IsProtected;
                this.Algorithm = xlWorkbookProtection.Algorithm;
                this.PasswordHash = xlWorkbookProtection.PasswordHash;
                this.SpinCount = xlWorkbookProtection.SpinCount;
                this.Base64EncodedSalt = xlWorkbookProtection.Base64EncodedSalt;
                this.AllowedElements = xlWorkbookProtection.AllowedElements;
            }
            return this;
        }

        public IXLWorkbookProtection DisallowElement(XLWorkbookProtectionElements element)
        {
            return AllowElement(element, allowed: false);
        }

        public IXLWorkbookProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
        {
            return Protect(String.Empty, algorithm);
        }

        public IXLWorkbookProtection Protect(XLWorkbookProtectionElements allowedElements)
            => Protect(string.Empty, DefaultProtectionAlgorithm, allowedElements);

        public IXLWorkbookProtection Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements)
            => Protect(string.Empty, algorithm, allowedElements);

        public IXLWorkbookProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm, XLWorkbookProtectionElements allowedElements = XLWorkbookProtectionElements.Windows)
        {
            if (IsProtected)
            {
                throw new InvalidOperationException("The workbook structure is already protected");
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

        public IXLWorkbookProtection Unprotect()
        {
            return Unprotect(String.Empty);
        }

        public IXLWorkbookProtection Unprotect(String password)
        {
            if (IsProtected)
            {
                password = password ?? "";

                if ("" != PasswordHash && "" == password)
                    throw new InvalidOperationException("The workbook structure is password protected");

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

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.AllowElement(XLWorkbookProtectionElements element, Boolean allowed) => AllowElement(element, allowed);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.AllowEverything() => AllowEverything();

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.AllowNone() => AllowNone();

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.DisallowElement(XLWorkbookProtectionElements element) => DisallowElement(element);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Protect(Algorithm algorithm) => Protect(algorithm);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Protect(String password, Algorithm algorithm) => Protect(password, algorithm);

        IXLWorkbookProtection IXLWorkbookProtection.Protect(XLWorkbookProtectionElements allowedElements) => Protect(allowedElements);

        IXLWorkbookProtection IXLWorkbookProtection.Protect(Algorithm algorithm, XLWorkbookProtectionElements allowedElements) => Protect(algorithm, allowedElements);

        IXLWorkbookProtection IXLWorkbookProtection.Protect(String password, Algorithm algorithm, XLWorkbookProtectionElements allowedElements) => Protect(password, algorithm, allowedElements);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Unprotect() => Unprotect();

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Unprotect(String password) => Unprotect(password);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.CopyFrom(IXLElementProtection<XLWorkbookProtectionElements> protectable) => CopyFrom(protectable);

        #endregion IXLProtectable interface
    }
}
