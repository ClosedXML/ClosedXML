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
            Algorithm = algorithm;
            AllowedElements = allowedElements;
        }

        public Algorithm Algorithm { get; internal set; }
        public XLWorkbookProtectionElements AllowedElements { get; set; }
        public bool IsPasswordProtected => IsProtected && !string.IsNullOrEmpty(PasswordHash);
        public bool IsProtected { get; internal set; }

        internal string Base64EncodedSalt { get; set; }
        internal string PasswordHash { get; set; }
        internal uint SpinCount { get; set; } = 100000;

        public IXLWorkbookProtection AllowElement(XLWorkbookProtectionElements element, bool allowed = true)
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
            return new XLWorkbookProtection(Algorithm, AllowedElements)
            {
                IsProtected = IsProtected,
                PasswordHash = PasswordHash,
                SpinCount = SpinCount,
                Base64EncodedSalt = Base64EncodedSalt
            };
        }

        public IXLWorkbookProtection CopyFrom(IXLElementProtection<XLWorkbookProtectionElements> workbookProtection)
        {
            if (workbookProtection is XLWorkbookProtection xlWorkbookProtection)
            {
                IsProtected = xlWorkbookProtection.IsProtected;
                Algorithm = xlWorkbookProtection.Algorithm;
                PasswordHash = xlWorkbookProtection.PasswordHash;
                SpinCount = xlWorkbookProtection.SpinCount;
                Base64EncodedSalt = xlWorkbookProtection.Base64EncodedSalt;
                AllowedElements = xlWorkbookProtection.AllowedElements;
            }
            return this;
        }

        public IXLWorkbookProtection DisallowElement(XLWorkbookProtectionElements element)
        {
            return AllowElement(element, allowed: false);
        }

        public IXLWorkbookProtection Protect()
        {
            return Protect(string.Empty);
        }

        public IXLWorkbookProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm, XLWorkbookProtectionElements allowedElements = XLWorkbookProtectionElements.Windows)
        {
            if (IsProtected)
            {
                throw new InvalidOperationException("The workbook structure is already protected");
            }
            else
            {
                IsProtected = true;

                password = password ?? "";

                Algorithm = algorithm;
                Base64EncodedSalt = Utils.CryptographicAlgorithms.GenerateNewSalt(Algorithm);
                PasswordHash = Utils.CryptographicAlgorithms.GetPasswordHash(Algorithm, password, Base64EncodedSalt, SpinCount);
            }

            AllowedElements = allowedElements;

            return this;
        }

        public IXLWorkbookProtection Unprotect()
        {
            return Unprotect(string.Empty);
        }

        public IXLWorkbookProtection Unprotect(string password)
        {
            if (IsProtected)
            {
                password = password ?? "";

                if (!string.IsNullOrEmpty(PasswordHash) && string.IsNullOrEmpty(password))
                    throw new InvalidOperationException("The workbook structure is password protected");

                var hash = Utils.CryptographicAlgorithms.GetPasswordHash(Algorithm, password, Base64EncodedSalt, SpinCount);
                if (hash != PasswordHash)
                    throw new ArgumentException("Invalid password");
                else
                {
                    IsProtected = false;
                    PasswordHash = string.Empty;
                    Base64EncodedSalt = string.Empty;
                }
            }

            return this;
        }

        #region IXLProtectable interface

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.AllowElement(XLWorkbookProtectionElements element, bool allowed) => AllowElement(element, allowed);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.AllowEverything() => AllowEverything();

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.AllowNone() => AllowNone();

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.DisallowElement(XLWorkbookProtectionElements element) => DisallowElement(element);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Protect() => Protect();

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Protect(string password, Algorithm algorithm) => Protect(password, algorithm);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Unprotect() => Unprotect();

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.Unprotect(string password) => Unprotect(password);

        IXLElementProtection<XLWorkbookProtectionElements> IXLElementProtection<XLWorkbookProtectionElements>.CopyFrom(IXLElementProtection<XLWorkbookProtectionElements> protectable) => CopyFrom(protectable);

        #endregion IXLProtectable interface
    }
}
