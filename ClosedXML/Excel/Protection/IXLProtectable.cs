#nullable disable

// Keep this file CodeMaid organised and cleaned
using System;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Excel
{
    public interface IXLProtectable<TProtection, in TElement> : IXLProtectable
        where TProtection : IXLElementProtection<TElement>
        where TElement : struct
    {
        TProtection Protection { get; set; }

        /// <summary>Protects this instance without a password.</summary>
        TProtection Protect(TElement allowedElements);

        /// <summary>Protects this instance without a password.</summary>
        new TProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm);

        /// <summary>Protects this instance with the specified password, password hash algorithm and set elements that the user is allowed to change.</summary>
        /// <param name="algorithm">The algorithm.</param>
        /// <param name="allowedElements">The allowed elements.</param>
        TProtection Protect(Algorithm algorithm, TElement allowedElements);

        /// <summary>Protects this instance using the specified password and password hash algorithm.</summary>
        /// <param name="password">The password.</param>
        /// <param name="algorithm">The algorithm.</param>
        new TProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm);

        /// <summary>Protects this instance with the specified password, password hash algorithm and set elements that the user is allowed to change.</summary>
        /// <param name="password">The password.</param>
        /// <param name="algorithm">The algorithm.</param>
        /// <param name="allowedElements">The allowed elements.</param>
        TProtection Protect(String password, Algorithm algorithm, TElement allowedElements);

        /// <summary>Unprotects this instance without a password.</summary>
        new TProtection Unprotect();

        /// <summary>Unprotects this instance using the specified password.</summary>
        /// <param name="password">The password.</param>
        new TProtection Unprotect(String password);
    }

    public interface IXLProtectable
    {
        /// <summary>Gets a value indicating whether this instance is protected with a password.</summary>
        /// <value>
        ///   <c>true</c> if this instance is password protected; otherwise, <c>false</c>.
        /// </value>
        Boolean IsPasswordProtected { get; }

        /// <summary>Gets a value indicating whether this instance is protected, either with or without a password.</summary>
        /// <value>
        ///   <c>true</c> if this instance is protected; otherwise, <c>false</c>.
        /// </value>
        Boolean IsProtected { get; }

        /// <summary>Protects this instance without a password.</summary>
        IXLElementProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm);

        /// <summary>Protects this instance using the specified password and password hash algorithm.</summary>
        /// <param name="password">The password.</param>
        /// <param name="algorithm">The algorithm.</param>
        IXLElementProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm);

        /// <summary>Unprotects this instance without a password.</summary>
        IXLElementProtection Unprotect();

        /// <summary>Unprotects this instance using the specified password.</summary>
        /// <param name="password">The password.</param>
        IXLElementProtection Unprotect(String password);
    }
}
