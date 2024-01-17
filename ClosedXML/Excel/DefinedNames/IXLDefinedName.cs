using System;

namespace ClosedXML.Excel
{
    /// <summary>
    /// A scope of <see cref="IXLDefinedName"/>. It determines where can be defined name resolved.
    /// </summary>
    public enum XLNamedRangeScope
    {
        /// <summary>
        /// Name is defined at the sheet level and is available only at the sheet
        /// it is defined or <see cref="IXLWorksheet.DefinedNames"/> collection or when referred
        /// with sheet specifier (e.g. <c>Sheet5!Name</c> when name is scoped to <c>Sheet5</c>).
        /// </summary>
        Worksheet,

        /// <summary>
        /// Name is defined at the workbook and is available everywhere.
        /// </summary>
        Workbook
    }

    public interface IXLDefinedName
    {
        /// <summary>
        /// Gets or sets the comment for this named range.
        /// </summary>
        /// <value>
        /// The comment for this named range.
        /// </value>
        String? Comment { get; set; }

        /// <summary>
        /// Checks if the named range contains invalid references (#REF!).
        /// <para>
        /// <example>Defined name with a formula <c>SUM(#REF!A1, Sheet7!B4)</c> would return
        /// <c>true</c>, because <c>#REF!A1</c> is an invalid reference.</example>
        /// </para>
        /// </summary>
        bool IsValid { get; }

        /// <summary>
        /// Gets or sets the name of the range.
        /// </summary>
        /// <value>
        /// The name of the range.
        /// </value>
        /// <exception cref="ArgumentException">Set value is not a valid name.</exception>
        /// <exception cref="InvalidOperationException">The name is colliding with a different name
        /// that is already defined in the collection.</exception>
        String Name { get; set; }

        /// <summary>
        /// Gets the ranges associated with this named range.
        /// <para>Note: A named range can point to multiple ranges.</para>
        /// </summary>
        IXLRanges Ranges { get; }

        /// <summary>
        /// A formula of the named range. In most cases, name is just a range (e.g.
        /// <c>Sheet5!$A$4</c>), but it can be a constant, lambda or other values.
        /// The name formula can contain a bang reference (e.g. reference without
        /// a sheet, but with exclamation mark <c>!$A$5</c>), but can't contain plain
        /// local cell references (i.e. references without a sheet like <c>A5</c>).
        /// </summary>
        String RefersTo { get; set; }

        /// <summary>
        /// Gets the scope of this named range.
        /// </summary>
        XLNamedRangeScope Scope { get; }

        /// <summary>
        /// Gets or sets the visibility of this named range.
        /// </summary>
        /// <value>
        ///   <c>true</c> if visible; otherwise, <c>false</c>.
        /// </value>
        Boolean Visible { get; set; }

        /// <summary>
        /// Copy sheet-scoped defined name to a different sheet. The references to the original
        /// sheet are changed to refer to the <paramref name="targetSheet"/>:
        /// <list type="bullet">
        ///   <item>Cell ranges (<c>Org!A1</c> will be <c>New!A1</c>).</item>
        ///   <item>Tables - if the target sheet contains a table of same size at same place as the original sheet.</item>
        ///   <item>Sheet-specified names (<c>Org!Name</c> will be <c>New!Name</c>, but the actual name won't be created).</item>
        /// </list>
        /// </summary>
        /// <param name="targetSheet">Target sheet where to copy the defined name.</param>
        /// <exception cref="InvalidOperationException">Defined name is workbook-scoped</exception>
        /// <exception cref="InvalidOperationException">Trying to copy defined name to the same sheet.</exception>
        IXLDefinedName CopyTo(IXLWorksheet targetSheet);

        /// <summary>
        /// Deletes this named range (not the cells).
        /// </summary>
        void Delete();

        IXLDefinedName SetRefersTo(String formula);

        IXLDefinedName SetRefersTo(IXLRangeBase range);

        IXLDefinedName SetRefersTo(IXLRanges ranges);
    }
}
