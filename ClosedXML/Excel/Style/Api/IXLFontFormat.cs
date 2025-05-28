namespace ClosedXML.Excel;

/// <summary>
/// A font formatting object. Format doesn't have to have all properties specified. Setting
/// a property to <c>null</c> clears the property.
/// </summary>
/// <typeparam name="TParent">A type of parent object for the fluent API.</typeparam>
public interface IXLFontFormat<out TParent>
    where TParent : class
{
    /// <summary>
    /// Gets or sets whether the formatting should have a bold font.
    /// </summary>
    bool? Bold { get; set; }

    /// <summary>
    /// A fluent way to set the <see cref="Bold"/> property to <c>true</c>.
    /// </summary>
    /// <returns>Parent object of the format.</returns>
    TParent SetBold();

    /// <summary>
    /// A fluent way to set the <see cref="Bold"/> property.
    /// </summary>
    /// <returns>Parent object of the format.</returns>
    TParent SetBold(bool? value);
}
