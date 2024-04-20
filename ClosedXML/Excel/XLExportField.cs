using System.Reflection;

namespace ClosedXML.Excel;

/// <summary>
/// Represents a field to be exported, along with its parent property name if applicable.
/// </summary>
public class XLExportField
{
    public XLExportField(PropertyInfo property, string? parentName = null)
    {
        Property = property;
        ParentName = parentName;
    }

    /// <summary>
    /// Gets or sets the name of the parent property.
    /// </summary>
    /// <remarks>
    /// The ParentName property refers to the name of the parent property if the field is nested within a complex object.
    /// </remarks>
    public string? ParentName { get; set; }

    /// <summary>
    /// Gets or sets the property information of the field.
    /// </summary>
    public PropertyInfo? Property { get; set; }
}
