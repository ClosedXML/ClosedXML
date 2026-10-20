namespace ClosedXML.Excel.Exceptions;

/// <summary>
/// Some limit (generally set in <see cref="LoadOptions"/>) was exceeded.
/// </summary>
internal class LimitExceededException(string message) : ClosedXMLException(message);
