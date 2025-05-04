namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A variable in the generated code.
/// </summary>
/// <param name="Type">C# type of variable, includes nullability.</param>
/// <param name="Name">Name of variable (unescaped).</param>
internal record Variable(string Type, string Name);
