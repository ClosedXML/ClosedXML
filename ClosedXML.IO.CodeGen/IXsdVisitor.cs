using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.Model.Elements;
using ClosedXML.IO.CodeGen.Model.SimpleTypes;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A visitor for visiting <see cref="INode">nodes</see> and running a code for each type of a node.
/// </summary>
/// <typeparam name="TResult">Type of returned value.</typeparam>
public interface IXsdVisitor<out TResult>
{
    TResult Visit(Schema schema);

    TResult Visit(SimpleType simpleType);

    TResult Visit(SimpleTypeList simpleType);

    TResult Visit(SimpleTypeUnion simpleType);

    TResult Visit(ElementDefinition elementDefinition);

    TResult Visit(GroupDefinition groupDefinition);

    TResult Visit(AttributeGroupDefinition attributeGroupDefinition);

    TResult Visit(ComplexTypeSequence complexType);

    TResult Visit(ComplexTypeChoice complexType);

    TResult Visit(ComplexTypeSimpleContent complexType);

    TResult Visit(ComplexTypeElement complexType);

    TResult Visit(AttributeElement attributeElement);

    TResult Visit(AttributeGroupReference attributeGroupReference);

    TResult Visit(Sequence sequence);

    TResult Visit(Choice choice);

    TResult Visit(ElementType elementType);

    TResult Visit(ElementReference elementReference);

    TResult Visit(GroupReference groupReference);

    TResult Visit(Any any);
}
