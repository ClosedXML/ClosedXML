using ClosedXML.IO.CodeGen.Model;
using ClosedXML.IO.CodeGen.Model.Elements;
using ClosedXML.IO.CodeGen.Model.SimpleTypes;
using ClosedXML.IO.CodeGen.Model.TopLevel;

namespace ClosedXML.IO.CodeGen;

/// <summary>
/// A visitor for visitor pattern.
/// </summary>
/// <typeparam name="TResult">Type of returned value.</typeparam>
public interface IXsdVisitor<out TResult>
{
    TResult Visit(Schema schema);

    TResult Visit(ComplexTypeSequence complexType);

    TResult Visit(ComplexTypeChoice complexType);

    TResult Visit(ComplexTypeSimpleContent complexType);

    TResult Visit(ComplexTypeElement complexType);

    TResult Visit(ElementDefinition elementDefinition);

    TResult Visit(GroupDefinition groupDefinition);

    TResult Visit(AttributeGroupDefinition attributeGroupDefinition);

    TResult Visit(AttributeElement attributeElement);

    TResult Visit(AttributeGroupReference attributeGroupReference);

    TResult Visit(SimpleType simpleType);

    TResult Visit(SimpleTypeList simpleTypeList);

    TResult Visit(SimpleTypeUnion simpleTypeUnion);

    TResult Visit(GroupReference groupReference);

    TResult Visit(Any any);

    TResult Visit(Choice choice);

    TResult Visit(ElementType elementType);

    TResult Visit(Sequence sequence);

    TResult Visit(ElementReference elementReference);
}
