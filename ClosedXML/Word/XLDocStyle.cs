using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ClosedXML.Word
{
    public class XLDocStyle : IXLDocStyle
    {
        public static void CreateAndAddCharacterStyle( StyleDefinitionsPart styleDefinitionsPart,
        string styleid, string stylename, string aliases = "" )
        {
            // Get access to the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;

            // Create a new character style and specify some of the attributes.
            Style style = new Style( )
            {
                Type = StyleValues.Character,
                StyleId = styleid,
                CustomStyle = true
            };

            // Create and add the child elements (properties of the style).
            Aliases aliases1 = new Aliases( ) { Val = aliases };
            StyleName styleName1 = new StyleName( ) { Val = stylename };
            LinkedStyle linkedStyle1 = new LinkedStyle( ) { Val = "Test" };
            if ( aliases != "" )
                style.Append( aliases1 );
            style.Append( styleName1 );
            style.Append( linkedStyle1 );

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties( );
            Bold bold1 = new Bold( );
            Color color1 = new Color( ) { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts( ) { Ascii = "Tahoma" };
            Italic italic1 = new Italic( );
            // Specify a 24 point size.
            FontSize fontSize1 = new FontSize( ) { Val = "48" };
            styleRunProperties1.Append( font1 );
            styleRunProperties1.Append( fontSize1 );
            styleRunProperties1.Append( color1 );
            styleRunProperties1.Append( bold1 );
            styleRunProperties1.Append( italic1 );

            // Add the run properties to the style.
            style.Append( styleRunProperties1 );

            // Add the style to the styles part.
            styles.Append( style );
        }

        // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
        public static StyleDefinitionsPart AddStylesPartToPackage( WordprocessingDocument doc )
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>( );
            Styles root = new Styles( );
            root.Save( part );
            return part;
        }
    }
}
