using System.Xml.Linq;

namespace MonoReport.Services
{
    internal static class WordXml
    {
        public static readonly XNamespace WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private const string WordprocessingNamespacePrefix = "w";

        public static readonly XName WordprocessingNamespaceDeclarationName = XNamespace.Xmlns + WordprocessingNamespacePrefix;



        public const string FieldBeginValue = "begin";
        public const string FieldSeparatorValue = "separate";
        public const string FieldEndValue = "end";
        public const string FalseValue = "false";
        public const string PreserveSpaceValue = "preserve";
        public const string DefaultHeaderFooterTypeValue = "default";
        public const string SingleUnderlineValue = "single";

        public static readonly XName SpaceAttributeName = XNamespace.Xml + "space";

        public static readonly XName RunElementName = WordprocessingNamespace + "r";
        public static readonly XName RunPropertiesElementName = WordprocessingNamespace + "rPr";
        public static readonly XName TextElementName = WordprocessingNamespace + "t";
        public static readonly XName SimpleFieldElementName = WordprocessingNamespace + "fldSimple";
        public static readonly XName FieldCharacterElementName = WordprocessingNamespace + "fldChar";
        public static readonly XName InstructionTextElementName = WordprocessingNamespace + "instrText";
        public static readonly XName ParagraphElementName = WordprocessingNamespace + "p";
        public static readonly XName TableRowElementName = WordprocessingNamespace + "tr";
        public static readonly XName TableCellElementName = WordprocessingNamespace + "tc";
        public static readonly XName TableElementName = WordprocessingNamespace + "tbl";
        public static readonly XName VerticalMergeElementName = WordprocessingNamespace + "vMerge";
        public static readonly XName DrawingElementName = WordprocessingNamespace + "drawing";
        public static readonly XName PictureElementName = WordprocessingNamespace + "pict";
        public static readonly XName BreakElementName = WordprocessingNamespace + "br";
        public static readonly XName TabElementName = WordprocessingNamespace + "tab";
        public static readonly XName TablePropertiesElementName = WordprocessingNamespace + "tblPr";
        public static readonly XName TableGridElementName = WordprocessingNamespace + "tblGrid";
        public static readonly XName GridColumnElementName = WordprocessingNamespace + "gridCol";
        public static readonly XName RunFontsElementName = WordprocessingNamespace + "rFonts";
        public static readonly XName BoldElementName = WordprocessingNamespace + "b";
        public static readonly XName ItalicElementName = WordprocessingNamespace + "i";
        public static readonly XName ColorElementName = WordprocessingNamespace + "color";
        public static readonly XName FontSizeElementName = WordprocessingNamespace + "sz";
        public static readonly XName HighlightElementName = WordprocessingNamespace + "highlight";
        public static readonly XName UnderlineElementName = WordprocessingNamespace + "u";
        public static readonly XName LanguageElementName = WordprocessingNamespace + "lang";
        public static readonly XName DocumentElementName = WordprocessingNamespace + "document";
        public static readonly XName BodyElementName = WordprocessingNamespace + "body";
        public static readonly XName HyperlinkElementName = WordprocessingNamespace + "hyperlink";
        public static readonly XName HeaderElementName = WordprocessingNamespace + "hdr";
        public static readonly XName FooterElementName = WordprocessingNamespace + "ftr";
        public static readonly XName SectionPropertiesElementName = WordprocessingNamespace + "sectPr";
        public static readonly XName HeaderReferenceElementName = WordprocessingNamespace + "headerReference";
        public static readonly XName FooterReferenceElementName = WordprocessingNamespace + "footerReference";
        public static readonly XName DirtyAttributeName = WordprocessingNamespace + "dirty";
        public static readonly XName InstructionAttributeName = WordprocessingNamespace + "instr";
        public static readonly XName FieldCharacterTypeAttributeName = WordprocessingNamespace + "fldCharType";
        public static readonly XName WidthAttributeName = WordprocessingNamespace + "w";
        public static readonly XName AsciiFontAttributeName = WordprocessingNamespace + "ascii";
        public static readonly XName HighAnsiFontAttributeName = WordprocessingNamespace + "hAnsi";
        public static readonly XName ValueAttributeName = WordprocessingNamespace + "val";
        public static readonly XName AnchorAttributeName = WordprocessingNamespace + "anchor";
        public static readonly XName TypeAttributeName = WordprocessingNamespace + "type";

        public static readonly XName RelationshipIdAttributeName = RelationshipsNamespace + "id";
    }
}
