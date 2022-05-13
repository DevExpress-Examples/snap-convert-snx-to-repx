using DevExpress.Snap.Core.API;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// A base class to generate content markup
    /// </summary>
    abstract class BufferedDocumentVisitor : DocumentVisitorBase
    {
        readonly StringBuilder buffer; //A buffer containing the resulting markup
        protected BufferedDocumentVisitor()
        {
            this.buffer = new StringBuilder();
        }
        protected StringBuilder Buffer { get { return buffer; } }
    }

    /// <summary>
    /// A class used to convert formatted text from Snap to markup supported in XtraReports
    /// </summary>
    class MarkupGenerator : BufferedDocumentVisitor
    {
        const char lastLowSpecial = '\x1f'; //lower boundary for special symbols
        const char firstHighSpecial = '\xffff'; //higher boundary for special symbols
        SnapDocument document; //The source document
        List<Field> processedFields; //A list containing already processed fields
        ReportGenerator generator; //Contains the resulting report

        public MarkupGenerator(ReportGenerator generator, SnapDocument document)
        {
            this.document = document;
            this.generator = generator;
            processedFields = new List<Field>();
        }


        /// <summary>
        /// Processes plain text
        /// </summary>
        /// <param name="text">DocumentText containing information about the current text portion</param>
        public override void Visit(DocumentText text)
        {
            //Get all opening tags
            List<MarkupTag> tags = GetTags(text.TextProperties);
            //Add them to the resulting markup
            for (int i = 0; i < tags.Count; i++)
                Buffer.Append(tags[i].OpeningTag);
            //Check if this text portion contains fields
            SnapSingleListItemEntity entity = GetSnapSingleItemField(text.Position);
            if (entity != null)
            {                         
                //If the field is not yet processed (added to markup), process it
                if (!processedFields.Contains(entity.Field))
                {
                    processedFields.Add(entity.Field);
                    ProcessField(entity);
                }
            }
            else
            {
                //Add text char by char
                int count = text.Length;
                for (int i = 0; i < count; i++)
                {
                    char ch = text.Text[i];
                    if (ch > lastLowSpecial && ch < firstHighSpecial)
                        Buffer.Append(ch);
                    else if (ch == '\x9' || ch == '\xA' || ch == '\xD')
                        Buffer.Append(ch);
                }
            }
            //Add closing tags to the resulting markup
            for (int i = tags.Count - 1; i >= 0; i--)
                Buffer.Append(tags[i].ClosingTag);
        }

        /// <summary>
        /// Returns a markup for a Snap entity
        /// </summary>
        /// <param name="entity">A Snap entity to process</param>
        protected virtual void ProcessField(SnapSingleListItemEntity entity)
        {
            //In the standard markup, fields should look like this: [UnitPrice!$0.00]
            //Get the data field name
            string dataFieldName = entity.DataFieldName;
            //Add an opening square bracket
            Buffer.Append("[");
            //If it is a parameter, it should start with ? (for example, [?parameter1])
            if (entity is SnapText && ((SnapText)entity).IsParameter)
                Buffer.Append("?");
            //Append the field name
            Buffer.Append(dataFieldName);
            //Check if the format string is specified
            if (entity is SnapText && !string.IsNullOrEmpty(((SnapText)entity).FormatString))
                //Add the format string to the field
                Buffer.AppendFormat("!{0}", ((SnapText)entity).FormatString);
            //Add a closing square bracket
            Buffer.Append("]");
        }

        /// <summary>
        /// Visits the paragraph start
        /// </summary>
        /// <param name="paragraphStart">An object containing paragraph properties</param>
        public override void Visit(DocumentParagraphStart paragraphStart)
        {
            //Add an opening paragraph tag with alignment settings
            Buffer.AppendFormat("<p {0}>", AlignmentConverter.Convert(paragraphStart.ParagraphProperties.Alignment));
        }

        /// <summary>
        /// Visits the paragraph end
        /// </summary>
        /// <param name="paragraphEnd">An object containing paragraph properties</param>
        public override void Visit(DocumentParagraphEnd paragraphEnd)
        {
            //Add a closing paragraph tag
            Buffer.Append("</p>");
        }

        /// <summary>
        /// Visits the hyperlink start
        /// </summary>
        /// <param name="hyperlinkStart">An object containing information about hyperlinks</param>
        public override void Visit(DocumentHyperlinkStart hyperlinkStart)
        {
            //Add a hyperlink tag
            Buffer.AppendFormat("<href value={0}>", hyperlinkStart.NavigateUri);
        }

        /// <summary>
        /// Visits the hyperlink end
        /// </summary>
        /// <param name="hyperlinkEnd">An object containing information about hyperlinks</param>
        public override void Visit(DocumentHyperlinkEnd hyperlinkEnd)
        {
            //Add a closing hyperlink tag
            Buffer.Append("</href>");
        }

        /// <summary>
        /// Checks if a Snap entity exists at the specified position
        /// </summary>
        /// <param name="position">A position to check</param>
        /// <returns>A SnapSingleListItemEntity instance</returns>
        private SnapSingleListItemEntity GetSnapSingleItemField(int position)
        {
            //Iterate through all fields in the document and check if the specified position is in the field's code range
            foreach (Field field in document.Fields)
                if (field.Range.Start.ToInt() <= position && field.CodeRange.End.ToInt() >= position)
                {
                    //If so, parse a field and check if it is SnapSingleListItemEntity (can be bound to a data field)
                    SnapEntity entity = document.ParseField(field);
                    if (entity is SnapSingleListItemEntity)
                    {    
                        //Return the found entity
                        return (SnapSingleListItemEntity)entity;
                    }
                }
            //If no field exists, return Null
            return null;
        }

        /// <summary>
        /// Collects tags based on text properties
        /// </summary>
        /// <param name="properties">An object containing information about formatting</param>
        /// <returns>A collection of tag objects</returns>
        List<MarkupTag> GetTags(ReadOnlyTextProperties properties)
        {
            //Create a list to store tag objects
            List<MarkupTag> list = new List<MarkupTag>();
            //Create a font tag
            MarkupTag fontTag = new MarkupTag("font");
            list.Add(fontTag);

            //Add a font name to the font tag
            string fontName = properties.FontName;
            fontTag.Parts.Add(string.Format("={0}{1}{0}", FontValueWrapperSymbol, fontName));

            //Add a font size to the font tag
            float fontSize = properties.FontSize;
            fontTag.Parts.Add(string.Format(" size={0:0.0}", fontSize));

            //Add a font color to the font tag
            Color foreColor = properties.ForeColor;
            fontTag.Parts.Add(string.Format(" color={0}", GetColorString(foreColor)));

            //Add a highlight color to the font tag
            Color backColor = properties.HighlightColor;
            //Get the string representation of the color
            string colorString = GetColorString(backColor);
            if (colorString != string.Empty)
                fontTag.Parts.Add(string.Format(" backcolor={0}", colorString));

            //If text is bold, add the corresponding tag
            if (properties.FontBold)
                list.Add(new MarkupTag("b"));
            //If text is italic, add the corresponding tag
            if (properties.FontItalic)
                list.Add(new MarkupTag("i"));
            //If text is underlined, add the corresponding tag
            if (properties.UnderlineType != UnderlineType.None)
                list.Add(new MarkupTag("u"));
            //If text is of the strikeout style, add the corresponding tag
            if (properties.StrikeoutType != StrikeoutType.None)
                list.Add(new MarkupTag("s"));
            //Return the list of tags
            return list;
        }

        /// <summary>
        /// Converts System.Drawing.Color to HTML representation
        /// </summary>
        /// <param name="color">The source color</param>
        /// <returns>HTML representation of the specified color</returns>
        private string GetColorString(Color color)
        {
            return ColorTranslator.ToHtml(color);
        }

        //The font name wrapper symbol
        public virtual string FontValueWrapperSymbol => "'";

        //The resulting markup
        public virtual string Text { 
            get 
            {                
                return string.Format("{0}</p>", Buffer.ToString());
            } 
        }


        /// <summary>
        /// This class stores information about tags used in a specific text portion
        /// </summary>
        class MarkupTag
        {
            //The base tag without any attributes
            string baseTag;
            //Additional parts (attributes, values, etc.)
            List<string> parts;
            public MarkupTag(string baseTag)
            {
                this.baseTag = baseTag;
                parts = new List<string>();
            }

            /// <summary>
            /// Generates an opening tag based on the base tag and additional parts
            /// </summary>
            /// <returns>A string containing a full opening tag</returns>
            private string GetOpeningTag()
            {
                //Create a string builder to construct a string
                StringBuilder sb = new StringBuilder("<");
                //Add a base tag
                sb.Append(baseTag);

                //Append additional parts
                for (int i = 0; i < parts.Count; i++)
                    sb.Append(parts[i]);

                //Close the tag
                sb.Append(">");

                //Return the generated value
                return sb.ToString();
            }

            public List<string> Parts { get => parts; } //A list of additional parts
            public string OpeningTag { get => GetOpeningTag(); } //An opening tag
            public string ClosingTag { get => string.Format("</{0}>", baseTag); } //A closing tag
        }
    }
}
