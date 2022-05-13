using DevExpress.Snap;
using DevExpress.Snap.Core.API;
using DevExpress.XtraPrinting;
using DevExpress.XtraReports.UI;
using DevExpress.XtraRichEdit.API.Layout;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// This class processes text portions
    /// </summary>
    class ParagraphProcessor : ContentProcessorBase
    {
        List<LayoutRow> rows; //LayoutRows that contain processed text parts
        public ParagraphProcessor(ReportGenerator generator, SnapDocumentServer server, List<LayoutRow> rows) : base(generator, server)
        {
            this.rows = rows;
        }

        /// <summary>
        /// Generates a markup and passes it to the corresponding XRControl
        /// </summary>
        public void ProcessParagraphs()
        {
            //Get a range of processed paragraphs
            DocumentRange range = Server.Document.CreateRange(rows[0].Range.Start, rows[rows.Count - 1].Range.Start + rows[rows.Count - 1].Range.Length - rows[0].Range.Start - 1);
            //Check if this text portion contains summaries
            DevExpress.Snap.Core.Fields.SummaryRunning summaryRunning = GetSummary(range);
            MarkupGenerator markupGenerator;
            //If a summary calculation exists, an expression is generated
            if (summaryRunning != DevExpress.Snap.Core.Fields.SummaryRunning.None)
                markupGenerator = new ExpressionMarkupGenerator(Generator, Server.Document);
            else
                //Otherwise, a markup used in mail-merge reports is generated
                markupGenerator = new MarkupGenerator(Generator, Server.Document);

            //Create a DocumentIterator instance to iterate through all text parts within the specified range
            DocumentIterator iterator = new DocumentIterator(range, true);
            while (iterator.MoveNext())
                iterator.Current.Accept(markupGenerator);

            //Get the resulting markup
            string markup = markupGenerator.Text;
            XRControl contentControl;
            //If the current control is already created (for example, XRTableCell), then use it
            if (Generator.CurrentControl != null)
                contentControl = Generator.CurrentControl;
            else
            {
                //Otherwise, create a new control, set its properties, and add it to the parent Band
                contentControl = GenerateControl();                
                contentControl.CanShrink = true;
                contentControl.Padding = GetPaddingInfo(range);
                contentControl.BoundsF = CalculateParagraphBounds(rows, Generator.CurrentBand);
                Generator.CurrentBand.Controls.Add(contentControl);
            }

            //If a summary calculation exists
            if (summaryRunning != DevExpress.Snap.Core.Fields.SummaryRunning.None && contentControl is XRLabel)
            {
                //Add an expression binding and specify the area for which summary is calculated
                contentControl.ExpressionBindings.Add(new ExpressionBinding("Text", markup));
                ((XRLabel)contentControl).Summary.Running = SummaryConverter.Convert(summaryRunning);
            }
            else
                //Otherwise, use the default markup
                contentControl.Text = markup;
        }

        /// <summary>
        /// Checks if summary functions are used in the specified range
        /// </summary>
        /// <param name="range">A document range to check</param>
        /// <returns>The SummaryRunning value</returns>
        private DevExpress.Snap.Core.Fields.SummaryRunning GetSummary(DocumentRange range)
        {
            //Get all fields in the specified range
            ReadOnlyFieldCollection fields = Server.Document.Fields.Get(range);
            //Iterate through fields
            foreach (Field field in fields)
            {
                //Check if it is a SnapText that has SummaryRunning specified
                SnapEntity entity = Server.Document.ParseField(field);
                if (entity is SnapText && ((SnapText)entity).SummaryRunning != DevExpress.Snap.Core.Fields.SummaryRunning.None
                    && ((SnapText)entity).SummaryFunc != DevExpress.Data.SummaryItemType.None)
                    //Return the SummaryRunning value
                    return ((SnapText)entity).SummaryRunning;
            }
            //If no fields with summaries are found, return None
            return DevExpress.Snap.Core.Fields.SummaryRunning.None;
        }

        /// <summary>
        /// Gets the resulting bounds for XRControl within the parent control
        /// </summary>
        /// <param name="rows">Layout rows containing text</param>
        /// <param name="parentControl">The parent control to calculate boundaries</param>
        /// <returns>Control boundaries</returns>
        private RectangleF CalculateParagraphBounds(List<LayoutRow> rows, XRControl parentControl)
        {
            //Get the paragraph width
            int paragraphWidth = rows[0].Bounds.Width;
            //Iterate through rows to calculate the resulting width and height
            int paragraphHeight = rows[0].Bounds.Height;
            for (int i = 1; i < rows.Count; i++)
            {
                if (rows[i].Bounds.Width > paragraphWidth)
                    paragraphWidth = rows[i].Bounds.Width;

                paragraphHeight += rows[i].Bounds.Height;
            }

            //Calculating initial boundaries
            Rectangle bounds = new Rectangle(Math.Max(rows[0].GetRelativeBounds(rows[0].Parent).Left, 0),
                                             0,
                                             paragraphWidth,
                                             paragraphHeight);

            //Convert boundaries to measurement units used in XtraReports
            bounds = XRConvert.Convert(bounds, GraphicsDpi.Twips, GraphicsDpi.HundredthsOfAnInch);
            //Take paddings into account
            bounds.Offset(-parentControl.Padding.Left, ReportGenerator.GetVerticalOffset(parentControl));
            //Return boundaries
            return bounds;
        }

        /// <summary>
        /// Creates a control to display content
        /// </summary>
        /// <returns>An XRControl for future use</returns>
        private XRControl GenerateControl()
        {
            //IMPORTANT! Auxiliary Snap controls, such as checkboxes, can be located in the middle of formatted text.
            //XtraReports do not support such layout, so textual representation is used
            XRLabel label = new XRLabel();
            //Allow HTML-style markup
            label.AllowMarkupText = true;
            //Remove borders
            label.Borders = BorderSide.None;            
            //Return the control
            return label;
        }
    }
}
