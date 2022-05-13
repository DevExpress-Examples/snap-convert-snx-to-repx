using DevExpress.Snap;
using DevExpress.XtraPrinting;
using DevExpress.XtraReports.UI;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// A base class for content processors
    /// </summary>
    abstract class ContentProcessorBase
    {
        SnapDocumentServer server; //A SnapDocumentServer containing the processed document
        ReportGenerator generator; //A ReportGenerator that stores the resulting XtraReport

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="generator">A ReportGenerator that stores the resulting XtraReport</param>
        /// <param name="server">A SnapDocumentServer containing the processed document</param>
        public ContentProcessorBase(ReportGenerator generator, SnapDocumentServer server)
        {
            //Assign values
            this.server = server;
            this.generator = generator;
        }

        /// <summary>
        /// Retrieves padding for current content
        /// </summary>
        /// <param name="range">The processed document range</param>
        /// <returns>PaddingInfo containing required paddings</returns>
        protected virtual PaddingInfo GetPaddingInfo(DocumentRange range)
        {
            //Get paragraph properties
            ParagraphProperties parProps = server.Document.BeginUpdateParagraphs(range);

            //Set initial spacing values
            int spacingBefore = 0;
            int spacingAfter = 0;

            //Obtain spacing
            if (parProps.SpacingBefore != null)
                spacingBefore = (int)XRConvert.Convert((float)parProps.SpacingBefore, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);

            if (parProps.SpacingAfter != null)
                spacingAfter = (int)XRConvert.Convert((float)parProps.SpacingAfter, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);

            server.Document.EndUpdateParagraphs(parProps);
            //Create and return a new PaddingInfo object containing required values
            PaddingInfo paddingInfo = new PaddingInfo(0, 0, spacingBefore, spacingAfter);
            return paddingInfo;
        }

        //Properties
        protected SnapDocumentServer Server => server;
        protected ReportGenerator Generator => generator;
    }
}
