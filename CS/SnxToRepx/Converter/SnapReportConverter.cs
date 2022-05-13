using DevExpress.DataAccess.Sql;
using DevExpress.Snap;
using DevExpress.Snap.Core.API;
using DevExpress.XtraReports.UI;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.IO;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// This class converts an SNX report to XtraReport
    /// </summary>
    public class SnapReportConverter: IDisposable
    {
        ReportGenerator generator; //The Report generator instance that contains the current report

        /// <summary>
        /// Converts an existing SnapDocument instance to XtraReport
        /// </summary>
        /// <param name="src">The source SnapDocument</param>
        /// <returns>An XtraReport instance</returns>
        public XtraReport Convert(SnapDocument src)
        {
            //Save a document to a byte array and call the GenerateReport method
            return GenerateReport(src.SaveDocument(SnapDocumentFormat.Snap));                
        }

        /// <summary>
        /// Converts an SNX document stored in the file system to XtraReport
        /// </summary>
        /// <param name="path">A path to the SNX file</param>
        /// <returns>An XtraReport instance</returns>
        public XtraReport Convert(string path)
        {
            return GenerateReport(File.ReadAllBytes(path));
        }

        /// <summary>
        /// Converts an SNX document stored in a stream to XtraReport
        /// </summary>
        /// <param name="stream">A Stream object containing the SNX document</param>
        /// <returns>An XtraReport instance</returns>
        public XtraReport Convert(Stream stream)
        {
            //Use a SnapDocumentServer instance to load a document and convert it to a byte array
            using (SnapDocumentServer server = new SnapDocumentServer())
            {
                server.Document.ConfigureDataConnection += OnConfigureDataConnection;
                server.LoadDocument(stream);
                server.Document.ConfigureDataConnection -= OnConfigureDataConnection;
                //Call the GenerateReport method to generate XtraReport                
                return GenerateReport(server.SnxBytes);
            }
        }

        /// <summary>
        /// Loads a document and converts it to XtraReport
        /// </summary>
        /// <param name="snxBytes">A byte array containing the input SNX document</param>
        /// <returns>An XtraReport instance</returns>
        private XtraReport GenerateReport(byte[] snxBytes)
        {
            //Set this property to ensure that subsequent tables are not merged
            DevExpress.XtraRichEdit.RichEditControlCompatibility.MergeSuccessiveTables = false;
            try
            {
                //Create a new SnapDocumentServer to operate with the document
                using (SnapDocumentServer server = new SnapDocumentServer())
                {
                    //Handle the ConfigureDataConnection event to allow users to specify data connection settings
                    server.Document.ConfigureDataConnection += OnConfigureDataConnection;
                    //Pass SnxBytes to server
                    server.SnxBytes = snxBytes;

                    server.Document.ConfigureDataConnection -= OnConfigureDataConnection;
                    //Return the loaded SnapDocument instance
                    SnapDocument source = server.Document;

                    //Set units to Document to simplify conversion to report units
                    source.Unit = DevExpress.Office.DocumentUnit.Document;

                    //Create an XtraReport instance and initialize ReportGenerator
                    XtraReport report = new XtraReport();
                    generator = new ReportGenerator(report);
                    //Apply global settings (paper, parameters, etc.) to the created report
                    generator.SetupReport(source);

                    //Ensure that SnapLists start in new paragraphs
                    PrepareSnapListsForExport(source);

                    //Create a SnapLayoutProcessor instance to process the document layout
                    SnapLayoutProcessor layoutProcessor = new SnapLayoutProcessor(generator, server);
                    //Start processing
                    layoutProcessor.ProcessDocument();

                    //Add header and footer data to the report
                    ProcessHeadersFooters(server);

                    //Empty bands may appear after the XtraReport layout is built. Call this method to remove redundant bands
                    generator.CleanEmptyBands();
                    //Return the resulting report
                    return report;
                }
            } catch
            {
                return null;
            }
        }

        /// <summary>
        /// Handles the ConfigureDataConnection event and forwards it to SnapReportConverter.ConfigureDataConnection
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnConfigureDataConnection(object sender, ConfigureDataConnectionEventArgs e)
        {
            if (ConfigureDataConnection != null)
                ConfigureDataConnection(this, e);
        }

        /// <summary>
        /// This method ensures that every SnapList starts with a new paragraph
        /// </summary>
        /// <param name="document">A source document</param>
        private void PrepareSnapListsForExport(SnapDocument document)
        {
            //IMPORTANT! Sometimes SnapLists may start on the same row with static content.
            //In this situation, it is difficult to process SnapLists along with other content.
            //To avoid this situation, ensure that every SnapList starts with a new paragraph.

            //Lock all updates
            document.BeginUpdate();
            //Iterate through all fields
            for (int i = 0; i < document.Fields.Count; i++)
            {
                Field field = document.Fields[i];
                //Parse the current field to check if it is a SnapList
                SnapEntity entity = document.ParseField(field);
                if (entity is SnapList)
                {
                    //Check if this SnapList starts at the beginning of a paragraph
                    Paragraph parentParagraph = document.Paragraphs.Get(field.Range.Start);
                    if (parentParagraph.Range.Start.ToInt() != field.Range.Start.ToInt())
                        //If not, insert a paragraph just before the field
                        document.Paragraphs.Insert(field.Range.Start);
                }
            }
            //Disable locking
            document.EndUpdate();
        }


        private void ProcessHeadersFooters(SnapDocumentServer sourceServer)
        {
            //IMPORTANT! There is no such term as First/Primary/Odd/Even for headers and footers in XtraReports.
            //Thus, only the primary header and footer are processed

            //Get the first section and check if it has a header of the Primary type
            Section section = sourceServer.Document.Sections[0];
            if (section.HasHeader(HeaderFooterType.Primary))
            {
                //If so, copy its content to a separate SnapDocumentServer and process its content
                using (SnapDocumentServer server = new SnapDocumentServer())
                {
                    SubDocument headerDocument = section.BeginUpdateHeader(HeaderFooterType.Primary);
                    server.Document.AppendDocumentContent(headerDocument.Range, InsertOptions.KeepSourceFormatting);                   
                    section.EndUpdateHeader(headerDocument);
                    ProcessSubDocument(server, BandKind.PageHeader);
                }
            }

            //Check if the section has a footer of the Primary type
            if (section.HasFooter(HeaderFooterType.Primary))
            {
                //If so, copy its content to a separate SnapDocumentServer and process its content
                using (SnapDocumentServer server = new SnapDocumentServer())
                {
                    SubDocument footerDocument = section.BeginUpdateFooter(HeaderFooterType.Primary);
                    server.Document.AppendDocumentContent(footerDocument.Range, InsertOptions.KeepSourceFormatting);                    
                    section.EndUpdateFooter(footerDocument);
                    ProcessSubDocument(server, BandKind.PageFooter);
                }
            }          
        }

        /// <summary>
        /// Copies header/footer information to PageHeaderBand/PageFooterBand
        /// </summary>
        /// <param name="server">A SnapDocumentServer instance containing the source document</param>
        /// <param name="targetBandKind">The target BandKind to create a Band</param>
        private void ProcessSubDocument(SnapDocumentServer server, BandKind targetBandKind)
        {
            //IMPORTANT! In Snap, you can place page fields at any position (even in formatted text)
            //XRPageInfo does not support this, so all fields are converted to static values before processing
            server.Document.UnlinkAllFields();
            //Specify the current band to place content into
            generator.CurrentBand = generator.CurrentReport.Bands.Create(targetBandKind);

            //Create a SnapLayoutProcessor instance to process the document layout
            SnapLayoutProcessor layoutProcessor = new SnapLayoutProcessor(generator, server);
            layoutProcessor.ProcessDocument();
        }

        public void Dispose()
        {
            ConfigureDataConnection = null;
            generator = null;
        }

        public event ConfigureDataConnectionEventHandler ConfigureDataConnection;
    }
}
