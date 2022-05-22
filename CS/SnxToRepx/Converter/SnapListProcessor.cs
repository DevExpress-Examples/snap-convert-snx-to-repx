using DevExpress.Snap;
using DevExpress.Snap.Core.API;
using DevExpress.XtraReports.UI;
using System.Linq;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// This class processes SnapList content
    /// </summary>
    class SnapListProcessor
    {
        SnapDocumentServer server; //The source SnapDocumentServer
        SnapList list; //The source SnapList
        ReportGenerator generator; //A ReportGenerator containing the constructed report

        public SnapListProcessor(ReportGenerator generator, SnapDocumentServer server, SnapList list)
        {
            this.server = server;
            this.list = list;
            this.generator = generator;
        }

        /// <summary>
        /// Processes a SnapList
        /// </summary>
        public void ProcessSnapList()
        {
            //Create a separate DetailReportBand for this list
            DetailReportBand listBand = new DetailReportBand();
            //Add it to the report and specify that this DetailReportBand is the currently processed report
            generator.CurrentReport.Bands.Add(listBand);
            generator.CurrentReport = listBand;

            //Copy data source settings
            CopyDataSettings();

            //Enable runtime customization for this list to access templates and other properties
            list.BeginUpdate();
            //Process the list header
            ProcessListHeader();
            //Process the row template
            ProcessRowTemplate();
            //Process each group one by one
            for (int i = 0; i < list.Groups.Count; i++)
            {
                SnapListGroupInfo info = list.Groups[i];
                ProcessGroupItem(info);
            }
            //Process the list footer
            ProcessListFooter();
            //Unlock the SnapList
            list.EndUpdate();
            //Apply sorting, filters, etc.
            CopyDataShapingSettings();
            //If no the DetailReportBand's ReportFooterBand was created, create and use it as the currently processed band
            if (listBand.Bands[BandKind.ReportFooter] == null)
                generator.CurrentBand = listBand.Bands.Create(BandKind.ReportFooter);
            else
                generator.CurrentBand = listBand.Bands[BandKind.ReportFooter];
        }

        /// <summary>
        /// Copies data source settings from the Snap list
        /// </summary>
        private void CopyDataSettings()
        {
            //Get data source information
            DataSourceInfo dsInfo = server.Document.DataSources[list.DataSourceName];
            XtraReportBase currentReport = generator.CurrentReport;

            //If data binding exists
            if (dsInfo != null)
            {
                //Specify the DataSource and DataMember properties
                currentReport.DataSource = dsInfo.DataSource;
                currentReport.DataMember = list.DataMember;

                //IMPORTANT! Detail SnapLists do not have the DataSource property specified. Instead, they use the parent's data source.
                //XtraReports require the DataSource property to be specified. Moreover, the DataMember property should contain the full path
                //to the data member (for example, "Products.ProductOrders" in XtraReports, while Snap uses "ProductOrders" in this case).

                //Check if the data source is null
                if (currentReport.DataSource == null && currentReport.Report != null && currentReport.Report.DataMember != string.Empty)
                {
                    //Get the parent's data source and use it as a data source for the current report
                    string dataMember = currentReport.DataMember;
                    currentReport.DataSource = currentReport.Report.DataSource;
                    //Combine the current and parent's data members to get the full path
                    currentReport.DataMember = $"{currentReport.Report.DataMember}.{dataMember}";
                }
                //Copy calculated fields from the data source info
                CopyCalculatedFields(dsInfo, list.DataMember);
            }
        }

        /// <summary>
        /// Copies calculated fields to XtraReport
        /// </summary>
        /// <param name="dsInfo">Data source information</param>
        /// <param name="dataMember">The current data member's name</param>
        private void CopyCalculatedFields(DataSourceInfo dsInfo, string dataMember)
        {
            //Iterate through calculated fields
            for (int i = 0; i < dsInfo.CalculatedFields.Count; i++)
            {
                DevExpress.Snap.Core.API.CalculatedField calcField = dsInfo.CalculatedFields[i];
                //If a calculated field does not yet exist 
                if (generator.GetCalculatedField(calcField.DisplayName) == null)
                {
                    //Create CalculatedField in XtraReport and copy field settings
                    DevExpress.XtraReports.UI.CalculatedField reportField = new DevExpress.XtraReports.UI.CalculatedField(calcField.DataSource, calcField.DataMember);
                    reportField.DisplayName = calcField.DisplayName;
                    reportField.FieldType = calcField.FieldType;
                    reportField.Expression = calcField.Expression;
                    reportField.Name = calcField.Name;

                    //Add this calculated field to the report
                    generator.CurrentReport.RootReport.CalculatedFields.Add(reportField);
                }
            }
        }

        /// <summary>
        /// Copies filter and sort settings
        /// </summary>
        private void CopyDataShapingSettings()
        {
            //Copy the first filter string
            if (list.Filters.Count > 0)
                generator.CurrentReport.FilterString = list.Filters[0];

            DetailBand band = (DetailBand)generator.CurrentReport.Bands[BandKind.Detail];
            //Iterate through all sort fields and add corresponding items to the current DetailBand
            foreach (SnapListGroupParam sortItem in list.Sorting)
            {     
                //Convert the sort order to XtraReports sort order
                XRColumnSortOrder sortOrder = ColumnSortOrderConverter.Convert(sortItem.SortOrder);
                band.SortFields.Add(new GroupField() { FieldName = sortItem.FieldName, SortOrder = sortOrder });
            }
        }

        /// <summary>
        /// Processes the list header
        /// </summary>
        private void ProcessListHeader()
        {
            //Get the list header
            SnapDocument listHeader = list.ListHeader;
            //Process the inner document and add its content to the ReportHeader band
            ProcessTemplateDocument(listHeader, BandKind.ReportHeader);
        }

        /// <summary>
        /// Processes the list footer
        /// </summary>
        private void ProcessListFooter()
        {
            //Get the list footer
            SnapDocument listFooter = list.ListFooter;
            //Process the inner document and add its content to the ReportFooter band
            ProcessTemplateDocument(listFooter, BandKind.ReportFooter);
        }

        /// <summary>
        /// Processes the row template
        /// </summary>
        private void ProcessRowTemplate()
        {
            //Get the row template
            SnapDocument rowTemplate = list.RowTemplate;
            //Process the inner document and add its content to DetailBand
            ProcessTemplateDocument(rowTemplate, BandKind.Detail);
        }

        /// <summary>
        /// Processes the group item
        /// </summary>
        /// <param name="info">Information about the current group item</param>
        private void ProcessGroupItem(SnapListGroupInfo info)
        {
            //If the group header exists, get the template and add its content to the GroupHeader band
            if (info.Header != null)
            {
                SnapDocument groupHeaderTemplate = info.Header;
                ProcessTemplateDocument(groupHeaderTemplate, BandKind.GroupHeader);
            }
            //Specify the initial group header to link headers and footers
            int groupLevel = 0;
            //Get the current GroupHeaderBand
            if (generator.CurrentBand is GroupHeaderBand)
            {
                GroupHeaderBand groupHeader = (GroupHeaderBand)generator.CurrentBand;                
                //Iterate through group items (each item may be grouped by multiple fields)
                for (int i = 0; i < info.Count; i++)
                {
                    //Get group parameters and add corresponding GroupFields
                    SnapListGroupParam groupParam = info[i];
                    groupHeader.GroupFields.Add(new GroupField(groupParam.FieldName, ColumnSortOrderConverter.Convert(groupParam.SortOrder)));
                }
                //Get the current grouping level
                groupLevel = groupHeader.Level;
            }
            //If the group footer exists, get the template and add its content to the GroupFooter band
            if (info.Footer != null)
            {
                SnapDocument groupHeaderTemplate = info.Footer;
                ProcessTemplateDocument(groupHeaderTemplate, BandKind.GroupFooter);
            }

            if (generator.CurrentBand is GroupFooterBand)
            {
                GroupFooterBand groupFooter = (GroupFooterBand)generator.CurrentBand;
                groupFooter.Level = groupLevel;
            }
        }       

        /// <summary>
        /// Processes the template and adds its content to the Band of the specified kind
        /// </summary>
        /// <param name="template">The source document</param>
        /// <param name="bandKind">The Band kind to create/access a Band</param>
        private void ProcessTemplateDocument(SnapDocument template, BandKind bandKind)
        {
            //If the template is not empty
            if (template.Range.Length > 1)
            {
                //Create the corresponding band and set it as an active Band
                Band currentBand = generator.CurrentReport.Bands.Create(bandKind);
                generator.CurrentBand = currentBand;

                //Create a temporary SnapDocumentServer and copy the template document there
                using (SnapDocumentServer tempServer = new SnapDocumentServer())
                {
                    tempServer.SnxBytes = template.SaveDocument(SnapDocumentFormat.Snap);
                    //Copy data sources
                    for (int i = 0; i < server.Document.DataSources.Count; i++)
                        if (server.Document.DataSources[i].DataSource != null &&
                            !tempServer.Document.DataSources.Any((ds) =>
                            ds.DataSourceName == server.Document.DataSources[i].DataSourceName))
                            tempServer.Document.DataSources.Add(server.Document.DataSources[i]);
                    //Create a SnapLayoutProcessor instance and process the document layout
                    SnapLayoutProcessor processor = new SnapLayoutProcessor(generator, tempServer);
                    processor.ProcessDocument();                    
                }

                //Restore the current Band (it may change during layout processing)
                generator.CurrentBand = currentBand;
                //Set the Band height to 0. It will automatically adjust its size based on content
                generator.CurrentBand.HeightF = 0F;
            }
        }
    }
}
