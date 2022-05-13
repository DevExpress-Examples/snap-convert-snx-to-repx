using DevExpress.Snap.Core.API;
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
    /// A class that contains information about the generated XtraReport
    /// </summary>
    class ReportGenerator
    {
        XtraReport report; //The XtraReport instance
        XtraReportBase currentReport; //The currently processed XtraReportBase instance (DetailReportBand or XtraReport)
        Band currentBand; //The currently populated Band

        public ReportGenerator(XtraReport report)
        {
            this.report = report;
        }

        /// <summary>
        /// Specifies global properties, which are applied to the entire report
        /// </summary>
        /// <param name="source">The source Snap document</param>
        public void SetupReport(SnapDocument source)
        {            
            //Create a DetailBand instance. An XtraReport object should always contain DetailBand
            CreateDetailBand();
            //Copy mail-merge settings (if any)
            CopyMailMergeOptions(source);
            //Copy page settings
            CopyPageSettings(source);
            //Copy parameters
            CopyParameters(source);

            //Specify initial values
            CurrentReport = report;
            CurrentBand = report.Bands[BandKind.Detail];
            CurrentControl = null;
        }

        /// <summary>
        /// Copies mail-merge settings to XtraReport
        /// </summary>
        /// <param name="source">The source Snap document</param>
        private void CopyMailMergeOptions(SnapDocument source)
        {
            //Get mail-merge options
            DevExpress.Snap.Core.Options.SnapMailMergeExportOptions options = source.CreateSnapMailMergeExportOptions();
            //Specify data settings
            report.DataSource = options.DataSource;
            report.DataMember = options.DataMember;
        }

        /// <summary>
        /// Specifies page settings
        /// </summary>
        /// <param name="source">The source Snap document</param>
        private void CopyPageSettings(SnapDocument source)
        {            
            //IMPORTANT! Snap document may contain multiple sections with different page settings. XtraReports do not support
            //this feature. So, only the first section's settings are taken into account
            Section section = source.Sections[0];

            //Get page settings
            report.PaperKind = section.Page.PaperKind;
            report.PageWidth = (int)XRConvert.Convert(section.Page.Width, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);
            report.PageHeight = (int)XRConvert.Convert(section.Page.Height, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);
            report.Landscape = section.Page.Landscape;

            //Get margins
            report.Margins.Left = (int)XRConvert.Convert(section.Margins.Left, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);
            report.Margins.Top = (int)XRConvert.Convert(section.Margins.Top, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);
            report.Margins.Right = (int)XRConvert.Convert(section.Margins.Right, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);
            report.Margins.Bottom = (int)XRConvert.Convert(section.Margins.Bottom, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);            
        }

        /// <summary>
        /// Copies parameters to XtraReport
        /// </summary>
        /// <param name="source">The source Snap document</param>
        private void CopyParameters(SnapDocument source)
        {
            //Iterate through parameters in the Snap document
            for (int i = 0; i < source.Parameters.Count; i++)
            {
                Parameter snapParameter = source.Parameters[i];
                //Create an XtraReport parameter and add it to the Parameters collection
                DevExpress.XtraReports.Parameters.Parameter reportParameter = new DevExpress.XtraReports.Parameters.Parameter();
                report.Parameters.Add(reportParameter);

                //Copy parameter settings
                reportParameter.Name = snapParameter.Name;
                reportParameter.Type = snapParameter.Type;
                reportParameter.Value = snapParameter.Value;
            }
        }

        /// <summary>
        /// Creates DetailBand
        /// </summary>
        private void CreateDetailBand()
        {            
            report.Bands.Create(BandKind.Detail);
        }

        /// <summary>
        /// Removes redundant empty bands
        /// </summary>
        public void CleanEmptyBands()
        {
            //Call the core method
            CleanEmptyBandsCore(report);
        }

        /// <summary>
        /// Removes redundant empty bands
        /// </summary>
        /// <param name="curReport">The currently processed XtraReportBase</param>
        private void CleanEmptyBandsCore(XtraReportBase curReport)
        {
            //Iterate through all bands
            for (int i = curReport.Bands.Count - 1; i >= 0; i--)
            {
                //If the current band is a nested report (DetailReportBand)
                if (curReport.Bands[i] is XtraReportBase)
                    //Call this method for the nested report to clean its bands
                    CleanEmptyBandsCore((XtraReportBase)curReport.Bands[i]);
                else
                //If no controls in a Band or this is not a mandatory band (TopMarginBand, BottomMarginBand, and DetailBand should always exist,
                //GroupHeaderBand specifies groups and can be empty)
                    if (curReport.Bands[i].Controls.Count == 0)
                    if (!(curReport.Bands[i] is TopMarginBand || curReport.Bands[i] is BottomMarginBand || curReport.Bands[i] is DetailBand
                        || curReport.Bands[i] is GroupHeaderBand))
                        //Remove a Band
                        curReport.Bands.RemoveAt(i);
                    else
                        //Resize a Band to follow its content size (it will be automatically resized afterwards based on its content)
                        curReport.Bands[i].HeightF = 0;                
            }
        }

        /// <summary>
        /// Locates a CalculatedField by its name
        /// </summary>
        /// <param name="displayName">A name of the field</param>
        /// <returns>The located CalculatedField instance</returns>
        public DevExpress.XtraReports.UI.CalculatedField GetCalculatedField(string displayName)
        {
            //IMPORTANT! In Snap, different data sources may contain calculated fields with the same name
            //To avoid potential conflicts, adding a field with the same name is disabled
            //Iterate through fields
            for (int i = 0; i < report.CalculatedFields.Count; i++)
                if (report.CalculatedFields[i].DisplayName == displayName)
                    //If the field's name equals to the specified name, return the field
                    return report.CalculatedFields[i];
            //Otherwise, return null
            return null;
        }

        /// <summary>
        /// Gets the available Y position to insert a new control
        /// </summary>
        /// <param name="control">A container control</param>
        /// <returns>An integer offset value</returns>
        public static int GetVerticalOffset(XRControl control)
        {
            //If no controls in the container, return 0
            if (control.Controls.Count == 0)             
                return 0;
            else
            {
                //Iterate through nested controls to get the bottom-most position
                int offset = (int)control.Controls[0].BoundsF.Bottom;
                for (int i = 1; i < control.Controls.Count; i++)
                    if (control.Controls[i].BoundsF.Bottom > offset)
                        offset = (int)control.Controls[i].BoundsF.Bottom;

                //Return the bottom-most position
                return offset;
            }
        }

        public XtraReportBase CurrentReport //The currently processed report
        {
            get => currentReport;
            set
            {
                if (currentReport != value)
                {
                    //Reset all values after changing the current report
                    currentReport = value;
                    currentBand = null;
                    CurrentControl = null;
                }
            } 
        }
        public Band CurrentBand //The currently processed band
        {
            get => currentBand;
            set
            {
                if (currentBand != value)
                {
                    //Reset the control value after changing the current band
                    currentBand = value;
                    CurrentControl = null;
                }
            }
        }
        public XRControl CurrentControl { get; set; } //The currently processed control
    }
}
