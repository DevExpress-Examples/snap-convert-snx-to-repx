using DevExpress.Data;
using DevExpress.XtraPrinting;
using DevExpress.XtraReports.UI;
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
    /// This class converts RichEdit border styles to XtraReport border styles
    /// </summary>
    static class BorderLineStyleConverter
    {
        //DOCS: Only a limited set of border styles is available in reports

        /// <summary>
        /// This method converts DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle to DevExpress.XtraPrinting.BorderDashStyle
        /// </summary>
        /// <param name="inputStyle">The source DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle</param>
        /// <returns>The resulting DevExpress.XtraPrinting.BorderDashStyle</returns>
        public static BorderDashStyle Convert(TableBorderLineStyle inputStyle)
        {
            switch (inputStyle)
            {
                case TableBorderLineStyle.Dashed:
                    return BorderDashStyle.Dash;
                case TableBorderLineStyle.DotDash:
                    return BorderDashStyle.DashDot;
                case TableBorderLineStyle.DotDotDash:
                    return BorderDashStyle.DashDotDot;
                case TableBorderLineStyle.Dotted:
                    return BorderDashStyle.Dot;
                case TableBorderLineStyle.Double:
                    return BorderDashStyle.Double;
                default:
                    return BorderDashStyle.Solid;
            }
        }
    }

    /// <summary>
    /// This class converts RichEdit alignment settings to XtraReport alignment
    /// </summary>
    static class AlignmentConverter
    {
        /// <summary>
        /// This method converts DevExpress.XtraRichEdit.API.Native.TableCellVerticalAlignment to DevExpress.XtraPrinting.TextAlignment
        /// </summary>
        /// <param name="vertAlign">The source DevExpress.XtraRichEdit.API.Native.TableCellVerticalAlignment</param>
        /// <returns>The resulting DevExpress.XtraPrinting.TextAlignment</returns>
        public static TextAlignment Convert(TableCellVerticalAlignment vertAlign)
        {
            switch (vertAlign)
            {
                case TableCellVerticalAlignment.Top:
                    return TextAlignment.TopLeft;
                case TableCellVerticalAlignment.Center:
                    return TextAlignment.MiddleLeft;
                case TableCellVerticalAlignment.Bottom:
                    return TextAlignment.BottomLeft;
                default:
                    return TextAlignment.TopLeft;
            }
        }

        /// <summary>
        /// This method converts DevExpress.XtraRichEdit.API.Native.ParagraphAlignment to the markup representation
        /// </summary>
        /// <param name="parAlign">The source DevExpress.XtraRichEdit.API.Native.ParagraphAlignment</param>
        /// <returns>The resulting markup tag with alignment</returns>
        public static string Convert(ParagraphAlignment parAlign)
        {
            switch (parAlign)
            {
                case ParagraphAlignment.Justify:
                    return "align=justify";
                case ParagraphAlignment.Right:
                    return "align=right";
                case ParagraphAlignment.Center:
                    return "align=center";
                case ParagraphAlignment.Left:
                    return "align=left";
                default:
                    return string.Empty;
            }
        }
    }

    /// <summary>
    /// This class converts the sorting order to XtraReports sorting order
    /// </summary>
    static class ColumnSortOrderConverter
    {
        /// <summary>
        /// This method converts DevExpress.Data.ColumnSortOrder to DevExpress.XtraReports.UI.XRColumnSortOrder
        /// </summary>
        /// <param name="sortOrder">The source DevExpress.Data.ColumnSortOrder</param>
        /// <returns>The resulting DevExpress.XtraReports.UI.XRColumnSortOrder</returns>
        public static XRColumnSortOrder Convert(ColumnSortOrder sortOrder)
        {
            switch (sortOrder)
            {
                case ColumnSortOrder.None:
                    return XRColumnSortOrder.None;
                case ColumnSortOrder.Descending:
                    return XRColumnSortOrder.Descending;
                case ColumnSortOrder.Ascending:
                default:
                    return XRColumnSortOrder.Ascending;                
            }
        }
    }

    /// <summary>
    /// This class converts Snap summary settings to XtraReports summary settings
    /// </summary>
    static class SummaryConverter
    {
        /// <summary>
        /// This method converts DevExpress.Data.SummaryItemType to the markup representation
        /// </summary>
        /// <param name="summaryFunc">The source DevExpress.XtraRichEdit.API.Native.ParagraphAlignment</param>
        /// <returns>The resulting markup tag with alignment</returns>
        public static string Convert(SummaryItemType summaryFunc) 
        {
            switch (summaryFunc)
            {
                case SummaryItemType.Average:
                    return "sumAvg";
                case SummaryItemType.Count:
                    return "sumCount";
                case SummaryItemType.Max:
                    return "sumMax";
                case SummaryItemType.Min:
                    return "sumMin";
                case SummaryItemType.Sum:
                case SummaryItemType.Custom:
                    return "sumSum";
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// This method converts DevExpress.Snap.Core.Fields.SummaryRunning to DevExpress.XtraReports.UI.SummaryRunning
        /// </summary>
        /// <param name="running">The source DevExpress.Snap.Core.Fields.SummaryRunning</param>
        /// <returns>The resulting DevExpress.XtraReports.UI.SummaryRunning</returns>
        public static SummaryRunning Convert(DevExpress.Snap.Core.Fields.SummaryRunning running)
        {
            switch (running)
            {
                case DevExpress.Snap.Core.Fields.SummaryRunning.Group:
                    return SummaryRunning.Group;
                case DevExpress.Snap.Core.Fields.SummaryRunning.Report:
                    return SummaryRunning.Report;
                default:
                    return SummaryRunning.None;
            }
        }
    }
}
