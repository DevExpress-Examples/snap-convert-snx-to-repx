using DevExpress.Snap;
using DevExpress.Snap.Core.API;
using DevExpress.XtraReports.UI;
using DevExpress.XtraRichEdit.API.Layout;
using DevExpress.XtraRichEdit.API.Native;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace SnxToRepx.Converter
{
    /// <summary>
    /// A delegate used to pass an action to nested processors
    /// </summary>
    /// <param name="rows">Layout rows with text content</param>
    /// <param name="tables">Layout tables</param>
    delegate void CellContentProcessorDelegate(LayoutRowCollection rows, LayoutTableCollection tables);

    /// <summary>
    /// This class processes the layout of the current document
    /// </summary>
    class SnapLayoutProcessor
    {
        SnapDocumentServer server; //A SnapDocumentServer instance containing the current document
        SnapListCollection snapLists, processedSnapLists; //Collections of SnapLists containing all and processed lists
        ReportGenerator generator; //A ReportGenerator containing the currently constructed XtraReport
        public SnapLayoutProcessor(ReportGenerator generator, SnapDocumentServer server)
        {
            this.server = server;
            this.generator = generator;
        }

        /// <summary>
        /// Iterates through pages and processes their content
        /// </summary>
        public void ProcessDocument()
        {
            //IMPORTANT! In Snap, you can create a multi-column layout with completely different
            //document elements. XtraReports do not support this layout, so a single-column report is generated.

            //Disable history to avoid potential document changing conflicts
            server.Options.DocumentCapabilities.Undo = DevExpress.XtraRichEdit.DocumentCapability.Disabled;
            
            //Create and initialize collections
            snapLists = new SnapListCollection(server.Document);
            snapLists.PrepareCollection();

            processedSnapLists = new SnapListCollection(server.Document);

            //Display field codes
            //This may help to distinguish between "template" and "result" fields
            server.Document.BeginUpdate();
            for (int i = 0; i < server.Document.Fields.Count; i++)
                server.Document.Fields[i].ShowCodes = true;
            server.Document.EndUpdate();

            //Iterate through pages
            for (int i = 0; i < server.DocumentLayout.GetPageCount(); i++)
            {
                LayoutPage page = server.DocumentLayout.GetPage(i);
                //Iterate through page areas
                foreach (LayoutPageArea pageArea in page.PageAreas)
                    //Iterate through columns
                    foreach (LayoutColumn column in pageArea.Columns)
                        //Process rows and tables in a column
                        ProcessLayoutElements(column.Rows, column.Tables);
            }
        }

        /// <summary>
        /// Processes tables and rows
        /// </summary>
        /// <param name="rows">A collection of LayoutRow elements</param>
        /// <param name="tables">A collection of tables</param>
        internal void ProcessLayoutElements(LayoutRowCollection rows, LayoutTableCollection tables)
        {
            //TableCells with both text and nested tables/charts are currently not supported. Only the nested table/chart is displayed
            //IMPORTANT! XRTableCell and XRLabel content is built using the markup and expressions. However, if they have nested controls, neither markup nor
            //expressions are in effect. Thus, only nested controls are visible in this case. In Snap, you can place any content in a table cell
            //(including multiple nesting table levels). There is no generic solution to automatically convert this content to XtraReports.

            //Create a collection of layout elements and add rows and tables to this collection in an appropriate order (after sorting)
            List<RangedLayoutElement> elements = new List<RangedLayoutElement>();
            elements.AddRange(rows);
            elements.AddRange(tables);
            elements.Sort(new RangedElementComparer());

            //If the collection is not empty
            if (elements.Count > 0)
            {      
                //Create a collection for text rows
                List<LayoutRow> paragraphRows = new List<LayoutRow>();

                //Iterate through elements
                for (int i = 0; i < elements.Count; i++)
                {
                    RangedLayoutElement element = elements[i];
                    //Check if an element contains SnapList
                    SnapList list = snapLists.GetSnapListByPosition(element.Range.Start);
                    if (list != null)
                        if (!processedSnapLists.ContainsField(list.Field))
                        {
                            //If a list is found and is not yet processed, process it
                            ProcessSnapList(list);
                        }
                        else
                            //Otherwise, continue execution
                            continue;
                    else
                    {
                        //If the current element is a table
                        if (element is LayoutTable)
                        {
                            //Create a TableProcessor instance and process the element
                            TableProcessor tableProcessor = new TableProcessor(generator, server, (LayoutTable)element, ProcessLayoutElements);
                            tableProcessor.ProcessTable();
                        }                            
                        else
                        //If the current element is a row
                            if (element is LayoutRow)
                        {
                            //Process it
                            LayoutRow row = (LayoutRow)element;
                            ProcessLayoutRow(elements, row, paragraphRows, i);                            
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Trims empty rows (containing empty paragraphs) from the beginning and from the end of a list
        /// </summary>
        /// <param name="rows">LayoutRows to process</param>
        private void CorrectParagraphRows(List<LayoutRow> rows)
        {
            List<LayoutRow> rowsToRemove = new List<LayoutRow>();

            //Iterate through rows in the forward direction
            for (int i = 0; i < rows.Count; i++)
                if (rows[i].Range.Length <= 1)
                    //If a row is empty, add it to the collection of rows to be removed
                    rowsToRemove.Add(rows[i]);
                else
                    //If a non-empty row is found, stop the process
                    break;

            //Remove empty rows
            foreach (LayoutRow row in rowsToRemove)
                rows.Remove(row);

            rowsToRemove.Clear();

            //Iterate through rows in the forward direction
            for (int i = rows.Count - 1; i >= 0; i--)
                if (rows[i].Range.Length <= 1)
                    rowsToRemove.Add(rows[i]);
                else
                    break;

            //Remove empty rows
            foreach (LayoutRow row in rowsToRemove)
                rows.Remove(row);
        }

        /// <summary>
        /// Processes the Snap List
        /// </summary>
        /// <param name="list">The source Snap list</param>
        private void ProcessSnapList(SnapList list)
        {
            //Get the currently processed report
            XtraReportBase report = generator.CurrentReport;

            //Create SnapListProcessor and process the list
            SnapListProcessor listProcessor = new SnapListProcessor(generator, server, list);
            listProcessor.ProcessSnapList();

            //Add this list to the collection of processed Snap lists
            processedSnapLists.Add(list.Field);

            //Obtain the current Band and XRControl
            Band band = generator.CurrentBand;
            XRControl control = generator.CurrentControl;

            //Restore the currently processed report, Band, and XRControl
            generator.CurrentReport = report;
            generator.CurrentBand = band;
            generator.CurrentControl = control;
        }

        /// <summary>
        /// Processes a single LayoutRow instance
        /// </summary>
        /// <param name="elements">A collection of layout elements</param>
        /// <param name="row">The current row</param>
        /// <param name="paragraphRows">A collection to store text rows</param>
        /// <param name="index">An index of the current row in the element collection</param>
        private void ProcessLayoutRow(List<RangedLayoutElement> elements, LayoutRow row, List<LayoutRow> paragraphRows, int index)
        {
            SnapChart chart = null;
            //Get a row range
            DocumentRange rowRange = server.Document.CreateRange(row.Range.Start, row.Range.Length);
            //Iterate through fields to check if they belong to this range
            for (int j = 0; j < server.Document.Fields.Count; j++)
            {
                Field field = server.Document.Fields[j];
                //If a field starts within the range, check the field type
                if (field.Range.Start.ToInt() >= row.Range.Start && field.Range.Start.ToInt() < row.Range.Start + row.Range.Length)
                {
                    SnapEntity entity = server.Document.ParseField(field);
                    //If it is a chart, store it in a variable
                    if (entity is SnapChart)
                    {
                        chart = ((SnapChart)entity);
                        break;
                    }
                }
            }
            //If no charts found, add this row to the text row collection
            if (chart == null)
                paragraphRows.Add(row);

            //If this is the last element, or the subsequent element is not a row, or if the next element contains a snap list
            if (index == elements.Count - 1 || !(elements[index + 1] is LayoutRow) || snapLists.GetSnapListByPosition(elements[index + 1].Range.Start) != null || chart != null)
            {
                //Remove empty rows
                CorrectParagraphRows(paragraphRows);
                if (paragraphRows.Count > 0)
                {
                    //Create a ParagraphProcessor instance and process rows
                    ParagraphProcessor paragraphProcessor = new ParagraphProcessor(generator, server, paragraphRows);
                    paragraphProcessor.ProcessParagraphs();
                    //Clear the collection of processed rows
                    paragraphRows.Clear();
                }
            }

            //If the chart was found
            if (chart != null)
            {
                //Process the chart
                ProcessChart(chart);
            }
        }

        /// <summary>
        /// Processes the chart element
        /// </summary>
        /// <param name="chart">The source chart</param>
        private void ProcessChart(SnapChart chart)
        {
            //Enable runtime customization
            chart.BeginUpdate();
            //Get the series collection
            DevExpress.XtraCharts.SeriesCollection series = chart.Series;
            //The only way to access the inner DevExpress.XtraCharts.Native.Chart object is to use System.Reflection
            //We use this object to copy all chart properties at once
            System.Reflection.PropertyInfo pi = typeof(DevExpress.XtraCharts.SeriesCollection).GetProperty("Chart", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            DevExpress.Snap.Core.Native.SNChart innerChart = pi?.GetValue(series) as DevExpress.Snap.Core.Native.SNChart;

            if (innerChart != null)
            {
                //Get the parent control
                XRControl parentControl = generator.CurrentControl;
                if (parentControl == null)
                    parentControl = generator.CurrentBand;

                //Create an XRChart instance and add it to the parent container
                XRChart reportChart = new XRChart();
                parentControl.Controls.Add(reportChart);

                //Get the chart size in measurement units used in XtraReports
                Size chartSize = XRConvert.Convert(chart.Size, DevExpress.XtraPrinting.GraphicsDpi.Pixel, DevExpress.XtraPrinting.GraphicsDpi.HundredthsOfAnInch);
                //Set boundaries
                reportChart.BoundsF = new RectangleF(new PointF(0, ReportGenerator.GetVerticalOffset(parentControl)), chartSize);

                //Assign properties from the source chart
                ((DevExpress.XtraCharts.Native.IChartContainer)reportChart).Chart.Assign(innerChart);
            }
            //Unlock the snap entity
            chart.EndUpdate();
        }
    }
}
