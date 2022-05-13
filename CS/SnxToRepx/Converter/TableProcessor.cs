using DevExpress.Snap;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrinting.Native;
using DevExpress.XtraReports.UI;
using DevExpress.XtraRichEdit.API.Layout;
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
    /// This class processes content in tables
    /// </summary>
    class TableProcessor : ContentProcessorBase
    {
        LayoutTable layoutTable; //The source table from the document layout
        CellContentProcessorDelegate cellAction; //An action that describes how to process cell content
        XRTable reportTable; //The target report table
        XRTableRow reportTableRow; //The target report table row

        public TableProcessor(ReportGenerator generator, SnapDocumentServer server, LayoutTable table, CellContentProcessorDelegate cellAction) : base(generator, server)
        {            
            this.layoutTable = table;
            this.cellAction = cellAction;
        }

        /// <summary>
        /// Processes the table
        /// </summary>
        public void ProcessTable()
        {
            //Store the current report, Band, and control
            XtraReportBase previousReport = Generator.CurrentReport;
            Band previousBand = Generator.CurrentBand;
            XRControl previousControl = Generator.CurrentControl;

            //Create a document position to retrieve the corresponding table from the document model
            DocumentPosition position = Server.Document.CreatePosition(layoutTable.Range.Start);
            TableCell firstTableCell = Server.Document.Tables.GetTableCell(position);
            Table modelTable = firstTableCell.Table;
            //Get the parent control to add a new table. If no control exists, use the current Band
            XRControl parentControl = previousControl;
            if (parentControl == null)
                parentControl = previousBand;
            //Calculate table boundaries relatively to its parent
            Rectangle bounds = layoutTable.GetRelativeBounds(layoutTable.Parent);
            //Correct table boundaries to fit into the report (in Snap, tables can start outside margins)
            Rectangle correctedBounds = new Rectangle(Math.Max(bounds.Left, 0),
                                                      0,
                                                      Math.Min(bounds.Width, layoutTable.Parent.Bounds.Width),
                                                      bounds.Height);
            //Convert boundaries to measurement units used in XtraReports
            correctedBounds = XRConvert.Convert(correctedBounds, GraphicsDpi.Twips, GraphicsDpi.HundredthsOfAnInch);
            //Get an offset and apply it
            int offset = ReportGenerator.GetVerticalOffset(parentControl);
            correctedBounds.Offset(0, offset);
            //Create the target XRTable, add it to the parent, and specify its boundaries.
            //IMPORTANT! You need to add a control first to properly apply measurement units to it. Otherwise, you
            //can get unexpected results
            reportTable = new XRTable();
            parentControl.Controls.Add(reportTable);
            reportTable.BoundsF = correctedBounds;
            //Process the table
            ProcessTableContent(modelTable);
            //Restore the previous control if the report is not changed
            //(it may change if the table contains a nested SnapList in a table cell)
            if (previousReport == Generator.CurrentReport && previousBand == Generator.CurrentBand)
                Generator.CurrentControl = previousControl;
        }

        /// <summary>
        /// Processes table content
        /// </summary>
        /// <param name="table">The source model table</param>
        private void ProcessTableContent(Table table)
        {
            //Lock updates and calculations in the report table.
            //This is required to avoid resizing each XRTableCell during the add operation
            reportTable.BeginInit();
            //Process rows
            ProcessRows(table);
            //Unlock updates
            reportTable.EndInit();
            //Process the table to merge cells and perform corrections
            PostProcessTable();
        }

        /// <summary>
        /// Applies final touches to the table
        /// </summary>
        private void PostProcessTable()
        {
            //Merge cells vertically
            ApplyMerging();
            //Correct nested control sizes
            CorrectTable();
        }

        /// <summary>
        /// Corrects sizes of nested controls
        /// </summary>
        private void CorrectTable()
        {
            //Iterate through rows
            for (int i = 0; i < reportTable.Rows.Count; i++)
            {
                //Iterate through cells in a row
                for (int j = 0; j < reportTable.Rows[i].Cells.Count; j++)
                {
                    XRTableCell currentCell = reportTable.Rows[i].Cells[j];

                    //If a row contains only a single cell, disable shrinking for this cell
                    //Such cells may be used as gaps between different content parts
                    if (reportTable.Rows[i].Cells.Count == 1)
                        currentCell.CanShrink = false;
                    //If a cell contains a single control, resize this control to fit into cell boundaries
                    if (currentCell.Controls.Count == 1)
                    {
                        XRControl nestedControl = currentCell.Controls[0];
                        CorrectControlSize(nestedControl);
                    }
                }
            }
        }

        /// <summary>
        /// Make control fit its parent boundaries
        /// </summary>
        /// <param name="control">A control to process</param>
        private void CorrectControlSize(XRControl control)
        {
            //The BestSizeEstimator calculates the best-fit size for the control
            control.BoundsF = BestSizeEstimator.GetBoundsToFitContainer(control);
        }

        /// <summary>
        /// Gets the layout row for a model row
        /// </summary>
        /// <param name="row">The source model row</param>
        /// <returns>The corresponding layout row</returns>
        private LayoutTableRow GetLayoutTableRow(DevExpress.XtraRichEdit.API.Native.TableRow row)
        {
            //Iterate through cells in a row
            for (int i = 0; i < row.Cells.Count; i++)
            {
                TableCell cell = row.Cells[i];

                //IMPORTANT! Sometimes it is impossible to get the LayoutTableRow directly using the TableRow range.
                //For example, if the first cell of the layout row is in the middle/end of a merged cell, the LayoutRow is Null.
                //That is why we need to check every cell to get the layout row
                //Try to access the layout row based on the cell range
                LayoutTableRow layoutRow = Server.DocumentLayout.GetElement<LayoutTableRow>(cell.ContentRange.Start);
                if (layoutRow != null)
                    return layoutRow;
            }
            //If no row is found, return Null
            return null;
        }

        /// <summary>
        /// Processes table rows
        /// </summary>
        /// <param name="table">The source model table</param>
        private void ProcessRows(Table table)
        {
            //Iterate through rows
            for (int i = 0; i < table.Rows.Count; i++)
            {
                //Get the corresponding layout row
                DevExpress.XtraRichEdit.API.Native.TableRow row = table.Rows[i];
                LayoutTableRow layoutRow = GetLayoutTableRow(row);                

                //Create an XRTableRow and add it to the resulting table
                reportTableRow = new XRTableRow();
                reportTable.Rows.Add(reportTableRow);
                //Set the row height
                if (layoutRow != null)
                    reportTableRow.HeightF = XRConvert.Convert(layoutRow.Bounds.Height, DevExpress.XtraPrinting.GraphicsDpi.Twips, DevExpress.XtraPrinting.GraphicsDpi.HundredthsOfAnInch);
                //Process cells in this row
                ProcessCells(row, reportTableRow);
            }
        }

        /// <summary>
        /// Processes cells in a row
        /// </summary>
        /// <param name="row">The source row</param>
        /// <param name="reportTableRow">The target row</param>
        private void ProcessCells(DevExpress.XtraRichEdit.API.Native.TableRow row, XRTableRow reportTableRow)
        {
            //Iterate through cells in the source row
            for (int i = 0; i < row.Cells.Count; i++)
            {
                TableCell cell = row.Cells[i];
                //Create the corresponding XRTableCell and add it to the row
                XRTableCell reportTableCell = new XRTableCell();
                reportTableRow.Cells.Add(reportTableCell);
                //Process cell settings
                ProcessTableCell(cell, reportTableCell);
                //Store the current band for future use
                Band currentBand = Generator.CurrentBand;

                //If the delegate method exists
                if (cellAction != null)
                {
                    //Specify the current control (it will be used as a container for inner content)
                    Generator.CurrentControl = reportTableCell;
                    //Get the corresponding layout cell
                    LayoutTableCell layoutCell = Server.DocumentLayout.GetElement<LayoutTableCell>(cell.ContentRange.Start);
                    //If this cell is not a part of the merged cell, process its content
                    if (layoutCell != null)
                        cellAction(layoutCell.Rows, layoutCell.NestedTables);
                }

                //If the current control is changed during processing content
                if (Generator.CurrentControl != reportTableCell)
                    //If the Band is the same, restore the current control
                    if (currentBand == Generator.CurrentBand)
                        Generator.CurrentControl = reportTableCell;
                    else
                    {
                        //IMPORTANT!If the Band is changed, this means that the current table cell contains an inner SnapList.
                        //However, there might be subsequent rows and cells, which should be processed as well.
                        //Thus, we create a new table, copy table settings and continue processing.

                        //If the cell has no inner controls or content, remove it
                        if (reportTableCell.Controls.Count == 0 && string.IsNullOrEmpty(reportTableCell.Text) && reportTableCell.ExpressionBindings.Count == 0)
                            reportTableRow.Cells.Remove(reportTableCell);
                        //Get the current row height
                        float rowHeight = reportTableRow.HeightF;
                        //If this row has no cells, remove it and correct the resulting table height
                        if (reportTableRow.Cells.Count == 0)
                        {                            
                            XRTable parentTable = reportTableRow.Table;
                            parentTable.Rows.Remove(reportTableRow);
                            parentTable.HeightF -= rowHeight;
                            parentTable.Band.HeightF = 0;
                        }
                        //If this is not the last row or cell:
                        if (i < row.Cells.Count - 1 || !row.IsLast)
                        {
                            //Get table boundaries
                            RectangleF bounds = reportTable.BoundsF;
                            //Calculate the remaining height
                            float totalHeight = 0;
                            for (int j = 0; j < reportTable.Rows.Count; j++)
                                totalHeight += reportTable.Rows[j].HeightF;
                            //Specify the final height
                            reportTable.HeightF = totalHeight;
                            //Unlock the table and process it
                            reportTable.EndInit();
                            PostProcessTable();

                            //Get the current band and offset
                            Band band = Generator.CurrentBand;
                            int offset = ReportGenerator.GetVerticalOffset(band);
                            //Create a new XRTable in the target Band and copy settings
                            reportTable = new XRTable();
                            band.Controls.Add(reportTable);
                            reportTable.BoundsF = new RectangleF(bounds.X, offset, bounds.Width, bounds.Height - totalHeight);
                            //If this is not the last cell in the current row, create a row and specify its height
                            if (i < row.Cells.Count - 1)
                            {
                                reportTableRow = new XRTableRow();
                                reportTable.Rows.Add(reportTableRow);
                                reportTableRow.HeightF = rowHeight;
                            }
                        }
                    }                    
            }
        }

        /// <summary>
        /// Specifies base settings and stores the merge state in the current cell
        /// </summary>
        /// <param name="cell">The source cell</param>
        /// <param name="reportTableCell">The target cell</param>
        private void ProcessTableCell(TableCell cell, XRTableCell reportTableCell)
        {
            float cellWidth = 0;
            //Store the vertical merge state
            reportTableCell.Tag = cell.VerticalMerging;
            //Check the merge state
            switch (cell.VerticalMerging)
            {
                //If this cell is in the middle/end of the merged cell, we cannot get the corresponding layout cell.
                //Instead, we get report table cells in this column to properly calculate the cell width.
                case VerticalMergingState.Continue:
                    //Get cells in a column
                    List<XRTableCell> cells = GetColumnCells(reportTableCell);
                    //Get the "upper" cell
                    XRTableCell prevCell = cells[cells.IndexOf(reportTableCell) - 1];
                    //Specify the width
                    cellWidth = prevCell.WidthF;

                    break;
                //In other cases, just take the cell width from the layout cell
                case VerticalMergingState.Restart:
                case VerticalMergingState.None:
                default:
                    LayoutTableCell layoutCell = Server.DocumentLayout.GetElement<LayoutTableCell>(cell.ContentRange.Start);
                    cellWidth = XRConvert.Convert(layoutCell.Bounds.Width, DevExpress.XtraPrinting.GraphicsDpi.Twips, DevExpress.XtraPrinting.GraphicsDpi.HundredthsOfAnInch);
                    break;
            }
            //Set the cell width, enable shrinking based on content, and allow HTML-style markup to format content
            reportTableCell.WidthF = cellWidth;
            reportTableCell.CanShrink = true;
            reportTableCell.AllowMarkupText = true;
            //Copy appearance settings
            CopyCellAppearance(cell, reportTableCell);
        }

        /// <summary>
        /// Applies merging to table cells
        /// </summary>
        private void ApplyMerging()
        {
            //Lock updates in the table
            reportTable.BeginInit();
            //Iterate through rows and cells
            foreach (XRTableRow row in reportTable.Rows)
            {
                foreach (XRTableCell cell in row.Cells)
                {
                    //Get the vertical merge state
                    if (cell.Tag is VerticalMergingState)
                    {
                        VerticalMergingState mergingState = (VerticalMergingState)cell.Tag;
                        //If the cell is in the beginning of the merged cell
                        if (mergingState == VerticalMergingState.Restart)
                        {
                            //Get all cells in a column
                            List<XRTableCell> cells = GetColumnCells(cell);
                            //Store the index where the cell starts
                            int startCellIndex = cells.IndexOf(cell);
                            int endCellIndex = startCellIndex;

                            //Iterate through cells in the column starting with the current cell
                            for (int i = startCellIndex + 1; i < cells.Count; i++)
                            {
                                //If all merged cells are processed, exit the loop
                                if (cells[i].Tag is VerticalMergingState && ((VerticalMergingState)cells[i].Tag) != VerticalMergingState.Continue)
                                    break;
                                
                                //Correct the width
                                cells[i].WidthF = cell.WidthF;                                
                                //Store the new "end" cell index
                                endCellIndex++;
                            }
                            //RowSpan does not have any effect if a cell contains nested controls (nested tables, charts, etc.)
                            //Calculate the RowSpan value based on indexes obtained in previous steps
                            cell.RowSpan = endCellIndex - startCellIndex + 1;                            
                        }
                    }
                    //Reset the cell tag
                    cell.Tag = null;
                }
            }
            //Unlock the table
            reportTable.EndInit();
        }

        /// <summary>
        /// Copies basic appearance settings of a table cell
        /// </summary>
        /// <param name="tableCell">The source table cell</param>
        /// <param name="reportTableCell">The target report table cell</param>
        private void CopyCellAppearance(TableCell tableCell, XRTableCell reportTableCell)
        {
            //Set the background color
            reportTableCell.BackColor = tableCell.BackgroundColor;
            //Set borders
            CopyBorders(tableCell, reportTableCell);
            //Set word wrapping
            reportTableCell.WordWrap = tableCell.WordWrap;
            //Specify alignment
            SetAlignment(tableCell, reportTableCell);
        }

        private void CopyBorders(TableCell tableCell, XRTableCell reportTableCell)
        {
            //Table borders of different styles/colors are not supported.
            //BorderStyle/BorderColor are obtained from the last visible border. Borders of different Colors/Styles are ignored.
            reportTableCell.Borders = BorderSide.None;

            //Specify borders individually
            SetCellBorder(reportTableCell, tableCell.Borders.Top, BorderSide.Top);
            SetCellBorder(reportTableCell, tableCell.Borders.Left, BorderSide.Left);
            SetCellBorder(reportTableCell, tableCell.Borders.Right, BorderSide.Right);
            SetCellBorder(reportTableCell, tableCell.Borders.Bottom, BorderSide.Bottom);                     
        }

        /// <summary>
        /// Applies settings to the corresponding border
        /// </summary>
        /// <param name="reportTableCell">The target report table cell</param>
        /// <param name="cellBorder">The source cell border</param>
        /// <param name="side">The target border side</param>
        private void SetCellBorder(XRTableCell reportTableCell, TableCellBorder cellBorder, BorderSide side)
        {
            //Check if the border is visible
            if (IsBorderVisible(cellBorder))
            {
                //Add a border
                reportTableCell.Borders |= side;
                //Set the border style
                SetCellBorderStyle(reportTableCell, cellBorder);
            }
        }

        /// <summary>
        /// Sets the border style
        /// </summary>
        /// <param name="reportTableCell">The target report table cell</param>
        /// <param name="cellBorder">The source cell border</param>
        private void SetCellBorderStyle(XRTableCell reportTableCell, TableCellBorder cellBorder)
        {
            //Set the border color
            reportTableCell.BorderColor = cellBorder.LineColor;
            //Set the border line style
            reportTableCell.BorderDashStyle = BorderLineStyleConverter.Convert(cellBorder.LineStyle);
        }

        /// <summary>
        /// Checks whether the border is visible
        /// </summary>
        /// <param name="border">The source cell border</param>
        /// <returns>True if border is visible. Otherwise, false</returns>
        private bool IsBorderVisible(TableCellBorder border)
        {
            //Check if color is empty or the line style is not specified
            return border.LineColor != Color.Empty && border.LineStyle != TableBorderLineStyle.None && border.LineStyle != TableBorderLineStyle.Nil;
        }

        /// <summary>
        /// Specifies cell alignment
        /// </summary>
        /// <param name="cell">The source cell</param>
        /// <param name="reportTableCell">The target cell</param>
        private void SetAlignment(TableCell cell, XRTableCell reportTableCell)
        {
            //Assign VerticalAlignment to TextAlignment. Horizontal alignment should always be Left because of the AllowHtmlMarkup property
            reportTableCell.TextAlignment = AlignmentConverter.Convert(cell.VerticalAlignment);
            //Get paddings
            PaddingInfo paddingInfo = GetPaddingInfo(cell.ContentRange);
            //Set paddings in report measurement units
            paddingInfo.Left = (int)XRConvert.Convert(cell.LeftPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);
            paddingInfo.Right = (int)XRConvert.Convert(cell.RightPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch);
            paddingInfo.Top = Math.Max((int)XRConvert.Convert(cell.TopPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch), paddingInfo.Top);
            paddingInfo.Bottom = Math.Max((int)XRConvert.Convert(cell.BottomPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch), paddingInfo.Bottom);
            //Specify paddings in the report table cell
            reportTableCell.Padding = paddingInfo;
        }

        /// <summary>
        /// Gets all cells in the column
        /// </summary>
        /// <param name="baseCell">The report cell to identify a column</param>
        /// <returns>A collection of table cells in a column</returns>
        private List<XRTableCell> GetColumnCells(XRTableCell baseCell)
        {
            //Get the target table
            XRTable reportTable = baseCell.Row.Table;
            //Create the output collection
            List<XRTableCell> cells = new List<XRTableCell>();
            //Iterate through rows and cells
            foreach (XRTableRow row in reportTable.Rows)
            {
                foreach (XRTableCell cell in row.Cells)
                    //Check if cells start in the same horizontal position
                    if (FloatsComparer.Default.FirstEqualsSecond(cell.LeftF, baseCell.LeftF))
                        //If so, add a cell to the collection
                        cells.Add(cell);
            }
            //Return the collection
            return cells;
        }
    }
}
