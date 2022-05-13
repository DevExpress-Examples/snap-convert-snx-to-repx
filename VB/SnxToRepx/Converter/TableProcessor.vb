Imports DevExpress.Snap
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraPrinting.Native
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraRichEdit.API.Layout
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Drawing

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This class processes content in tables
    ''' </summary>
    Friend Class TableProcessor
        Inherits ContentProcessorBase

        Private layoutTable As LayoutTable 'The source table from the document layout

        Private cellAction As CellContentProcessorDelegate 'An action that describes how to process cell content

        Private reportTable As XRTable 'The target report table

        Private reportTableRow As XRTableRow 'The target report table row

        Public Sub New(ByVal generator As ReportGenerator, ByVal server As SnapDocumentServer, ByVal table As LayoutTable, ByVal cellAction As CellContentProcessorDelegate)
            MyBase.New(generator, server)
            layoutTable = table
            Me.cellAction = cellAction
        End Sub

        ''' <summary>
        ''' Processes the table
        ''' </summary>
        Public Sub ProcessTable()
            'Store the current report, Band, and control
            Dim previousReport As XtraReportBase = Generator.CurrentReport
            Dim previousBand As Band = Generator.CurrentBand
            Dim previousControl As XRControl = Generator.CurrentControl
            'Create a document position to retrieve the corresponding table from the document model
            Dim position As DocumentPosition = Server.Document.CreatePosition(layoutTable.Range.Start)
            Dim firstTableCell As TableCell = Server.Document.Tables.GetTableCell(position)
            Dim modelTable As Table = firstTableCell.Table
            'Get the parent control to add a new table. If no control exists, use the current Band
            Dim parentControl As XRControl = previousControl
            If parentControl Is Nothing Then parentControl = previousBand
            'Calculate table boundaries relatively to its parent
            Dim bounds As Rectangle = layoutTable.GetRelativeBounds(layoutTable.Parent)
            'Correct table boundaries to fit into the report (in Snap, tables can start outside margins)
            Dim correctedBounds As Rectangle = New Rectangle(Math.Max(bounds.Left, 0), 0, Math.Min(bounds.Width, layoutTable.Parent.Bounds.Width), bounds.Height)
            'Convert boundaries to measurement units used in XtraReports
            correctedBounds = GraphicsUnitConverter.Convert(correctedBounds, GraphicsDpi.Twips, GraphicsDpi.HundredthsOfAnInch)
            'Get an offset and apply it
            Dim offset As Integer = ReportGenerator.GetVerticalOffset(parentControl)
            correctedBounds.Offset(0, offset)
            'Create the target XRTable, add it to the parent, and specify its boundaries.
            'IMPORTANT! You need to add a control first to properly apply measurement units to it. Otherwise, you
            'can get unexpected results
            reportTable = New XRTable()
            parentControl.Controls.Add(reportTable)
            reportTable.BoundsF = correctedBounds
            'Process the table
            ProcessTableContent(modelTable)
            'Restore the previous control if the report is not changed
            '(it may change if the table contains a nested SnapList in a table cell)
            If previousReport Is Generator.CurrentReport AndAlso previousBand Is Generator.CurrentBand Then Generator.CurrentControl = previousControl
        End Sub

        ''' <summary>
        ''' Processes table content
        ''' </summary>
        ''' <param name="table">The source model table</param>
        Private Sub ProcessTableContent(ByVal table As Table)
            'Lock updates and calculations in the report table.
            'This is required to avoid resizing each XRTableCell during the add operation
            reportTable.BeginInit()
            'Process rows
            ProcessRows(table)
            'Unlock updates
            reportTable.EndInit()
            'Process the table to merge cells and perform corrections
            PostProcessTable()
        End Sub

        ''' <summary>
        ''' Applies final touches to the table
        ''' </summary>
        Private Sub PostProcessTable()
            'Merge cells vertically
            ApplyMerging()
            'Correct nested control sizes
            CorrectTable()
        End Sub

        ''' <summary>
        ''' Corrects sizes of nested controls
        ''' </summary>
        Private Sub CorrectTable()
            'Iterate through rows
            For i As Integer = 0 To reportTable.Rows.Count - 1
                'Iterate through cells in a row
                For j As Integer = 0 To reportTable.Rows(i).Cells.Count - 1
                    Dim currentCell As XRTableCell = reportTable.Rows(i).Cells(j)
                    'If a row contains only a single cell, disable shrinking for this cell
                    'Such cells may be used as gaps between different content parts
                    If reportTable.Rows(i).Cells.Count = 1 Then currentCell.CanShrink = False
                    'If a cell contains a single control, resize this control to fit into cell boundaries
                    If currentCell.Controls.Count = 1 Then
                        Dim nestedControl As XRControl = currentCell.Controls(0)
                        CorrectControlSize(nestedControl)
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' Make control fit its parent boundaries
        ''' </summary>
        ''' <param name="control">A control to process</param>
        Private Sub CorrectControlSize(ByVal control As XRControl)
            'The BestSizeEstimator calculates the best-fit size for the control
            control.BoundsF = BestSizeEstimator.GetBoundsToFitContainer(control)
        End Sub

        ''' <summary>
        ''' Gets the layout row for a model row
        ''' </summary>
        ''' <param name="row">The source model row</param>
        ''' <returns>The corresponding layout row</returns>
        Private Function GetLayoutTableRow(ByVal row As DevExpress.XtraRichEdit.API.Native.TableRow) As LayoutTableRow
            'Iterate through cells in a row
            For i As Integer = 0 To row.Cells.Count - 1
                Dim cell As TableCell = row.Cells(i)
                'IMPORTANT! Sometimes it is impossible to get the LayoutTableRow directly using the TableRow range.
                'For example, if the first cell of the layout row is in the middle/end of a merged cell, the LayoutRow is Null.
                'That is why we need to check every cell to get the layout row
                'Try to access the layout row based on the cell range
                Dim layoutRow As LayoutTableRow = Server.DocumentLayout.GetElement(Of LayoutTableRow)(cell.ContentRange.Start)
                If layoutRow IsNot Nothing Then Return layoutRow
            Next

            'If no row is found, return Null
            Return Nothing
        End Function

        ''' <summary>
        ''' Processes table rows
        ''' </summary>
        ''' <param name="table">The source model table</param>
        Private Sub ProcessRows(ByVal table As Table)
            'Iterate through rows
            For i As Integer = 0 To table.Rows.Count - 1
                'Get the corresponding layout row
                Dim row As DevExpress.XtraRichEdit.API.Native.TableRow = table.Rows(i)
                Dim layoutRow As LayoutTableRow = GetLayoutTableRow(row)
                'Create an XRTableRow and add it to the resulting table
                reportTableRow = New XRTableRow()
                reportTable.Rows.Add(reportTableRow)
                'Set the row height
                If layoutRow IsNot Nothing Then reportTableRow.HeightF = GraphicsUnitConverter.Convert(layoutRow.Bounds.Height, GraphicsDpi.Twips, GraphicsDpi.HundredthsOfAnInch)
                'Process cells in this row
                ProcessCells(row, reportTableRow)
            Next
        End Sub

        ''' <summary>
        ''' Processes cells in a row
        ''' </summary>
        ''' <param name="row">The source row</param>
        ''' <param name="reportTableRow">The target row</param>
        Private Sub ProcessCells(ByVal row As DevExpress.XtraRichEdit.API.Native.TableRow, ByVal reportTableRow As XRTableRow)
            'Iterate through cells in the source row
            For i As Integer = 0 To row.Cells.Count - 1
                Dim cell As TableCell = row.Cells(i)
                'Create the corresponding XRTableCell and add it to the row
                Dim reportTableCell As XRTableCell = New XRTableCell()
                reportTableRow.Cells.Add(reportTableCell)
                'Process cell settings
                ProcessTableCell(cell, reportTableCell)
                'Store the current band for future use
                Dim currentBand As Band = Generator.CurrentBand
                'If the delegate method exists
                If cellAction IsNot Nothing Then
                    'Specify the current control (it will be used as a container for inner content)
                    Generator.CurrentControl = reportTableCell
                    'Get the corresponding layout cell
                    Dim layoutCell As LayoutTableCell = Server.DocumentLayout.GetElement(Of LayoutTableCell)(cell.ContentRange.Start)
                    'If this cell is not a part of the merged cell, process its content
                    If layoutCell IsNot Nothing Then cellAction(layoutCell.Rows, layoutCell.NestedTables)
                End If

                'If the current control is changed during processing content
                If Generator.CurrentControl IsNot reportTableCell Then
                    'If the Band is the same, restore the current control
                    If currentBand Is Generator.CurrentBand Then
                        Generator.CurrentControl = reportTableCell
                    Else
                        'IMPORTANT!If the Band is changed, this means that the current table cell contains an inner SnapList.
                        'However, there might be subsequent rows and cells, which should be processed as well.
                        'Thus, we create a new table, copy table settings and continue processing.
                        'If the cell has no inner controls or content, remove it
                        If reportTableCell.Controls.Count = 0 AndAlso String.IsNullOrEmpty(reportTableCell.Text) AndAlso reportTableCell.ExpressionBindings.Count = 0 Then reportTableRow.Cells.Remove(reportTableCell)
                        'Get the current row height
                        Dim rowHeight As Single = reportTableRow.HeightF
                        'If this row has no cells, remove it and correct the resulting table height
                        If reportTableRow.Cells.Count = 0 Then
                            Dim parentTable As XRTable = reportTableRow.Table
                            parentTable.Rows.Remove(reportTableRow)
                            parentTable.HeightF -= rowHeight
                            parentTable.Band.HeightF = 0
                        End If

                        'If this is not the last row or cell:
                        If i < row.Cells.Count - 1 OrElse Not row.IsLast Then
                            'Get table boundaries
                            Dim bounds As RectangleF = reportTable.BoundsF
                            'Calculate the remaining height
                            Dim totalHeight As Single = 0
                            For j As Integer = 0 To reportTable.Rows.Count - 1
                                totalHeight += reportTable.Rows(j).HeightF
                            Next

                            'Specify the final height
                            reportTable.HeightF = totalHeight
                            'Unlock the table and process it
                            reportTable.EndInit()
                            PostProcessTable()
                            'Get the current band and offset
                            Dim band As Band = Generator.CurrentBand
                            Dim offset As Integer = ReportGenerator.GetVerticalOffset(band)
                            'Create a new XRTable in the target Band and copy settings
                            reportTable = New XRTable()
                            band.Controls.Add(reportTable)
                            reportTable.BoundsF = New RectangleF(bounds.X, offset, bounds.Width, bounds.Height - totalHeight)
                            'If this is not the last cell in the current row, create a row and specify its height
                            If i < row.Cells.Count - 1 Then
                                reportTableRow = New XRTableRow()
                                reportTable.Rows.Add(reportTableRow)
                                reportTableRow.HeightF = rowHeight
                            End If
                        End If
                    End If
                End If
            Next
        End Sub

        ''' <summary>
        ''' Specifies base settings and stores the merge state in the current cell
        ''' </summary>
        ''' <param name="cell">The source cell</param>
        ''' <param name="reportTableCell">The target cell</param>
        Private Sub ProcessTableCell(ByVal cell As TableCell, ByVal reportTableCell As XRTableCell)
            Dim cellWidth As Single = 0
            'Store the vertical merge state
            reportTableCell.Tag = cell.VerticalMerging
            'Check the merge state
            Select Case cell.VerticalMerging
                'If this cell is in the middle/end of the merged cell, we cannot get the corresponding layout cell.
                'Instead, we get report table cells in this column to properly calculate the cell width.
                Case VerticalMergingState.Continue
                    'Get cells in a column
                    Dim cells As List(Of XRTableCell) = GetColumnCells(reportTableCell)
                    'Get the "upper" cell
                    Dim prevCell As XRTableCell = cells(cells.IndexOf(reportTableCell) - 1)
                    'Specify the width
                    'In other cases, just take the cell width from the layout cell
                    cellWidth = prevCell.WidthF
                Case Else
                    Dim layoutCell As LayoutTableCell = Server.DocumentLayout.GetElement(Of LayoutTableCell)(cell.ContentRange.Start)
                    cellWidth = GraphicsUnitConverter.Convert(layoutCell.Bounds.Width, GraphicsDpi.Twips, GraphicsDpi.HundredthsOfAnInch)
            End Select

            'Set the cell width, enable shrinking based on content, and allow HTML-style markup to format content
            reportTableCell.WidthF = cellWidth
            reportTableCell.CanShrink = True
            reportTableCell.AllowMarkupText = True
            'Copy appearance settings
            CopyCellAppearance(cell, reportTableCell)
        End Sub

        ''' <summary>
        ''' Applies merging to table cells
        ''' </summary>
        Private Sub ApplyMerging()
            'Lock updates in the table
            reportTable.BeginInit()
            'Iterate through rows and cells
            For Each row As XRTableRow In reportTable.Rows
                For Each cell As XRTableCell In row.Cells
                    'Get the vertical merge state
                    If TypeOf cell.Tag Is VerticalMergingState Then
                        Dim mergingState As VerticalMergingState = CType(cell.Tag, VerticalMergingState)
                        'If the cell is in the beginning of the merged cell
                        If mergingState = VerticalMergingState.Restart Then
                            'Get all cells in a column
                            Dim cells As List(Of XRTableCell) = GetColumnCells(cell)
                            'Store the index where the cell starts
                            Dim startCellIndex As Integer = cells.IndexOf(cell)
                            Dim endCellIndex As Integer = startCellIndex
                            'Iterate through cells in the column starting with the current cell
                            For i As Integer = startCellIndex + 1 To cells.Count - 1
                                'If all merged cells are processed, exit the loop
                                If TypeOf cells(i).Tag Is VerticalMergingState AndAlso CType(cells(i).Tag, VerticalMergingState) <> VerticalMergingState.Continue Then Exit For
                                'Correct the width
                                cells(i).WidthF = cell.WidthF
                                'Store the new "end" cell index
                                endCellIndex += 1
                            Next

                            'RowSpan does not have any effect if a cell contains nested controls (nested tables, charts, etc.)
                            'Calculate the RowSpan value based on indexes obtained in previous steps
                            cell.RowSpan = endCellIndex - startCellIndex + 1
                        End If
                    End If

                    'Reset the cell tag
                    cell.Tag = Nothing
                Next
            Next

            'Unlock the table
            reportTable.EndInit()
        End Sub

        ''' <summary>
        ''' Copies basic appearance settings of a table cell
        ''' </summary>
        ''' <param name="tableCell">The source table cell</param>
        ''' <param name="reportTableCell">The target report table cell</param>
        Private Sub CopyCellAppearance(ByVal tableCell As TableCell, ByVal reportTableCell As XRTableCell)
            'Set the background color
            reportTableCell.BackColor = tableCell.BackgroundColor
            'Set borders
            CopyBorders(tableCell, reportTableCell)
            'Set word wrapping
            reportTableCell.WordWrap = tableCell.WordWrap
            'Specify alignment
            SetAlignment(tableCell, reportTableCell)
        End Sub

        Private Sub CopyBorders(ByVal tableCell As TableCell, ByVal reportTableCell As XRTableCell)
            'Table borders of different styles/colors are not supported.
            'BorderStyle/BorderColor are obtained from the last visible border. Borders of different Colors/Styles are ignored.
            reportTableCell.Borders = BorderSide.None
            'Specify borders individually
            SetCellBorder(reportTableCell, tableCell.Borders.Top, BorderSide.Top)
            SetCellBorder(reportTableCell, tableCell.Borders.Left, BorderSide.Left)
            SetCellBorder(reportTableCell, tableCell.Borders.Right, BorderSide.Right)
            SetCellBorder(reportTableCell, tableCell.Borders.Bottom, BorderSide.Bottom)
        End Sub

        ''' <summary>
        ''' Applies settings to the corresponding border
        ''' </summary>
        ''' <param name="reportTableCell">The target report table cell</param>
        ''' <param name="cellBorder">The source cell border</param>
        ''' <param name="side">The target border side</param>
        Private Sub SetCellBorder(ByVal reportTableCell As XRTableCell, ByVal cellBorder As TableCellBorder, ByVal side As BorderSide)
            'Check if the border is visible
            If IsBorderVisible(cellBorder) Then
                'Add a border
                reportTableCell.Borders = reportTableCell.Borders Or side
                'Set the border style
                SetCellBorderStyle(reportTableCell, cellBorder)
            End If
        End Sub

        ''' <summary>
        ''' Sets the border style
        ''' </summary>
        ''' <param name="reportTableCell">The target report table cell</param>
        ''' <param name="cellBorder">The source cell border</param>
        Private Sub SetCellBorderStyle(ByVal reportTableCell As XRTableCell, ByVal cellBorder As TableCellBorder)
            'Set the border color
            reportTableCell.BorderColor = cellBorder.LineColor
            'Set the border line style
            reportTableCell.BorderDashStyle = BorderLineStyleConverter.Convert(cellBorder.LineStyle)
        End Sub

        ''' <summary>
        ''' Checks whether the border is visible
        ''' </summary>
        ''' <param name="border">The source cell border</param>
        ''' <returns>True if border is visible. Otherwise, false</returns>
        Private Function IsBorderVisible(ByVal border As TableCellBorder) As Boolean
            'Check if color is empty or the line style is not specified
            Return border.LineColor <> Color.Empty AndAlso border.LineStyle <> TableBorderLineStyle.None AndAlso border.LineStyle <> TableBorderLineStyle.Nil
        End Function

        ''' <summary>
        ''' Specifies cell alignment
        ''' </summary>
        ''' <param name="cell">The source cell</param>
        ''' <param name="reportTableCell">The target cell</param>
        Private Sub SetAlignment(ByVal cell As TableCell, ByVal reportTableCell As XRTableCell)
            'Assign VerticalAlignment to TextAlignment. Horizontal alignment should always be Left because of the AllowHtmlMarkup property
            reportTableCell.TextAlignment = AlignmentConverter.Convert(cell.VerticalAlignment)
            'Get paddings
            Dim paddingInfo As PaddingInfo = GetPaddingInfo(cell.ContentRange)
            'Set paddings in report measurement units
            paddingInfo.Left = CInt(GraphicsUnitConverter.Convert(cell.LeftPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            paddingInfo.Right = CInt(GraphicsUnitConverter.Convert(cell.RightPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            paddingInfo.Top = Math.Max(CInt(GraphicsUnitConverter.Convert(cell.TopPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch)), paddingInfo.Top)
            paddingInfo.Bottom = Math.Max(CInt(GraphicsUnitConverter.Convert(cell.BottomPadding, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch)), paddingInfo.Bottom)
            'Specify paddings in the report table cell
            reportTableCell.Padding = paddingInfo
        End Sub

        ''' <summary>
        ''' Gets all cells in the column
        ''' </summary>
        ''' <param name="baseCell">The report cell to identify a column</param>
        ''' <returns>A collection of table cells in a column</returns>
        Private Function GetColumnCells(ByVal baseCell As XRTableCell) As List(Of XRTableCell)
            'Get the target table
            Dim reportTable As XRTable = baseCell.Row.Table
            'Create the output collection
            Dim cells As List(Of XRTableCell) = New List(Of XRTableCell)()
            'Iterate through rows and cells
            For Each row As XRTableRow In reportTable.Rows
                For Each cell As XRTableCell In row.Cells
                    'Check if cells start in the same horizontal position
                    'If so, add a cell to the collection
                    If FloatsComparer.Default.FirstEqualsSecond(cell.LeftF, baseCell.LeftF) Then cells.Add(cell)
                Next
            Next

            'Return the collection
            Return cells
        End Function
    End Class
End Namespace
