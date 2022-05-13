Imports DevExpress.Snap
Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraRichEdit.API.Layout
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Collections.Generic
Imports System.Drawing

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' A delegate used to pass an action to nested processors
    ''' </summary>
    ''' <param name="rows">Layout rows with text content</param>
    ''' <param name="tables">Layout tables</param>
    Friend Delegate Sub CellContentProcessorDelegate(ByVal rows As LayoutRowCollection, ByVal tables As LayoutTableCollection)

    ''' <summary>
    ''' This class processes the layout of the current document
    ''' </summary>
    Friend Class SnapLayoutProcessor

        Private server As SnapDocumentServer 'A SnapDocumentServer instance containing the current document

        Private snapLists, processedSnapLists As SnapListCollection 'Collections of SnapLists containing all and processed lists

        Private generator As ReportGenerator 'A ReportGenerator containing the currently constructed XtraReport

        Public Sub New(ByVal generator As ReportGenerator, ByVal server As SnapDocumentServer)
            Me.server = server
            Me.generator = generator
        End Sub

        ''' <summary>
        ''' Iterates through pages and processes their content
        ''' </summary>
        Public Sub ProcessDocument()
            'IMPORTANT! In Snap, you can create a multi-column layout with completely different
            'document elements. XtraReports do not support this layout, so a single-column report is generated.
            'Disable history to avoid potential document changing conflicts
            server.Options.DocumentCapabilities.Undo = DevExpress.XtraRichEdit.DocumentCapability.Disabled
            'Create and initialize collections
            snapLists = New SnapListCollection(server.Document)
            snapLists.PrepareCollection()
            processedSnapLists = New SnapListCollection(server.Document)
            'Display field codes
            'This may help to distinguish between "template" and "result" fields
            server.Document.BeginUpdate()
            For i As Integer = 0 To server.Document.Fields.Count - 1
                server.Document.Fields(i).ShowCodes = True
            Next

            server.Document.EndUpdate()
            'Iterate through pages
            For i As Integer = 0 To server.DocumentLayout.GetPageCount() - 1
                Dim page As LayoutPage = server.DocumentLayout.GetPage(i)
                'Iterate through page areas
                For Each pageArea As LayoutPageArea In page.PageAreas
                    'Iterate through columns
                    For Each column As LayoutColumn In pageArea.Columns
                        'Process rows and tables in a column
                        ProcessLayoutElements(column.Rows, column.Tables)
                    Next
                Next
            Next
        End Sub

        ''' <summary>
        ''' Processes tables and rows
        ''' </summary>
        ''' <param name="rows">A collection of LayoutRow elements</param>
        ''' <param name="tables">A collection of tables</param>
        Friend Sub ProcessLayoutElements(ByVal rows As LayoutRowCollection, ByVal tables As LayoutTableCollection)
            'TableCells with both text and nested tables/charts are currently not supported. Only the nested table/chart is displayed
            'IMPORTANT! XRTableCell and XRLabel content is built using the markup and expressions. However, if they have nested controls, neither markup nor
            'expressions are in effect. Thus, only nested controls are visible in this case. In Snap, you can place any content in a table cell
            '(including multiple nesting table levels). There is no generic solution to automatically convert this content to XtraReports.
            'Create a collection of layout elements and add rows and tables to this collection in an appropriate order (after sorting)
            Dim elements As List(Of RangedLayoutElement) = New List(Of RangedLayoutElement)()
            elements.AddRange(rows)
            elements.AddRange(tables)
            elements.Sort(New RangedElementComparer())
            'If the collection is not empty
            If elements.Count > 0 Then
                'Create a collection for text rows
                Dim paragraphRows As List(Of LayoutRow) = New List(Of LayoutRow)()
                'Iterate through elements
                For i As Integer = 0 To elements.Count - 1
                    Dim element As RangedLayoutElement = elements(i)
                    'Check if an element contains SnapList
                    Dim list As SnapList = snapLists.GetSnapListByPosition(element.Range.Start)
                    If list IsNot Nothing Then
                        If Not processedSnapLists.ContainsField(list.Field) Then
                            'If a list is found and is not yet processed, process it
                            ProcessSnapList(list)
                        Else
                            'Otherwise, continue execution
                            Continue For
                        End If
                    Else
                        'If the current element is a table
                        If TypeOf element Is LayoutTable Then
                            'Create a TableProcessor instance and process the element
                            Dim tableProcessor As TableProcessor = New TableProcessor(generator, server, CType(element, LayoutTable), AddressOf Me.ProcessLayoutElements)
                            tableProcessor.ProcessTable()
                        'If the current element is a row
                        ElseIf TypeOf element Is LayoutRow Then
                            'Process it
                            Dim row As LayoutRow = CType(element, LayoutRow)
                            ProcessLayoutRow(elements, row, paragraphRows, i)
                        End If
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Trims empty rows (containing empty paragraphs) from the beginning and from the end of a list
        ''' </summary>
        ''' <param name="rows">LayoutRows to process</param>
        Private Sub CorrectParagraphRows(ByVal rows As List(Of LayoutRow))
            Dim rowsToRemove As List(Of LayoutRow) = New List(Of LayoutRow)()
            'Iterate through rows in the forward direction
            For i As Integer = 0 To rows.Count - 1
                If rows(i).Range.Length <= 1 Then
                    'If a row is empty, add it to the collection of rows to be removed
                    rowsToRemove.Add(rows(i))
                Else
                    'If a non-empty row is found, stop the process
                    Exit For
                End If
            Next

            'Remove empty rows
            For Each row As LayoutRow In rowsToRemove
                rows.Remove(row)
            Next

            rowsToRemove.Clear()
            'Iterate through rows in the forward direction
            For i As Integer = rows.Count - 1 To 0 Step -1
                If rows(i).Range.Length <= 1 Then
                    rowsToRemove.Add(rows(i))
                Else
                    Exit For
                End If
            Next

            'Remove empty rows
            For Each row As LayoutRow In rowsToRemove
                rows.Remove(row)
            Next
        End Sub

        ''' <summary>
        ''' Processes the Snap List
        ''' </summary>
        ''' <param name="list">The source Snap list</param>
        Private Sub ProcessSnapList(ByVal list As SnapList)
            'Get the currently processed report
            Dim report As XtraReportBase = generator.CurrentReport
            'Create SnapListProcessor and process the list
            Dim listProcessor As SnapListProcessor = New SnapListProcessor(generator, server, list)
            listProcessor.ProcessSnapList()
            'Add this list to the collection of processed Snap lists
            processedSnapLists.Add(list.Field)
            'Obtain the current Band and XRControl
            Dim band As Band = generator.CurrentBand
            Dim control As XRControl = generator.CurrentControl
            'Restore the currently processed report, Band, and XRControl
            generator.CurrentReport = report
            generator.CurrentBand = band
            generator.CurrentControl = control
        End Sub

        ''' <summary>
        ''' Processes a single LayoutRow instance
        ''' </summary>
        ''' <param name="elements">A collection of layout elements</param>
        ''' <param name="row">The current row</param>
        ''' <param name="paragraphRows">A collection to store text rows</param>
        ''' <param name="index">An index of the current row in the element collection</param>
        Private Sub ProcessLayoutRow(ByVal elements As List(Of RangedLayoutElement), ByVal row As LayoutRow, ByVal paragraphRows As List(Of LayoutRow), ByVal index As Integer)
            Dim chart As SnapChart = Nothing
            'Get a row range
            Dim rowRange As DocumentRange = server.Document.CreateRange(row.Range.Start, row.Range.Length)
            'Iterate through fields to check if they belong to this range
            For j As Integer = 0 To server.Document.Fields.Count - 1
                Dim field As Field = server.Document.Fields(j)
                'If a field starts within the range, check the field type
                If field.Range.Start.ToInt() >= row.Range.Start AndAlso field.Range.Start.ToInt() < row.Range.Start + row.Range.Length Then
                    Dim entity As SnapEntity = server.Document.ParseField(field)
                    'If it is a chart, store it in a variable
                    If TypeOf entity Is SnapChart Then
                        chart = CType(entity, SnapChart)
                        Exit For
                    End If
                End If
            Next

            'If no charts found, add this row to the text row collection
            If chart Is Nothing Then paragraphRows.Add(row)
            'If this is the last element, or the subsequent element is not a row, or if the next element contains a snap list
            If index = elements.Count - 1 OrElse Not(TypeOf elements(index + 1) Is LayoutRow) OrElse snapLists.GetSnapListByPosition(elements(index + 1).Range.Start) IsNot Nothing OrElse chart IsNot Nothing Then
                'Remove empty rows
                CorrectParagraphRows(paragraphRows)
                If paragraphRows.Count > 0 Then
                    'Create a ParagraphProcessor instance and process rows
                    Dim paragraphProcessor As ParagraphProcessor = New ParagraphProcessor(generator, server, paragraphRows)
                    paragraphProcessor.ProcessParagraphs()
                    'Clear the collection of processed rows
                    paragraphRows.Clear()
                End If
            End If

            'If the chart was found
            If chart IsNot Nothing Then
                'Process the chart
                ProcessChart(chart)
            End If
        End Sub

        ''' <summary>
        ''' Processes the chart element
        ''' </summary>
        ''' <param name="chart">The source chart</param>
        Private Sub ProcessChart(ByVal chart As SnapChart)
            'Enable runtime customization
            chart.BeginUpdate()
            'Get the series collection
            Dim series As DevExpress.XtraCharts.SeriesCollection = chart.Series
            'The only way to access the inner DevExpress.XtraCharts.Native.Chart object is to use System.Reflection
            'We use this object to copy all chart properties at once
            Dim pi As System.Reflection.PropertyInfo = GetType(DevExpress.XtraCharts.SeriesCollection).GetProperty("Chart", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)
            Dim innerChart As Core.Native.SNChart = TryCast(pi?.GetValue(series), Core.Native.SNChart)
            If innerChart IsNot Nothing Then
                'Get the parent control
                Dim parentControl As XRControl = generator.CurrentControl
                If parentControl Is Nothing Then parentControl = generator.CurrentBand
                'Create an XRChart instance and add it to the parent container
                Dim reportChart As XRChart = New XRChart()
                parentControl.Controls.Add(reportChart)
                'Get the chart size in measurement units used in XtraReports
                Dim chartSize As Size = DevExpress.XtraPrinting.GraphicsUnitConverter.Convert(chart.Size, DevExpress.XtraPrinting.GraphicsDpi.Pixel, DevExpress.XtraPrinting.GraphicsDpi.HundredthsOfAnInch)
                'Set boundaries
                reportChart.BoundsF = New RectangleF(New PointF(0, ReportGenerator.GetVerticalOffset(parentControl)), chartSize)
                'Assign properties from the source chart
                CType(reportChart, DevExpress.XtraCharts.Native.IChartContainer).Chart.Assign(innerChart)
            End If

            'Unlock the snap entity
            chart.EndUpdate()
        End Sub
    End Class
End Namespace
