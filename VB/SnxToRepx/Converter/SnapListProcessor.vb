Imports DevExpress.Snap
Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraReports.UI

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This class processes SnapList content
    ''' </summary>
    Friend Class SnapListProcessor

        Private server As DevExpress.Snap.SnapDocumentServer 'The source SnapDocumentServer

        Private list As DevExpress.Snap.Core.API.SnapList 'The source SnapList

        Private generator As SnxToRepx.Converter.ReportGenerator 'A ReportGenerator containing the constructed report

        Public Sub New(ByVal generator As SnxToRepx.Converter.ReportGenerator, ByVal server As DevExpress.Snap.SnapDocumentServer, ByVal list As DevExpress.Snap.Core.API.SnapList)
            Me.server = server
            Me.list = list
            Me.generator = generator
        End Sub

        ''' <summary>
        ''' Processes a SnapList
        ''' </summary>
        Public Sub ProcessSnapList()
            'Create a separate DetailReportBand for this list
            Dim listBand As DevExpress.XtraReports.UI.DetailReportBand = New DevExpress.XtraReports.UI.DetailReportBand()
            'Add it to the report and specify that this DetailReportBand is the currently processed report
            Me.generator.CurrentReport.Bands.Add(listBand)
            Me.generator.CurrentReport = listBand
            'Copy data source settings
            Me.CopyDataSettings()
            'Enable runtime customization for this list to access templates and other properties
            Me.list.BeginUpdate()
            'Process the list header
            Me.ProcessListHeader()
            'Process the row template
            Me.ProcessRowTemplate()
            'Process each group one by one
            For i As Integer = 0 To Me.list.Groups.Count - 1
                Dim info As DevExpress.Snap.Core.API.SnapListGroupInfo = Me.list.Groups(i)
                Me.ProcessGroupItem(info)
            Next

            'Process the list footer
            Me.ProcessListFooter()
            'Unlock the SnapList
            Me.list.EndUpdate()
            'Apply sorting, filters, etc.
            Me.CopyDataShapingSettings()
            'If no the DetailReportBand's ReportFooterBand was created, create and use it as the currently processed band
            If listBand.Bands(DevExpress.XtraReports.UI.BandKind.ReportFooter) Is Nothing Then
                Me.generator.CurrentBand = listBand.Bands.Create(DevExpress.XtraReports.UI.BandKind.ReportFooter)
            Else
                Me.generator.CurrentBand = listBand.Bands(DevExpress.XtraReports.UI.BandKind.ReportFooter)
            End If
        End Sub

        ''' <summary>
        ''' Copies data source settings from the Snap list
        ''' </summary>
        Private Sub CopyDataSettings()
            'Get data source information
            Dim dsInfo As DevExpress.Snap.Core.API.DataSourceInfo = Me.server.Document.DataSources(Me.list.DataSourceName)
            Dim currentReport As DevExpress.XtraReports.UI.XtraReportBase = Me.generator.CurrentReport
            'If data binding exists
            If dsInfo IsNot Nothing Then
                'Specify the DataSource and DataMember properties
                currentReport.DataSource = dsInfo.DataSource
                currentReport.DataMember = Me.list.DataMember
                'IMPORTANT! Detail SnapLists do not have the DataSource property specified. Instead, they use the parent's data source.
                'XtraReports require the DataSource property to be specified. Moreover, the DataMember property should contain the full path
                'to the data member (for example, "Products.ProductOrders" in XtraReports, while Snap uses "ProductOrders" in this case).
                'Check if the data source is null
                If currentReport.DataSource Is Nothing AndAlso currentReport.Report IsNot Nothing AndAlso Not Equals(currentReport.Report.DataMember, String.Empty) Then
                    'Get the parent's data source and use it as a data source for the current report
                    Dim dataMember As String = currentReport.DataMember
                    currentReport.DataSource = currentReport.Report.DataSource
                    'Combine the current and parent's data members to get the full path
                    currentReport.DataMember = $"{currentReport.Report.DataMember}.{dataMember}"
                End If

                'Copy calculated fields from the data source info
                Me.CopyCalculatedFields(dsInfo, Me.list.DataMember)
            End If
        End Sub

        ''' <summary>
        ''' Copies calculated fields to XtraReport
        ''' </summary>
        ''' <param name="dsInfo">Data source information</param>
        ''' <param name="dataMember">The current data member's name</param>
        Private Sub CopyCalculatedFields(ByVal dsInfo As DevExpress.Snap.Core.API.DataSourceInfo, ByVal dataMember As String)
            'Iterate through calculated fields
            For i As Integer = 0 To dsInfo.CalculatedFields.Count - 1
                Dim calcField As DevExpress.Snap.Core.API.CalculatedField = dsInfo.CalculatedFields(i)
                'If a calculated field does not yet exist 
                If Me.generator.GetCalculatedField(calcField.DisplayName) Is Nothing Then
                    'Create CalculatedField in XtraReport and copy field settings
                    Dim reportField As DevExpress.XtraReports.UI.CalculatedField = New DevExpress.XtraReports.UI.CalculatedField(calcField.DataSource, calcField.DataMember)
                    reportField.DisplayName = calcField.DisplayName
                    reportField.FieldType = calcField.FieldType
                    reportField.Expression = calcField.Expression
                    reportField.Name = calcField.Name
                    'Add this calculated field to the report
                    Me.generator.CurrentReport.RootReport.CalculatedFields.Add(reportField)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Copies filter and sort settings
        ''' </summary>
        Private Sub CopyDataShapingSettings()
            'Copy the first filter string
            If Me.list.Filters.Count > 0 Then Me.generator.CurrentReport.FilterString = Me.list.Filters(0)
            Dim band As DevExpress.XtraReports.UI.DetailBand = CType(Me.generator.CurrentReport.Bands(DevExpress.XtraReports.UI.BandKind.Detail), DevExpress.XtraReports.UI.DetailBand)
            'Iterate through all sort fields and add corresponding items to the current DetailBand
            For Each sortItem As DevExpress.Snap.Core.API.SnapListGroupParam In Me.list.Sorting
                'Convert the sort order to XtraReports sort order
                Dim sortOrder As DevExpress.XtraReports.UI.XRColumnSortOrder = SnxToRepx.Converter.ColumnSortOrderConverter.Convert(sortItem.SortOrder)
                band.SortFields.Add(New DevExpress.XtraReports.UI.GroupField() With {.FieldName = sortItem.FieldName, .SortOrder = sortOrder})
            Next
        End Sub

        ''' <summary>
        ''' Processes the list header
        ''' </summary>
        Private Sub ProcessListHeader()
            'Get the list header
            Dim listHeader As DevExpress.Snap.Core.API.SnapDocument = Me.list.ListHeader
            'Process the inner document and add its content to the ReportHeader band
            Me.ProcessTemplateDocument(listHeader, DevExpress.XtraReports.UI.BandKind.ReportHeader)
        End Sub

        ''' <summary>
        ''' Processes the list footer
        ''' </summary>
        Private Sub ProcessListFooter()
            'Get the list footer
            Dim listFooter As DevExpress.Snap.Core.API.SnapDocument = Me.list.ListFooter
            'Process the inner document and add its content to the ReportFooter band
            Me.ProcessTemplateDocument(listFooter, DevExpress.XtraReports.UI.BandKind.ReportFooter)
        End Sub

        ''' <summary>
        ''' Processes the row template
        ''' </summary>
        Private Sub ProcessRowTemplate()
            'Get the row template
            Dim rowTemplate As DevExpress.Snap.Core.API.SnapDocument = Me.list.RowTemplate
            'Process the inner document and add its content to DetailBand
            Me.ProcessTemplateDocument(rowTemplate, DevExpress.XtraReports.UI.BandKind.Detail)
        End Sub

        ''' <summary>
        ''' Processes the group item
        ''' </summary>
        ''' <param name="info">Information about the current group item</param>
        Private Sub ProcessGroupItem(ByVal info As DevExpress.Snap.Core.API.SnapListGroupInfo)
            'If the group header exists, get the template and add its content to the GroupHeader band
            If info.Header IsNot Nothing Then
                Dim groupHeaderTemplate As DevExpress.Snap.Core.API.SnapDocument = info.Header
                Me.ProcessTemplateDocument(groupHeaderTemplate, DevExpress.XtraReports.UI.BandKind.GroupHeader)
            End If

            'Specify the initial group header to link headers and footers
            Dim groupLevel As Integer = 0
            'Get the current GroupHeaderBand
            If TypeOf Me.generator.CurrentBand Is DevExpress.XtraReports.UI.GroupHeaderBand Then
                Dim groupHeader As DevExpress.XtraReports.UI.GroupHeaderBand = CType(Me.generator.CurrentBand, DevExpress.XtraReports.UI.GroupHeaderBand)
                'Iterate through group items (each item may be grouped by multiple fields)
                For i As Integer = 0 To info.Count - 1
                    'Get group parameters and add corresponding GroupFields
                    Dim groupParam As DevExpress.Snap.Core.API.SnapListGroupParam = info(i)
                    groupHeader.GroupFields.Add(New DevExpress.XtraReports.UI.GroupField(groupParam.FieldName, SnxToRepx.Converter.ColumnSortOrderConverter.Convert(groupParam.SortOrder)))
                Next

                'Get the current grouping level
                groupLevel = groupHeader.Level
            End If

            'If the group footer exists, get the template and add its content to the GroupFooter band
            If info.Footer IsNot Nothing Then
                Dim groupHeaderTemplate As DevExpress.Snap.Core.API.SnapDocument = info.Footer
                Me.ProcessTemplateDocument(groupHeaderTemplate, DevExpress.XtraReports.UI.BandKind.GroupFooter)
            End If

            If TypeOf Me.generator.CurrentBand Is DevExpress.XtraReports.UI.GroupFooterBand Then
                Dim groupFooter As DevExpress.XtraReports.UI.GroupFooterBand = CType(Me.generator.CurrentBand, DevExpress.XtraReports.UI.GroupFooterBand)
                groupFooter.Level = groupLevel
            End If
        End Sub

        ''' <summary>
        ''' Processes the template and adds its content to the Band of the specified kind
        ''' </summary>
        ''' <param name="template">The source document</param>
        ''' <param name="bandKind">The Band kind to create/access a Band</param>
        Private Sub ProcessTemplateDocument(ByVal template As DevExpress.Snap.Core.API.SnapDocument, ByVal bandKind As DevExpress.XtraReports.UI.BandKind)
            'If the template is not empty
            If template.Range.Length > 1 Then
                'Create the corresponding band and set it as an active Band
                Dim currentBand As DevExpress.XtraReports.UI.Band = Me.generator.CurrentReport.Bands.Create(bandKind)
                Me.generator.CurrentBand = currentBand
                'Create a temporary SnapDocumentServer and copy the template document there
                Using tempServer As DevExpress.Snap.SnapDocumentServer = New DevExpress.Snap.SnapDocumentServer()
                    tempServer.SnxBytes = template.SaveDocument(DevExpress.Snap.Core.API.SnapDocumentFormat.Snap)
                    'Copy data sources
                    For i As Integer = 0 To Me.server.Document.DataSources.Count - 1
                        If Me.server.Document.DataSources(CInt((i))).DataSource IsNot Nothing Then tempServer.Document.DataSources.Add(Me.server.Document.DataSources(i))
                    Next

                    'Create a SnapLayoutProcessor instance and process the document layout
                    Dim processor As SnxToRepx.Converter.SnapLayoutProcessor = New SnxToRepx.Converter.SnapLayoutProcessor(Me.generator, tempServer)
                    processor.ProcessDocument()
                End Using

                'Restore the current Band (it may change during layout processing)
                Me.generator.CurrentBand = currentBand
                'Set the Band height to 0. It will automatically adjust its size based on content
                Me.generator.CurrentBand.HeightF = 0F
            End If
        End Sub
    End Class
End Namespace
