Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraRichEdit.API.Native

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' A class that contains information about the generated XtraReport
    ''' </summary>
    Friend Class ReportGenerator

        Private report As XtraReport 'The XtraReport instance

        Private currentReportField As XtraReportBase 'The currently processed XtraReportBase instance (DetailReportBand or XtraReport)

        Private currentBandField As Band 'The currently populated Band

        Public Sub New(ByVal report As XtraReport)
            Me.report = report
        End Sub

        ''' <summary>
        ''' Specifies global properties, which are applied to the entire report
        ''' </summary>
        ''' <param name="source">The source Snap document</param>
        Public Sub SetupReport(ByVal source As SnapDocument)
            'Create a DetailBand instance. An XtraReport object should always contain DetailBand
            CreateDetailBand()
            'Copy mail-merge settings (if any)
            CopyMailMergeOptions(source)
            'Copy page settings
            CopyPageSettings(source)
            'Copy parameters
            CopyParameters(source)
            'Specify initial values
            CurrentReport = report
            CurrentBand = report.Bands(BandKind.Detail)
            CurrentControl = Nothing
        End Sub

        ''' <summary>
        ''' Copies mail-merge settings to XtraReport
        ''' </summary>
        ''' <param name="source">The source Snap document</param>
        Private Sub CopyMailMergeOptions(ByVal source As SnapDocument)
            'Get mail-merge options
            Dim options As DevExpress.Snap.Core.Options.SnapMailMergeExportOptions = source.CreateSnapMailMergeExportOptions()
            'Specify data settings
            report.DataSource = options.DataSource
            report.DataMember = options.DataMember
        End Sub

        ''' <summary>
        ''' Specifies page settings
        ''' </summary>
        ''' <param name="source">The source Snap document</param>
        Private Sub CopyPageSettings(ByVal source As SnapDocument)
            'IMPORTANT! Snap document may contain multiple sections with different page settings. XtraReports do not support
            'this feature. So, only the first section's settings are taken into account
            Dim section As Section = source.Sections(0)
            'Get page settings
            report.PaperKind = section.Page.PaperKind
            report.PageWidth = CInt(GraphicsUnitConverter.Convert(section.Page.Width, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            report.PageHeight = CInt(GraphicsUnitConverter.Convert(section.Page.Height, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            report.Landscape = section.Page.Landscape
            'Get margins
            report.Margins.Left = CInt(GraphicsUnitConverter.Convert(section.Margins.Left, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            report.Margins.Top = CInt(GraphicsUnitConverter.Convert(section.Margins.Top, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            report.Margins.Right = CInt(GraphicsUnitConverter.Convert(section.Margins.Right, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            report.Margins.Bottom = CInt(GraphicsUnitConverter.Convert(section.Margins.Bottom, GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
        End Sub

        ''' <summary>
        ''' Copies parameters to XtraReport
        ''' </summary>
        ''' <param name="source">The source Snap document</param>
        Private Sub CopyParameters(ByVal source As SnapDocument)
            'Iterate through parameters in the Snap document
            For i As Integer = 0 To source.Parameters.Count - 1
                Dim snapParameter As Parameter = source.Parameters(i)
                'Create an XtraReport parameter and add it to the Parameters collection
                Dim reportParameter As DevExpress.XtraReports.Parameters.Parameter = New DevExpress.XtraReports.Parameters.Parameter()
                report.Parameters.Add(reportParameter)
                'Copy parameter settings
                reportParameter.Name = snapParameter.Name
                reportParameter.Type = snapParameter.Type
                reportParameter.Value = snapParameter.Value
            Next
        End Sub

        ''' <summary>
        ''' Creates DetailBand
        ''' </summary>
        Private Sub CreateDetailBand()
            report.Bands.Create(BandKind.Detail)
        End Sub

        ''' <summary>
        ''' Removes redundant empty bands
        ''' </summary>
        Public Sub CleanEmptyBands()
            'Call the core method
            CleanEmptyBandsCore(report)
        End Sub

        ''' <summary>
        ''' Removes redundant empty bands
        ''' </summary>
        ''' <param name="curReport">The currently processed XtraReportBase</param>
        Private Sub CleanEmptyBandsCore(ByVal curReport As XtraReportBase)
            'Iterate through all bands
            For i As Integer = curReport.Bands.Count - 1 To 0 Step -1
                'If the current band is a nested report (DetailReportBand)
                If TypeOf curReport.Bands(i) Is XtraReportBase Then
                    'Call this method for the nested report to clean its bands
                    CleanEmptyBandsCore(CType(curReport.Bands(i), XtraReportBase))
                'If no controls in a Band or this is not a mandatory band (TopMarginBand, BottomMarginBand, and DetailBand should always exist,
                'GroupHeaderBand specifies groups and can be empty)
                ElseIf curReport.Bands(i).Controls.Count = 0 Then
                    If Not(TypeOf curReport.Bands(i) Is TopMarginBand OrElse TypeOf curReport.Bands(i) Is BottomMarginBand OrElse TypeOf curReport.Bands(i) Is DetailBand OrElse TypeOf curReport.Bands(i) Is GroupHeaderBand) Then
                        'Remove a Band
                        curReport.Bands.RemoveAt(i)
                    Else
                        'Resize a Band to follow its content size (it will be automatically resized afterwards based on its content)
                        curReport.Bands(i).HeightF = 0
                    End If
                End If
            Next
        End Sub

        ''' <summary>
        ''' Locates a CalculatedField by its name
        ''' </summary>
        ''' <param name="displayName">A name of the field</param>
        ''' <returns>The located CalculatedField instance</returns>
        Public Function GetCalculatedField(ByVal displayName As String) As DevExpress.XtraReports.UI.CalculatedField
            'IMPORTANT! In Snap, different data sources may contain calculated fields with the same name
            'To avoid potential conflicts, adding a field with the same name is disabled
            'Iterate through fields
            For i As Integer = 0 To report.CalculatedFields.Count - 1
                'If the field's name equals to the specified name, return the field
                If Equals(report.CalculatedFields(i).DisplayName, displayName) Then Return report.CalculatedFields(i)
            Next

            'Otherwise, return null
            Return Nothing
        End Function

        ''' <summary>
        ''' Gets the available Y position to insert a new control
        ''' </summary>
        ''' <param name="control">A container control</param>
        ''' <returns>An integer offset value</returns>
        Public Shared Function GetVerticalOffset(ByVal control As XRControl) As Integer
            'If no controls in the container, return 0
            If control.Controls.Count = 0 Then
                Return 0
            Else
                'Iterate through nested controls to get the bottom-most position
                Dim offset As Integer = CInt(control.Controls(0).BoundsF.Bottom)
                For i As Integer = 1 To control.Controls.Count - 1
                    If control.Controls(i).BoundsF.Bottom > offset Then offset = CInt(control.Controls(i).BoundsF.Bottom)
                Next

                'Return the bottom-most position
                Return offset
            End If
        End Function

        Public Property CurrentReport As XtraReportBase 'The currently processed report
            Get
                Return currentReportField
            End Get

            Set(ByVal value As XtraReportBase)
                If currentReportField IsNot value Then
                    'Reset all values after changing the current report
                    currentReportField = value
                    currentBandField = Nothing
                    CurrentControl = Nothing
                End If
            End Set
        End Property

        Public Property CurrentBand As Band 'The currently processed band
            Get
                Return currentBandField
            End Get

            Set(ByVal value As Band)
                If currentBandField IsNot value Then
                    'Reset the control value after changing the current band
                    currentBandField = value
                    CurrentControl = Nothing
                End If
            End Set
        End Property

        Public Property CurrentControl As XRControl 'The currently processed control
    End Class
End Namespace
