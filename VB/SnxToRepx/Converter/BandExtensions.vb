Imports DevExpress.XtraReports.UI
Imports System.Runtime.CompilerServices

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This class contains useful extensions for the BandCollection class
    ''' </summary>
    Friend Module BandExtensions

        ''' <summary>
        ''' Creates a new band of the specified kind and adds it to the collection
        ''' </summary>
        ''' <param name="col">A BandCollection instance that should hold a new Band</param>
        ''' <param name="kind">A band kind to create a new Band</param>
        ''' <returns>The Band of the specified kind</returns>
        <Extension()>
        Public Function Create(ByVal col As BandCollection, ByVal kind As BandKind) As Band
            Dim band As Band = Nothing
            'Check the BandKind value and create a new band
            Select Case kind
                Case BandKind.ReportHeader
                    band = New ReportHeaderBand()
                Case BandKind.Detail
                    'Each report can contain only a single DetailBand.
                    'If it does not yet exist, create a new band
                    If col(kind) Is Nothing Then
                        band = New DetailBand()
                    Else
                        'Otherwise, return the existing DetailBand
                        band = col(kind)
                    End If

                Case BandKind.GroupHeader
                    band = New GroupHeaderBand()
                Case BandKind.GroupFooter
                    band = New GroupFooterBand()
                Case BandKind.ReportFooter
                    band = New ReportFooterBand()
                Case BandKind.TopMargin
                    band = New TopMarginBand()
                Case BandKind.BottomMargin
                    band = New BottomMarginBand()
                Case BandKind.PageHeader
                    band = New PageHeaderBand()
                Case BandKind.PageFooter
                    band = New PageFooterBand()
                Case Else
            End Select

            'If a band was successfully created and it is not yet added to the collection,
            If band IsNot Nothing AndAlso Not col.Contains(band) Then
                'add this band to the collection and set its height to 0
                'because it will automatically resize itself based on content
                col.Add(band)
                band.HeightF = 0
            End If

            'Return the created band
            Return band
        End Function
    End Module
End Namespace
