Imports DevExpress.Data
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Drawing

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This class converts RichEdit border styles to XtraReport border styles
    ''' </summary>
    Friend Module BorderLineStyleConverter

        'DOCS: Only a limited set of border styles is available in reports
        ''' <summary>
        ''' This method converts DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle to DevExpress.XtraPrinting.BorderDashStyle
        ''' </summary>
        ''' <param name="inputStyle">The source DevExpress.XtraRichEdit.API.Native.TableBorderLineStyle</param>
        ''' <returns>The resulting DevExpress.XtraPrinting.BorderDashStyle</returns>
        Public Function Convert(ByVal inputStyle As TableBorderLineStyle) As BorderDashStyle
            Select Case inputStyle
                Case TableBorderLineStyle.Dashed
                    Return BorderDashStyle.Dash
                Case TableBorderLineStyle.DotDash
                    Return BorderDashStyle.DashDot
                Case TableBorderLineStyle.DotDotDash
                    Return BorderDashStyle.DashDotDot
                Case TableBorderLineStyle.Dotted
                    Return BorderDashStyle.Dot
                Case TableBorderLineStyle.Double
                    Return BorderDashStyle.Double
                Case Else
                    Return BorderDashStyle.Solid
            End Select
        End Function
    End Module

    ''' <summary>
    ''' This class converts RichEdit alignment settings to XtraReport alignment
    ''' </summary>
    Friend Module AlignmentConverter

        ''' <summary>
        ''' This method converts DevExpress.XtraRichEdit.API.Native.TableCellVerticalAlignment to DevExpress.XtraPrinting.TextAlignment
        ''' </summary>
        ''' <param name="vertAlign">The source DevExpress.XtraRichEdit.API.Native.TableCellVerticalAlignment</param>
        ''' <returns>The resulting DevExpress.XtraPrinting.TextAlignment</returns>
        Public Function Convert(ByVal vertAlign As TableCellVerticalAlignment) As TextAlignment
            Select Case vertAlign
                Case TableCellVerticalAlignment.Top
                    Return TextAlignment.TopLeft
                Case TableCellVerticalAlignment.Center
                    Return TextAlignment.MiddleLeft
                Case TableCellVerticalAlignment.Bottom
                    Return TextAlignment.BottomLeft
                Case Else
                    Return TextAlignment.TopLeft
            End Select
        End Function

        ''' <summary>
        ''' This method converts DevExpress.XtraRichEdit.API.Native.ParagraphAlignment to the markup representation
        ''' </summary>
        ''' <param name="parAlign">The source DevExpress.XtraRichEdit.API.Native.ParagraphAlignment</param>
        ''' <returns>The resulting markup tag with alignment</returns>
        Public Function Convert(ByVal parAlign As ParagraphAlignment) As String
            Select Case parAlign
                Case ParagraphAlignment.Justify
                    Return "align=justify"
                Case ParagraphAlignment.Right
                    Return "align=right"
                Case ParagraphAlignment.Center
                    Return "align=center"
                Case ParagraphAlignment.Left
                    Return "align=left"
                Case Else
                    Return String.Empty
            End Select
        End Function
    End Module

    ''' <summary>
    ''' This class converts the sorting order to XtraReports sorting order
    ''' </summary>
    Friend Module ColumnSortOrderConverter

        ''' <summary>
        ''' This method converts DevExpress.Data.ColumnSortOrder to DevExpress.XtraReports.UI.XRColumnSortOrder
        ''' </summary>
        ''' <param name="sortOrder">The source DevExpress.Data.ColumnSortOrder</param>
        ''' <returns>The resulting DevExpress.XtraReports.UI.XRColumnSortOrder</returns>
        Public Function Convert(ByVal sortOrder As ColumnSortOrder) As XRColumnSortOrder
            Select Case sortOrder
                Case ColumnSortOrder.None
                    Return XRColumnSortOrder.None
                Case ColumnSortOrder.Descending
                    Return XRColumnSortOrder.Descending
                Case Else
                    Return XRColumnSortOrder.Ascending
            End Select
        End Function
    End Module

    ''' <summary>
    ''' This class converts Snap summary settings to XtraReports summary settings
    ''' </summary>
    Friend Module SummaryConverter

        ''' <summary>
        ''' This method converts DevExpress.Data.SummaryItemType to the markup representation
        ''' </summary>
        ''' <param name="summaryFunc">The source DevExpress.XtraRichEdit.API.Native.ParagraphAlignment</param>
        ''' <returns>The resulting markup tag with alignment</returns>
        Public Function Convert(ByVal summaryFunc As SummaryItemType) As String
            Select Case summaryFunc
                Case SummaryItemType.Average
                    Return "sumAvg"
                Case SummaryItemType.Count
                    Return "sumCount"
                Case SummaryItemType.Max
                    Return "sumMax"
                Case SummaryItemType.Min
                    Return "sumMin"
                Case SummaryItemType.Sum, SummaryItemType.Custom
                    Return "sumSum"
                Case Else
                    Return String.Empty
            End Select
        End Function

        ''' <summary>
        ''' This method converts DevExpress.Snap.Core.Fields.SummaryRunning to DevExpress.XtraReports.UI.SummaryRunning
        ''' </summary>
        ''' <param name="running">The source DevExpress.Snap.Core.Fields.SummaryRunning</param>
        ''' <returns>The resulting DevExpress.XtraReports.UI.SummaryRunning</returns>
        Public Function Convert(ByVal running As DevExpress.Snap.Core.Fields.SummaryRunning) As SummaryRunning
            Select Case running
                Case DevExpress.Snap.Core.Fields.SummaryRunning.Group
                    Return SummaryRunning.Group
                Case DevExpress.Snap.Core.Fields.SummaryRunning.Report
                    Return SummaryRunning.Report
                Case Else
                    Return SummaryRunning.None
            End Select
        End Function
    End Module
End Namespace
