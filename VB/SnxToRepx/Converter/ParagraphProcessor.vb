Imports DevExpress.Snap
Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraRichEdit.API.Layout
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.Collections.Generic
Imports System.Drawing

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This class processes text portions
    ''' </summary>
    Friend Class ParagraphProcessor
        Inherits ContentProcessorBase

        Private rows As List(Of LayoutRow) 'LayoutRows that contain processed text parts

        Public Sub New(ByVal generator As ReportGenerator, ByVal server As SnapDocumentServer, ByVal rows As List(Of LayoutRow))
            MyBase.New(generator, server)
            Me.rows = rows
        End Sub

        ''' <summary>
        ''' Generates a markup and passes it to the corresponding XRControl
        ''' </summary>
        Public Sub ProcessParagraphs()
            'Get a range of processed paragraphs
            Dim range As DocumentRange = Server.Document.CreateRange(rows(0).Range.Start, rows(rows.Count - 1).Range.Start + rows(rows.Count - 1).Range.Length - rows(0).Range.Start - 1)
            'Check if this text portion contains summaries
            Dim summaryRunning As Core.Fields.SummaryRunning = GetSummary(range)
            Dim markupGenerator As MarkupGenerator
            'If a summary calculation exists, an expression is generated
            If summaryRunning <> Core.Fields.SummaryRunning.None Then
                markupGenerator = New ExpressionMarkupGenerator(Generator, Server.Document)
            Else
                'Otherwise, a markup used in mail-merge reports is generated
                markupGenerator = New MarkupGenerator(Generator, Server.Document)
            End If

            'Create a DocumentIterator instance to iterate through all text parts within the specified range
            Dim iterator As DocumentIterator = New DocumentIterator(range, True)
            While iterator.MoveNext()
                iterator.Current.Accept(markupGenerator)
            End While

            'Get the resulting markup
            Dim markup As String = markupGenerator.Text
            Dim contentControl As XRControl
            'If the current control is already created (for example, XRTableCell), then use it
            If Generator.CurrentControl IsNot Nothing Then
                contentControl = Generator.CurrentControl
            Else
                'Otherwise, create a new control, set its properties, and add it to the parent Band
                contentControl = GenerateControl()
                contentControl.CanShrink = True
                contentControl.Padding = GetPaddingInfo(range)
                contentControl.BoundsF = CalculateParagraphBounds(rows, Generator.CurrentBand)
                Generator.CurrentBand.Controls.Add(contentControl)
            End If

            'If a summary calculation exists
            If summaryRunning <> Core.Fields.SummaryRunning.None AndAlso TypeOf contentControl Is XRLabel Then
                'Add an expression binding and specify the area for which summary is calculated
                contentControl.ExpressionBindings.Add(New ExpressionBinding("Text", markup))
                CType(contentControl, XRLabel).Summary.Running = SummaryConverter.Convert(summaryRunning)
            Else
                'Otherwise, use the default markup
                contentControl.Text = markup
            End If
        End Sub

        ''' <summary>
        ''' Checks if summary functions are used in the specified range
        ''' </summary>
        ''' <param name="range">A document range to check</param>
        ''' <returns>The SummaryRunning value</returns>
        Private Function GetSummary(ByVal range As DocumentRange) As Core.Fields.SummaryRunning
            'Get all fields in the specified range
            Dim fields As ReadOnlyFieldCollection = Server.Document.Fields.Get(range)
            'Iterate through fields
            For Each field As Field In fields
                'Check if it is a SnapText that has SummaryRunning specified
                Dim entity As SnapEntity = Server.Document.ParseField(field)
                'Return the SummaryRunning value
                If TypeOf entity Is SnapText AndAlso CType(entity, SnapText).SummaryRunning <> Core.Fields.SummaryRunning.None AndAlso CType(entity, SnapText).SummaryFunc <> DevExpress.Data.SummaryItemType.None Then Return CType(entity, SnapText).SummaryRunning
            Next

            'If no fields with summaries are found, return None
            Return Core.Fields.SummaryRunning.None
        End Function

        ''' <summary>
        ''' Gets the resulting bounds for XRControl within the parent control
        ''' </summary>
        ''' <param name="rows">Layout rows containing text</param>
        ''' <param name="parentControl">The parent control to calculate boundaries</param>
        ''' <returns>Control boundaries</returns>
        Private Function CalculateParagraphBounds(ByVal rows As List(Of LayoutRow), ByVal parentControl As XRControl) As RectangleF
            'Get the paragraph width
            Dim paragraphWidth As Integer = rows(0).Bounds.Width
            'Iterate through rows to calculate the resulting width and height
            Dim paragraphHeight As Integer = rows(0).Bounds.Height
            For i As Integer = 1 To rows.Count - 1
                If rows(i).Bounds.Width > paragraphWidth Then paragraphWidth = rows(i).Bounds.Width
                paragraphHeight += rows(i).Bounds.Height
            Next

            'Calculating initial boundaries
            Dim bounds As Rectangle = New Rectangle(Math.Max(rows(0).GetRelativeBounds(rows(0).Parent).Left, 0), 0, paragraphWidth, paragraphHeight)
            'Convert boundaries to measurement units used in XtraReports
            bounds = GraphicsUnitConverter.Convert(bounds, GraphicsDpi.Twips, GraphicsDpi.HundredthsOfAnInch)
            'Take paddings into account
            bounds.Offset(-parentControl.Padding.Left, ReportGenerator.GetVerticalOffset(parentControl))
            'Return boundaries
            Return bounds
        End Function

        ''' <summary>
        ''' Creates a control to display content
        ''' </summary>
        ''' <returns>An XRControl for future use</returns>
        Private Function GenerateControl() As XRControl
            'IMPORTANT! Auxiliary Snap controls, such as checkboxes, can be located in the middle of formatted text.
            'XtraReports do not support such layout, so textual representation is used
            Dim label As XRLabel = New XRLabel()
            'Allow HTML-style markup
            label.AllowMarkupText = True
            'Remove borders
            label.Borders = BorderSide.None
            'Return the control
            Return label
        End Function
    End Class
End Namespace
