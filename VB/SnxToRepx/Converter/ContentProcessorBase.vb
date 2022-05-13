Imports DevExpress.Snap
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraRichEdit.API.Native

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' A base class for content processors
    ''' </summary>
    Friend MustInherit Class ContentProcessorBase

        Private serverField As SnapDocumentServer 'A SnapDocumentServer containing the processed document

        Private generatorField As ReportGenerator 'A ReportGenerator that stores the resulting XtraReport

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="generator">A ReportGenerator that stores the resulting XtraReport</param>
        ''' <param name="server">A SnapDocumentServer containing the processed document</param>
        Public Sub New(ByVal generator As ReportGenerator, ByVal server As SnapDocumentServer)
            'Assign values
            serverField = server
            generatorField = generator
        End Sub

        ''' <summary>
        ''' Retrieves padding for current content
        ''' </summary>
        ''' <param name="range">The processed document range</param>
        ''' <returns>PaddingInfo containing required paddings</returns>
        Protected Overridable Function GetPaddingInfo(ByVal range As DocumentRange) As PaddingInfo
            'Get paragraph properties
            Dim parProps As ParagraphProperties = serverField.Document.BeginUpdateParagraphs(range)
            'Set initial spacing values
            Dim spacingBefore As Integer = 0
            Dim spacingAfter As Integer = 0
            'Obtain spacing
            If parProps.SpacingBefore IsNot Nothing Then spacingBefore = CInt(GraphicsUnitConverter.Convert(CSng(parProps.SpacingBefore), GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            If parProps.SpacingAfter IsNot Nothing Then spacingAfter = CInt(GraphicsUnitConverter.Convert(CSng(parProps.SpacingAfter), GraphicsDpi.Document, GraphicsDpi.HundredthsOfAnInch))
            serverField.Document.EndUpdateParagraphs(parProps)
            'Create and return a new PaddingInfo object containing required values
            Dim paddingInfo As PaddingInfo = New PaddingInfo(0, 0, spacingBefore, spacingAfter)
            Return paddingInfo
        End Function

        'Properties
        Protected ReadOnly Property Server As SnapDocumentServer
            Get
                Return serverField
            End Get
        End Property

        Protected ReadOnly Property Generator As ReportGenerator
            Get
                Return generatorField
            End Get
        End Property
    End Class
End Namespace
