Imports DevExpress.DataAccess.Sql
Imports DevExpress.Snap
Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraReports.UI
Imports DevExpress.XtraRichEdit.API.Native
Imports System
Imports System.IO

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This class converts an SNX report to XtraReport
    ''' </summary>
    Public Class SnapReportConverter
        Implements IDisposable

        Private generator As ReportGenerator 'The Report generator instance that contains the current report

        ''' <summary>
        ''' Converts an existing SnapDocument instance to XtraReport
        ''' </summary>
        ''' <param name="src">The source SnapDocument</param>
        ''' <returns>An XtraReport instance</returns>
        Public Function Convert(ByVal src As SnapDocument) As XtraReport
            'Save a document to a byte array and call the GenerateReport method
            Return GenerateReport(src.SaveDocument(SnapDocumentFormat.Snap))
        End Function

        ''' <summary>
        ''' Converts an SNX document stored in the file system to XtraReport
        ''' </summary>
        ''' <param name="path">A path to the SNX file</param>
        ''' <returns>An XtraReport instance</returns>
        Public Function Convert(ByVal path As String) As XtraReport
            Return GenerateReport(File.ReadAllBytes(path))
        End Function

        ''' <summary>
        ''' Converts an SNX document stored in a stream to XtraReport
        ''' </summary>
        ''' <param name="stream">A Stream object containing the SNX document</param>
        ''' <returns>An XtraReport instance</returns>
        Public Function Convert(ByVal stream As Stream) As XtraReport
            'Use a SnapDocumentServer instance to load a document and convert it to a byte array
            Using server As SnapDocumentServer = New SnapDocumentServer()
                AddHandler server.Document.ConfigureDataConnection, AddressOf OnConfigureDataConnection
                server.LoadDocument(stream)
                RemoveHandler server.Document.ConfigureDataConnection, AddressOf OnConfigureDataConnection
                'Call the GenerateReport method to generate XtraReport                
                Return GenerateReport(server.SnxBytes)
            End Using
        End Function

        ''' <summary>
        ''' Loads a document and converts it to XtraReport
        ''' </summary>
        ''' <param name="snxBytes">A byte array containing the input SNX document</param>
        ''' <returns>An XtraReport instance</returns>
        Private Function GenerateReport(ByVal snxBytes As Byte()) As XtraReport
            'Set this property to ensure that subsequent tables are not merged
            DevExpress.XtraRichEdit.RichEditControlCompatibility.MergeSuccessiveTables = False
            Try
                'Create a new SnapDocumentServer to operate with the document
                Using server As SnapDocumentServer = New SnapDocumentServer()
                    'Handle the ConfigureDataConnection event to allow users to specify data connection settings
                    AddHandler server.Document.ConfigureDataConnection, AddressOf OnConfigureDataConnection
                    'Pass SnxBytes to server
                    server.SnxBytes = snxBytes
                    RemoveHandler server.Document.ConfigureDataConnection, AddressOf OnConfigureDataConnection
                    'Return the loaded SnapDocument instance
                    Dim source As SnapDocument = server.Document
                    'Set units to Document to simplify conversion to report units
                    source.Unit = DevExpress.Office.DocumentUnit.Document
                    'Create an XtraReport instance and initialize ReportGenerator
                    Dim report As XtraReport = New XtraReport()
                    generator = New ReportGenerator(report)
                    'Apply global settings (paper, parameters, etc.) to the created report
                    generator.SetupReport(source)
                    'Ensure that SnapLists start in new paragraphs
                    PrepareSnapListsForExport(source)
                    'Create a SnapLayoutProcessor instance to process the document layout
                    Dim layoutProcessor As SnapLayoutProcessor = New SnapLayoutProcessor(generator, server)
                    'Start processing
                    layoutProcessor.ProcessDocument()
                    'Add header and footer data to the report
                    ProcessHeadersFooters(server)
                    'Empty bands may appear after the XtraReport layout is built. Call this method to remove redundant bands
                    generator.CleanEmptyBands()
                    'Return the resulting report
                    Return report
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Handles the ConfigureDataConnection event and forwards it to SnapReportConverter.ConfigureDataConnection
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        Private Sub OnConfigureDataConnection(ByVal sender As Object, ByVal e As ConfigureDataConnectionEventArgs)
            RaiseEvent ConfigureDataConnection(Me, e)
        End Sub

        ''' <summary>
        ''' This method ensures that every SnapList starts with a new paragraph
        ''' </summary>
        ''' <param name="document">A source document</param>
        Private Sub PrepareSnapListsForExport(ByVal document As SnapDocument)
            'IMPORTANT! Sometimes SnapLists may start on the same row with static content.
            'In this situation, it is difficult to process SnapLists along with other content.
            'To avoid this situation, ensure that every SnapList starts with a new paragraph.
            'Lock all updates
            document.BeginUpdate()
            'Iterate through all fields
            For i As Integer = 0 To document.Fields.Count - 1
                Dim field As Field = document.Fields(i)
                'Parse the current field to check if it is a SnapList
                Dim entity As SnapEntity = document.ParseField(field)
                If TypeOf entity Is SnapList Then
                    'Check if this SnapList starts at the beginning of a paragraph
                    Dim parentParagraph As Paragraph = document.Paragraphs.Get(field.Range.Start)
                    'If not, insert a paragraph just before the field
                    If parentParagraph.Range.Start.ToInt() <> field.Range.Start.ToInt() Then document.Paragraphs.Insert(field.Range.Start)
                End If
            Next

            'Disable locking
            document.EndUpdate()
        End Sub

        Private Sub ProcessHeadersFooters(ByVal sourceServer As SnapDocumentServer)
            'IMPORTANT! There is no such term as First/Primary/Odd/Even for headers and footers in XtraReports.
            'Thus, only the primary header and footer are processed
            'Get the first section and check if it has a header of the Primary type
            Dim section As Section = sourceServer.Document.Sections(0)
            If section.HasHeader(HeaderFooterType.Primary) Then
                'If so, copy its content to a separate SnapDocumentServer and process its content
                Using server As SnapDocumentServer = New SnapDocumentServer()
                    Dim headerDocument As SubDocument = section.BeginUpdateHeader(HeaderFooterType.Primary)
                    server.Document.AppendDocumentContent(headerDocument.Range, InsertOptions.KeepSourceFormatting)
                    section.EndUpdateHeader(headerDocument)
                    ProcessSubDocument(server, BandKind.PageHeader)
                End Using
            End If

            'Check if the section has a footer of the Primary type
            If section.HasFooter(HeaderFooterType.Primary) Then
                'If so, copy its content to a separate SnapDocumentServer and process its content
                Using server As SnapDocumentServer = New SnapDocumentServer()
                    Dim footerDocument As SubDocument = section.BeginUpdateFooter(HeaderFooterType.Primary)
                    server.Document.AppendDocumentContent(footerDocument.Range, InsertOptions.KeepSourceFormatting)
                    section.EndUpdateFooter(footerDocument)
                    ProcessSubDocument(server, BandKind.PageFooter)
                End Using
            End If
        End Sub

        ''' <summary>
        ''' Copies header/footer information to PageHeaderBand/PageFooterBand
        ''' </summary>
        ''' <param name="server">A SnapDocumentServer instance containing the source document</param>
        ''' <param name="targetBandKind">The target BandKind to create a Band</param>
        Private Sub ProcessSubDocument(ByVal server As SnapDocumentServer, ByVal targetBandKind As BandKind)
            'IMPORTANT! In Snap, you can place page fields at any position (even in formatted text)
            'XRPageInfo does not support this, so all fields are converted to static values before processing
            server.Document.UnlinkAllFields()
            'Specify the current band to place content into
            generator.CurrentBand = generator.CurrentReport.Bands.Create(targetBandKind)
            'Create a SnapLayoutProcessor instance to process the document layout
            Dim layoutProcessor As SnapLayoutProcessor = New SnapLayoutProcessor(generator, server)
            layoutProcessor.ProcessDocument()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ConfigureDataConnectionEvent = Nothing
            generator = Nothing
        End Sub

        Public Event ConfigureDataConnection As ConfigureDataConnectionEventHandler
    End Class
End Namespace
