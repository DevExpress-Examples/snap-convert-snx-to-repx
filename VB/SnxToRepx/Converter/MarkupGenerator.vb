Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' A base class to generate content markup
    ''' </summary>
    Friend MustInherit Class BufferedDocumentVisitor
        Inherits DocumentVisitorBase

        Private ReadOnly bufferField As StringBuilder 'A buffer containing the resulting markup

        Protected Sub New()
            bufferField = New StringBuilder()
        End Sub

        Protected ReadOnly Property Buffer As StringBuilder
            Get
                Return bufferField
            End Get
        End Property
    End Class

    ''' <summary>
    ''' A class used to convert formatted text from Snap to markup supported in XtraReports
    ''' </summary>
    Friend Class MarkupGenerator
        Inherits BufferedDocumentVisitor

        Const lastLowSpecial As Char = Microsoft.VisualBasic.Strings.ChrW(31) 'lower boundary for special symbols

        Const firstHighSpecial As Char = Microsoft.VisualBasic.Strings.ChrW(65535) 'higher boundary for special symbols

        Private document As SnapDocument 'The source document

        Private processedFields As List(Of Field) 'A list containing already processed fields

        Private generator As ReportGenerator 'Contains the resulting report

        Public Sub New(ByVal generator As ReportGenerator, ByVal document As SnapDocument)
            Me.document = document
            Me.generator = generator
            processedFields = New List(Of Field)()
        End Sub

        ''' <summary>
        ''' Processes plain text
        ''' </summary>
        ''' <param name="text">DocumentText containing information about the current text portion</param>
        Public Overrides Sub Visit(ByVal text As DocumentText)
            'Get all opening tags
            Dim tags As List(Of MarkupTag) = GetTags(text.TextProperties)
            'Add them to the resulting markup
            For i As Integer = 0 To tags.Count - 1
                Buffer.Append(tags(i).OpeningTag)
            Next

            'Check if this text portion contains fields
            Dim entity As SnapSingleListItemEntity = GetSnapSingleItemField(text.Position)
            If entity IsNot Nothing Then
                'If the field is not yet processed (added to markup), process it
                If Not processedFields.Contains(entity.Field) Then
                    processedFields.Add(entity.Field)
                    ProcessField(entity)
                End If
            Else
                'Add text char by char
                Dim count As Integer = text.Length
                For i As Integer = 0 To count - 1
                    Dim ch As Char = text.Text(i)
                    If ch > lastLowSpecial AndAlso ch < firstHighSpecial Then
                        Buffer.Append(ch)
                    ElseIf ch = Microsoft.VisualBasic.Strings.ChrW(9) OrElse ch = Microsoft.VisualBasic.Strings.ChrW(10) OrElse ch = Microsoft.VisualBasic.Strings.ChrW(13) Then
                        Buffer.Append(ch)
                    End If
                Next
            End If

            'Add closing tags to the resulting markup
            For i As Integer = tags.Count - 1 To 0 Step -1
                Buffer.Append(tags(i).ClosingTag)
            Next
        End Sub

        ''' <summary>
        ''' Returns a markup for a Snap entity
        ''' </summary>
        ''' <param name="entity">A Snap entity to process</param>
        Protected Overridable Sub ProcessField(ByVal entity As SnapSingleListItemEntity)
            'In the standard markup, fields should look like this: [UnitPrice!$0.00]
            'Get the data field name
            Dim dataFieldName As String = entity.DataFieldName
            'Add an opening square bracket
            Buffer.Append("[")
            'If it is a parameter, it should start with ? (for example, [?parameter1])
            If TypeOf entity Is SnapText AndAlso CType(entity, SnapText).IsParameter Then Buffer.Append("?")
            'Append the field name
            Buffer.Append(dataFieldName)
            'Check if the format string is specified
            'Add the format string to the field
            If TypeOf entity Is SnapText AndAlso Not String.IsNullOrEmpty(CType(entity, SnapText).FormatString) Then Buffer.AppendFormat("!{0}", CType(entity, SnapText).FormatString)
            'Add a closing square bracket
            Buffer.Append("]")
        End Sub

        ''' <summary>
        ''' Visits the paragraph start
        ''' </summary>
        ''' <param name="paragraphStart">An object containing paragraph properties</param>
        Public Overrides Sub Visit(ByVal paragraphStart As DocumentParagraphStart)
            'Add an opening paragraph tag with alignment settings
            Buffer.AppendFormat("<p {0}>", AlignmentConverter.Convert(paragraphStart.ParagraphProperties.Alignment))
        End Sub

        ''' <summary>
        ''' Visits the paragraph end
        ''' </summary>
        ''' <param name="paragraphEnd">An object containing paragraph properties</param>
        Public Overrides Sub Visit(ByVal paragraphEnd As DocumentParagraphEnd)
            'Add a closing paragraph tag
            Buffer.Append("</p>")
        End Sub

        ''' <summary>
        ''' Visits the hyperlink start
        ''' </summary>
        ''' <param name="hyperlinkStart">An object containing information about hyperlinks</param>
        Public Overrides Sub Visit(ByVal hyperlinkStart As DocumentHyperlinkStart)
            'Add a hyperlink tag
            Buffer.AppendFormat("<href value={0}>", hyperlinkStart.NavigateUri)
        End Sub

        ''' <summary>
        ''' Visits the hyperlink end
        ''' </summary>
        ''' <param name="hyperlinkEnd">An object containing information about hyperlinks</param>
        Public Overrides Sub Visit(ByVal hyperlinkEnd As DocumentHyperlinkEnd)
            'Add a closing hyperlink tag
            Buffer.Append("</href>")
        End Sub

        ''' <summary>
        ''' Checks if a Snap entity exists at the specified position
        ''' </summary>
        ''' <param name="position">A position to check</param>
        ''' <returns>A SnapSingleListItemEntity instance</returns>
        Private Function GetSnapSingleItemField(ByVal position As Integer) As SnapSingleListItemEntity
            'Iterate through all fields in the document and check if the specified position is in the field's code range
            For Each field As Field In document.Fields
                If field.Range.Start.ToInt() <= position AndAlso field.CodeRange.End.ToInt() >= position Then
                    'If so, parse a field and check if it is SnapSingleListItemEntity (can be bound to a data field)
                    Dim entity As SnapEntity = document.ParseField(field)
                    If TypeOf entity Is SnapSingleListItemEntity Then
                        'Return the found entity
                        Return CType(entity, SnapSingleListItemEntity)
                    End If
                End If
            Next

            'If no field exists, return Null
            Return Nothing
        End Function

        ''' <summary>
        ''' Collects tags based on text properties
        ''' </summary>
        ''' <param name="properties">An object containing information about formatting</param>
        ''' <returns>A collection of tag objects</returns>
        Private Function GetTags(ByVal properties As ReadOnlyTextProperties) As List(Of MarkupTag)
            'Create a list to store tag objects
            Dim list As List(Of MarkupTag) = New List(Of MarkupTag)()
            'Create a font tag
            Dim fontTag As MarkupTag = New MarkupTag("font")
            list.Add(fontTag)
            'Add a font name to the font tag
            Dim fontName As String = properties.FontName
            fontTag.Parts.Add(String.Format("={0}{1}{0}", FontValueWrapperSymbol, fontName))
            'Add a font size to the font tag
            Dim fontSize As Single = properties.FontSize
            fontTag.Parts.Add(String.Format(" size={0:0.0}", fontSize))
            'Add a font color to the font tag
            Dim foreColor As Color = properties.ForeColor
            fontTag.Parts.Add(String.Format(" color={0}", GetColorString(foreColor)))
            'Add a highlight color to the font tag
            Dim backColor As Color = properties.HighlightColor
            'Get the string representation of the color
            Dim colorString As String = GetColorString(backColor)
            If Not Equals(colorString, String.Empty) Then fontTag.Parts.Add(String.Format(" backcolor={0}", colorString))
            'If text is bold, add the corresponding tag
            If properties.FontBold Then list.Add(New MarkupTag("b"))
            'If text is italic, add the corresponding tag
            If properties.FontItalic Then list.Add(New MarkupTag("i"))
            'If text is underlined, add the corresponding tag
            If properties.UnderlineType <> UnderlineType.None Then list.Add(New MarkupTag("u"))
            'If text is of the strikeout style, add the corresponding tag
            If properties.StrikeoutType <> StrikeoutType.None Then list.Add(New MarkupTag("s"))
            'Return the list of tags
            Return list
        End Function

        ''' <summary>
        ''' Converts System.Drawing.Color to HTML representation
        ''' </summary>
        ''' <param name="color">The source color</param>
        ''' <returns>HTML representation of the specified color</returns>
        Private Function GetColorString(ByVal color As Color) As String
            Return ColorTranslator.ToHtml(color)
        End Function

        'The font name wrapper symbol
        Public Overridable ReadOnly Property FontValueWrapperSymbol As String
            Get
                Return "'"
            End Get
        End Property

        'The resulting markup
        Public Overridable ReadOnly Property Text As String
            Get
                Return String.Format("{0}</p>", Buffer.ToString())
            End Get
        End Property

        ''' <summary>
        ''' This class stores information about tags used in a specific text portion
        ''' </summary>
        Private Class MarkupTag

            'The base tag without any attributes
            Private baseTag As String

            'Additional parts (attributes, values, etc.)
            Private partsField As List(Of String)

            Public Sub New(ByVal baseTag As String)
                Me.baseTag = baseTag
                partsField = New List(Of String)()
            End Sub

            ''' <summary>
            ''' Generates an opening tag based on the base tag and additional parts
            ''' </summary>
            ''' <returns>A string containing a full opening tag</returns>
            Private Function GetOpeningTag() As String
                'Create a string builder to construct a string
                Dim sb As StringBuilder = New StringBuilder("<")
                'Add a base tag
                sb.Append(baseTag)
                'Append additional parts
                For i As Integer = 0 To partsField.Count - 1
                    sb.Append(partsField(i))
                Next

                'Close the tag
                sb.Append(">")
                'Return the generated value
                Return sb.ToString()
            End Function

            Public ReadOnly Property Parts As List(Of String)
                Get
                    Return partsField
                End Get
            End Property 'A list of additional parts

            Public ReadOnly Property OpeningTag As String
                Get
                    Return GetOpeningTag()
                End Get
            End Property 'An opening tag

            Public ReadOnly Property ClosingTag As String
                Get
                    Return String.Format("</{0}>", baseTag)
                End Get
            End Property 'A closing tag
        End Class
    End Class
End Namespace
