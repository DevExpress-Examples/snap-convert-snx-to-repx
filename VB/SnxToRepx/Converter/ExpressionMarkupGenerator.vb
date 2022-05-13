Imports DevExpress.Snap.Core.API

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This class generates an extended markup, which can be used in ExpressionBindings
    ''' </summary>
    Friend Class ExpressionMarkupGenerator
        Inherits MarkupGenerator

        Public Sub New(ByVal generator As ReportGenerator, ByVal document As SnapDocument)
            MyBase.New(generator, document)
        End Sub

        ''' <summary>
        ''' Processes a field and returns a corresponding Expression part
        ''' </summary>
        ''' <param name="entity">A Snap entity to process</param>
        Protected Overrides Sub ProcessField(ByVal entity As SnapSingleListItemEntity)
            'IMPORTANT! This approach is required to properly process summaries. Background: the converter uses HTML Markup
            'available in reports to format content. Standard fields are passed as mail-merge fields. Since it is impossible to
            'specify summaries in the standard mail-merge report, we need to construct an Expression instead of standard markup.
            'It uses a slightly different mechanism. Specifically, static string content should be wrapped with '. Dynamic content (fields)
            'should be outside static strings, and we need to use + to concatenate different parts. For example:
            'Static markup: <p align=center>[UnitPrice]</p>
            'Expression: '<p align=center>' + [UnitPrice] + '</p>'
            'Add the closing apostrophe and the plus sign (see the example above)
            Buffer.Append("' + ")
            'Check if it is a SnapText instance. SnapText contains format strings, summary settings, etc.
            Dim snapText As SnapText = TryCast(entity, SnapText)
            If snapText Is Nothing Then
                'If not, generate the standard field representation ( [fieldName] )
                MyBase.ProcessField(entity)
            Else
                'If so, get SnapText related properties
                Dim dataFieldName As String = snapText.DataFieldName
                Dim formatString As String = "{0:" & snapText.FormatString & "}"
                Dim isParameter As Boolean = snapText.IsParameter
                'if it is a parameter, add the ? sign (used in XtraReports to determine whether a control is bound to a parameter)
                If isParameter Then
                    dataFieldName = String.Format("?{0}", dataFieldName)
                Else
                    'If not, check if it contains summary information
                    Dim summaryFunc As String = SummaryConverter.Convert(snapText.SummaryFunc)
                    If Not Equals(summaryFunc, String.Empty) Then
                        'Add a corresponding expression function (sumSum, sumCount, etc.)
                        dataFieldName = String.Format("{0}([{1}])", summaryFunc, dataFieldName)
                    Else
                        'If no summary added, use the standard field format
                        dataFieldName = String.Format("[{0}]", dataFieldName)
                    End If
                End If

                'Check if the SnapText has a format string
                If String.IsNullOrEmpty(snapText.FormatString) Then
                    Buffer.Append(dataFieldName)
                Else
                    'Use the FormatString expression function to format the resulting value
                    Buffer.AppendFormat("FormatString('{0}', {1})", formatString, dataFieldName)
                End If
            End If

            'Add an opening apostrophe
            Buffer.Append(" + '")
        End Sub

        'Some tags in markup use ' for strings (e.g., font name). To properly use it in expressions,
        'we need to return a double apostrophe.
        Public Overrides ReadOnly Property FontValueWrapperSymbol As String
            Get
                Return "''"
            End Get
        End Property

        'Returns the expression with markup
        Public Overrides ReadOnly Property Text As String
            Get
                Return String.Format("'{0}'", MyBase.Text)
            End Get
        End Property
    End Class
End Namespace
