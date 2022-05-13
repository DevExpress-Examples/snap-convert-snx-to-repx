Imports DevExpress.Snap.Core.API
Imports DevExpress.XtraRichEdit.API.Native
Imports System.Collections.Generic

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This is a collection of SnapList fields with additional logic
    ''' </summary>
    Friend Class SnapListCollection
        Inherits List(Of Field)

        Private source As SnapDocument 'The source document

        Public Sub New(ByVal source As SnapDocument)
            Me.source = source
        End Sub

        ''' <summary>
        ''' This method collects top-level SnapLists
        ''' </summary>
        Public Sub PrepareCollection()
            'Clear the collection
            Clear()
            'Iterate through fields
            For Each field As Field In source.Fields
                'Get a top-level field for the current field
                Dim topField As Field = GetTopLevelField(field)
                If Not ContainsField(topField) Then
                    'If the field is not yet added and it is a SnapList, add it to the collection
                    Dim entity As SnapEntity = source.ParseField(topField)
                    If entity IsNot Nothing AndAlso TypeOf entity Is SnapList Then Add(topField)
                End If
            Next

            'Sort lists based on their ranges (SnapLists can be in the reverse order in the document model)
            Sort(New FieldComparer())
        End Sub

        ''' <summary>
        ''' Gets a SnapList instance based on its index in the collection
        ''' </summary>
        ''' <param name="index"></param>
        ''' <returns></returns>
        Public Function GetSnapList(ByVal index As Integer) As SnapList
            'Get and parse the source field
            Dim field As Field = Me(index)
            Dim entity As SnapEntity = source.ParseField(field)
            'Return SnapList
            Return CType(entity, SnapList)
        End Function

        ''' <summary>
        ''' Gets a SnapList instance containing the specified position
        ''' </summary>
        ''' <param name="position">A position to check</param>
        ''' <returns>The SnapList</returns>
        Public Function GetSnapListByPosition(ByVal position As Integer) As SnapList
            'Iterate through elements in the collection
            For i As Integer = 0 To Count - 1
                'Get a field and compare its range with the specified position
                Dim field As Field = Me(i)
                'If they match, return SnapList
                If field.Range.Start.ToInt() <= position AndAlso field.Range.End.ToInt() >= position Then Return CType(source.ParseField(field), SnapList)
            Next

            'If no SnapList is found, return Null
            Return Nothing
        End Function

        ''' <summary>
        ''' Gets the top-level field for the current field
        ''' </summary>
        ''' <param name="field">The currently processed field</param>
        ''' <returns>The top-level field</returns>
        Private Function GetTopLevelField(ByVal field As Field) As Field
            'Access the parent field
            Dim parentField As Field = field
            While parentField.Parent IsNot Nothing
                parentField = parentField.Parent
            End While

            'Return the resulting field
            Return parentField
        End Function

        ''' <summary>
        ''' Checks whether the collection contains a specific field
        ''' </summary>
        ''' <param name="field">A field to locate</param>
        ''' <returns>True if the collection contains the field. Otherwise, false.</returns>
        Public Function ContainsField(ByVal field As Field) As Boolean
            'Iterate through fields in the collection
            For i As Integer = 0 To Count - 1
                'Get a field and compare its range with the specified position
                Dim addedField As Field = Me(i)
                If addedField.Range.Start.ToInt() = field.Range.Start.ToInt() AndAlso addedField.Range.Start.ToInt() = field.Range.Start.ToInt() Then Return True
            Next

            Return False
        End Function
    End Class

    ''' <summary>
    ''' A comparer used to sort fields by their positions
    ''' </summary>
    Friend Class FieldComparer
        Implements IComparer(Of Field)

        Public Function Compare(ByVal x As Field, ByVal y As Field) As Integer Implements IComparer(Of Field).Compare
            'Use the default Integer comparer to compare positions
            Return Comparer(Of Integer).Default.Compare(x.Range.Start.ToInt(), y.Range.Start.ToInt())
        End Function
    End Class
End Namespace
