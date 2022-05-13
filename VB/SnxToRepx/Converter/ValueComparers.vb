Imports DevExpress.XtraRichEdit.API.Layout
Imports System.Collections.Generic

Namespace SnxToRepx.Converter

    ''' <summary>
    ''' This is a comparer to sort a collection of RangedLayoutElements
    ''' </summary>
    Friend Class RangedElementComparer
        Implements IComparer(Of RangedLayoutElement)

        ''' <summary>
        ''' Compares two RangedLayoutElements by their ranges
        ''' </summary>
        ''' <param name="x">The first element</param>
        ''' <param name="y">The second element</param>
        ''' <returns>The result of comparison</returns>
        Public Function Compare(ByVal x As RangedLayoutElement, ByVal y As RangedLayoutElement) As Integer Implements IComparer(Of RangedLayoutElement).Compare
            'Use the default integer comparer to compare start positions for elements
            Return Comparer(Of Integer).Default.Compare(x.Range.Start, y.Range.Start)
        End Function
    End Class
End Namespace
