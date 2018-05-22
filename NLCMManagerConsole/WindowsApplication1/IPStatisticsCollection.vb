
Public Class IPStatisticsCollection
    Implements IEnumerable, ICollection, IEnumerator


    Private position As Integer = -1
    Private _Items As New List(Of IPStatisticsObj)

    ReadOnly Property Items() As List(Of IPStatisticsObj)
        Get
            Return _Items
        End Get
    End Property

    Public Sub Add(item As IPStatisticsObj)
        If Not IsNothing(item) Then
            _Items.Add(item)
        End If
    End Sub

    Public Sub CopyTo(array As Array, index As Integer) Implements ICollection.CopyTo
        array = _Items.ToArray
    End Sub

    Public ReadOnly Property Count As Integer Implements ICollection.Count
        Get
            Return _Items.Count
        End Get
    End Property

    Public ReadOnly Property IsSynchronized As Boolean Implements ICollection.IsSynchronized
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property SyncRoot As Object Implements ICollection.SyncRoot
        Get
            Return Me
        End Get
    End Property

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return CType(Me, IEnumerator)
    End Function

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            Return _Items.GetEnumerator.Current
        End Get
    End Property

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        position += 1
        Return (position)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        position = -1
    End Sub
End Class
