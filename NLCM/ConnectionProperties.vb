Imports System.Data.OleDb

Namespace com.data
    Public Class ConnectionsProperty
        Dim _Name As String
        Dim _ConectionString As String
        Dim _Connection As OleDbConnection
        Dim _Adap As OleDbDataAdapter
        Dim _Command As OleDbCommand

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                _Name = value
            End Set
        End Property

        Public Property ConectionString() As String
            Get
                Return _ConectionString
            End Get
            Set(ByVal value As String)
                _ConectionString = value
                Try
                    SetUpConnection()
                Catch ex As Exception

                End Try
            End Set
        End Property

        Public ReadOnly Property Connection() As OleDbConnection
            Get
                Return _Connection
            End Get
        End Property

        Public Property Adap() As OleDbDataAdapter
            Get
                Return _Adap
            End Get
            Set(ByVal value As OleDbDataAdapter)
                _Adap = value
            End Set
        End Property

        Public Property Command() As OleDbCommand
            Get
                Return _Command
            End Get
            Set(ByVal value As OleDbCommand)
                _Command = value
            End Set
        End Property

        Public Sub SetUpConnection()
            _Connection = New OleDbConnection(_ConectionString)
        End Sub

    End Class
End Namespace