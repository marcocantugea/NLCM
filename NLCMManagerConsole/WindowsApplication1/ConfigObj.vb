Public Class ConfigObj

    Private _IDConfig As Integer = -1
    Private _parameter As String
    Private _param_val As String
    Private _active As Boolean


    Public Property IDConfig As Integer
        Get
            Return _IDConfig
        End Get
        Set(value As Integer)
            _IDConfig = value
        End Set
    End Property

    Public Property parameter As String
        Get
            Return _parameter
        End Get
        Set(value As String)
            _parameter = value
        End Set
    End Property

    Public Property param_val As String
        Get
            Return _param_val
        End Get
        Set(value As String)
            _param_val = value
        End Set
    End Property
    Public Property active As Boolean
        Get
            Return _active
        End Get
        Set(value As Boolean)
            _active = value
        End Set
    End Property

End Class
