Public Class IPStatisticsObj

    Private _IDRec As Integer = -1
    Private _IP_ADDRESS As String
    Private _DATE_RECORD As Date
    Private _MBCONSUMPTION_IN As Long
    Private _MBCONSUMPTION_OUT As Long
    Private _SESSIONUSERLOGED As String

    Public Property IDRec As Integer
        Get
            Return _IDRec
        End Get
        Set(value As Integer)
            _IDRec = value
        End Set
    End Property

    Public Property IP_ADDRESS As String
        Get
            Return _IP_ADDRESS
        End Get
        Set(value As String)
            _IP_ADDRESS = value
        End Set
    End Property

    Public Property DATE_RECORD As Date
        Get
            Return _DATE_RECORD
        End Get
        Set(value As Date)
            _DATE_RECORD = value
        End Set
    End Property

    Public Property MBCONSUMPTION_IN As Long
        Get
            Return _MBCONSUMPTION_IN
        End Get
        Set(value As Long)
            _MBCONSUMPTION_IN = value
        End Set
    End Property
    Public Property MBCONSUMPTION_OUT As Long
        Get
            Return _MBCONSUMPTION_OUT
        End Get
        Set(value As Long)
            _MBCONSUMPTION_OUT = value
        End Set
    End Property

    Public Property SESSIONUSERLOGED() As String
        Get
            Return _SESSIONUSERLOGED

        End Get
        Set(value As String)
            _SESSIONUSERLOGED = value
        End Set
    End Property





End Class
