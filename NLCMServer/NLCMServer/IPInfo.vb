Public Class IPInfo

    Private _IDIP As Integer = -1
    Private _IP_ADDRESS As String
    Private _LOCATION As String
    Private _ONLINE As Boolean
    Private _ACTIVE As Boolean
    Private _MONITOR As Boolean
    Private _Interval_REC As Integer
    Private _ADAPTERNAME As String
    Private _MACADDRESS As String

    Public Property MACADDRESS As String
        Get
            Return _MACADDRESS
        End Get
        Set(value As String)
            _MACADDRESS = value
        End Set
    End Property
    Public Property ADAPTERNAME As String
        Get
            Return _ADAPTERNAME
        End Get
        Set(value As String)
            _ADAPTERNAME = value
        End Set
    End Property

    Public Property Interval_REC As Integer
        Get
            Return _Interval_REC
        End Get
        Set(value As Integer)
            _Interval_REC = value
        End Set
    End Property
    Public Property IDIP As Integer
        Get
            Return _IDIP

        End Get
        Set(value As Integer)
            _IDIP = value
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

    Public Property LOCATION As String
        Get
            Return _LOCATION
        End Get
        Set(value As String)
            _LOCATION = value
        End Set
    End Property
    Public Property ONLINE As Boolean
        Get
            Return _ONLINE
        End Get
        Set(value As Boolean)
            _ONLINE = value
        End Set
    End Property

    Public Property ACTIVE As Boolean
        Get
            Return _ACTIVE
        End Get
        Set(value As Boolean)
            _ACTIVE = value
        End Set
    End Property

    Public Property MONITOR As Boolean
        Get
            Return _MONITOR
        End Get
        Set(value As Boolean)
            _MONITOR = value
        End Set
    End Property

End Class
