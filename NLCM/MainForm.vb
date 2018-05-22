Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Imports System.Runtime
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports Microsoft.Win32


Public Class MainForm
    'General timer for statictis
    Private WithEvents _timer As New Timer
    'Timer for internet connectivity connection
    Private WithEvents _timerInternetConnection As New Timer
    Private _startTime As DateTime = DateTime.MinValue
    Private _isNetworkOnline As Boolean
    Dim properties As IPGlobalProperties = IPGlobalProperties.GetIPGlobalProperties
    Dim ipstat As IPGlobalStatistics = Nothing
    Dim start_r_packets As Decimal
    Dim end_r_packets As Decimal
    Dim fNetworkInterfaces() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
    Dim start_received_bytes As Long
    Dim start_sent_bytes As Long
    Dim end_received_bytes As Long
    Dim end_sent_bytesas As Long
    Dim _intervalTimerTick As Long
    Dim _PCInfo As New IPStatisticsObj
    Private _EnableMonitor As Boolean = True
    Dim _ADO As New ADO
    Dim _Adaptername As String
    Dim _appconfig As New AppConfigFileSettings
    Dim _localconnection As String = System.Configuration.ConfigurationSettings.AppSettings("online")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Config As New ConfigObj
        Config.parameter = "-7"
        Config.param_val = "-7"
        Config.active = True
        Config.IDConfig = -7

        Try
            _ADO.GetParameter(Config, "dnsloop")
        Catch ex As Exception
            ''TODO: Log
            'TODO : Remove hard code
            Config.parameter = "dnsloop"
            Config.param_val = "127.0.0.0"
        End Try

        'Label1.Text = Config.parameter & " = " & Config.param_val

        ConnectInternet()
    End Sub

    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)
        If Not Me.IsHandleCreated Then
            Me.CreateHandle()
            value = False
        End If

        'set timer to minutes
        _timer.Interval = 1000
        _timerInternetConnection.Interval = 180000
        'initialize timers
        _timer.Start()
        _timerInternetConnection.Start()

        'initilizar monitor of network
        StartMonitor()
        ''Label1.Text = "..."

        MyBase.SetVisibleCore(value)
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
       

    End Sub


    'main time ticker
    Private Sub TimerTick(sender As Object, e As EventArgs) Handles _timer.Tick
        'get the interval to save the record and convert into minutes
        Dim intervaltoRecord As Long = GetIntervalRecord() * 60
        'checks if the timer meets the interval configured in the app.conf
        If _intervalTimerTick = intervaltoRecord Then
            ''Label1.Text = "entro"
            'Set tie interval to 0
            _intervalTimerTick = 0
            'Check is the Monitoring is enable
            CheckMonitoringEngine()
            ' if is enable will record the data
            If _EnableMonitor Then
                CloseEverything(Nothing, Nothing)
            End If

        Else
            If _intervalTimerTick = 30 Then
                'Label1.Text = "..."
            End If
            _intervalTimerTick += 1
        End If

    End Sub

    'Timer to check configuration on the database
    Private Sub timerInternetConnectionTick(sender As Object, e As EventArgs) Handles _timerInternetConnection.Tick
        'Get the database online configuration
        Dim online As Boolean
        Try
            online = _ADO.GetOnlineParameter(getIPAddr)
        Catch ex As Exception
            ''TODO: lOG
            ''TODO: REMOVE HARD CODE
            online = True
        End Try
        ' _localconnection = System.Configuration.ConfigurationSettings.AppSettings("online")
        'if database field is false will check the local configuration
        If Not online Then
            'if the local configuration is 1 (connected) will procedure to disconnect
            If _localconnection.Equals("1") Then
                'change dns to a internal loop
                DisconnecInternet()
                'change the parameter in the local configuration
                _appconfig.UpdateAppSettings("online", "0")
                _localconnection = "0"
                'Label1.Text = "Disconnected from internet..."
            End If
        Else
            'if the local configuration is 0(disconected) will procedure to connect internet
            If _localconnection.Equals("0") Then
                'change dns for internet dns
                ConnectInternet()
                'change the parameter in the local configuration
                _appconfig.UpdateAppSettings("online", "1")
                _localconnection = "1"
                'Label1.Text = "Connected to internet..."
            End If
        End If
    End Sub

#Region "Network statictis"

    'Function to get the configuration of the app.setting about the interval time in minutes that will record
    'the consumption the parameter is "Interval-REC"
    Private Function GetIntervalRecord() As Long
        ' variable tu return
        Dim IntervalRecord As Long = 0
        'get the configuration from the database
        Dim _ADO_IntervalRec As Integer
        Try
            _ADO_IntervalRec = _ADO.GetIntervalRec(getIPAddr)
        Catch ex As Exception
            ''TODO: Log
            ''TODO : Remove hard corde
            _ADO_IntervalRec = 1
        End Try
        'if the database configuration is set will set the return variable
        If _ADO_IntervalRec > 0 Then
            IntervalRecord = _ADO_IntervalRec
        Else
            'if the variable is not set will take the app.config to setup the interval record
            ' the configuration variable is Interval-Rec 
            Dim Setting_intervalrecord As String = System.Configuration.ConfigurationSettings.AppSettings("Interval-Rec")
            If Not IsNothing(Setting_intervalrecord) Then
                Try
                    IntervalRecord = Long.Parse(Setting_intervalrecord)
                Catch ex As Exception
                    IntervalRecord = 5
                End Try
            End If
        End If


        Return IntervalRecord
    End Function

    'function to check the parameter MONITOR in the database to record the consumption
    Private Sub CheckMonitoringEngine()
        'get the ip adress for check the parameter monitor on the database
        Dim checkmonitor As Boolean = _ADO.CheckmMonitorIsEnable(getIPAddr)
        If Not checkmonitor Then
            _EnableMonitor = False
        Else
            _EnableMonitor = True
        End If
    End Sub

    'function to get the actual ip address of the host
    Private Function getIPAddr() As String
        'get the ip adress for check the parameter monitor on the database
        Dim myHost As String = System.Net.Dns.GetHostName
        Dim ipEntry As IPHostEntry = System.Net.Dns.GetHostEntry(myHost)
        Dim addr As IPAddress() = ipEntry.AddressList
        Return addr(addr.Length - 1).ToString
    End Function

    'function to inizialize the network monitoring 
    Private Sub StartMonitor()
        ipstat = properties.GetIPv4GlobalStatistics()
        GetConnectionInfo()
    End Sub
    'function to get the statistics of the network will start count the MB and set it  up on globals vabriables
    ' also will add the ip address on the mail table configuration
    Public Sub GetConnectionInfo()

        Try
            Dim myHost As String = System.Net.Dns.GetHostName
            Dim ipEntry As IPHostEntry = System.Net.Dns.GetHostEntry(myHost)
            Dim addr As IPAddress() = ipEntry.AddressList
            AddHandler NetworkChange.NetworkAddressChanged, AddressOf NetworkChange_NetworkAvailabilityChanged
            _isNetworkOnline = NetworkInterface.GetIsNetworkAvailable()
            If addr.Length > 0 Then
                'obtiene la ip
                start_r_packets = Convert.ToDecimal(ipstat.ReceivedPackets)
                _PCInfo.IP_ADDRESS = addr(addr.Length - 1).ToString
                _PCInfo.SESSIONUSERLOGED = Environment.UserName
                Dim adapter As NetworkInterface = fNetworkInterfaces(0)
                _Adaptername = adapter.Name
                start_received_bytes = fNetworkInterfaces(0).GetIPv4Statistics.BytesReceived
                start_sent_bytes = fNetworkInterfaces(0).GetIPv4Statistics.BytesSent

                start_received_bytes = (start_received_bytes / 1048576 * 100000) / 100000
                start_sent_bytes = (start_sent_bytes / 1048576 * 100000) / 100000

            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    'function to make the cut of the MB captured in the time of the function GetConnectionInfo() was invoke
    ' also save the record of that consumption on the database and if the ip doesn't not exist on the configuration table
    ' will be added whit the default info
    Public Sub CloseEverything(sender As Object, e As EventArgs)
        _timer.Stop()
        ipstat = properties.GetIPv4GlobalStatistics
        end_r_packets = Convert.ToDecimal(ipstat.ReceivedPackets)
        _PCInfo.DATE_RECORD = Date.Now
        end_r_packets = end_r_packets - start_r_packets
        end_received_bytes = fNetworkInterfaces(0).GetIPv4Statistics.BytesReceived
        end_sent_bytesas = fNetworkInterfaces(0).GetIPv4Statistics.BytesSent
        end_received_bytes = (end_received_bytes / 1048576 * 100000) / 100000
        end_sent_bytesas = (end_sent_bytesas / 1048576 * 100000) / 100000

        _PCInfo.MBCONSUMPTION_IN = end_received_bytes - start_received_bytes
        _PCInfo.MBCONSUMPTION_OUT = end_sent_bytesas - start_sent_bytes

        _intervalTimerTick = 0

        'Record on database
        Try
            'save record in database
            _ADO.RecordIP(_PCInfo)
            'check if the ip address if is registered on the table
            Dim chek_ip As Boolean = _ADO.CheckIPExist(_PCInfo.IP_ADDRESS)
            If Not chek_ip Then
                'fill new IPInfo object to default values and save it on the data base
                Dim new_ipinfo As New IPInfo
                new_ipinfo.IP_ADDRESS = _PCInfo.IP_ADDRESS
                new_ipinfo.LOCATION = "pending"
                new_ipinfo.MONITOR = True
                new_ipinfo.ONLINE = True
                new_ipinfo.ACTIVE = True
                new_ipinfo.Interval_REC = 1
                new_ipinfo.ADAPTERNAME = _Adaptername
                _ADO.RecordIPInfo(new_ipinfo)
            End If
            'Label1.Text = "State Saved."
        Catch ex As Exception
            ''TODO: LOG
            ''TODO: make a buffer transacciton to save info
        End Try


        _timer.Start()
    End Sub

    Private Sub NetworkChange_NetworkAvailabilityChanged(sender As Object, e As NetworkAvailabilityEventArgs)
        _isNetworkOnline = e.IsAvailable
    End Sub

    <System.Runtime.InteropServices.DllImport("wininet.dll")>
    Public Shared Function InternetGetConnectedState(ByVal description As Integer, ReservedValue As Integer)

    End Function

#End Region

#Region "Change Internet Config"

    Private Sub ConnectInternet()
        Dim applychanges As Boolean = True

        Dim dnsloop As New ConfigObj
        dnsloop.parameter = "-7"
        dnsloop.param_val = "-7"
        dnsloop.active = True
        dnsloop.IDConfig = -7
        Try
            _ADO.GetParameter(dnsloop, "dnsloop")
        Catch ex As Exception
            ''TODO: Log
            applychanges = False
        End Try


        Dim dns1 As New ConfigObj
        dns1.parameter = "-7"
        dns1.param_val = "-7"
        dns1.active = True
        dns1.IDConfig = -7
        Try
            _ADO.GetParameter(dns1, "dns1")
        Catch ex As Exception
            ''TODO: Log
            applychanges = False
        End Try


        Dim dns2 As New ConfigObj
        dns2.parameter = "-7"
        dns2.param_val = "-7"
        dns2.active = True
        dns2.IDConfig = -7
        Try
            _ADO.GetParameter(dns2, "dns2")
        Catch ex As Exception
            ''TODO: Log
            applychanges = False
        End Try


        Dim dns3 As New ConfigObj
        dns3.parameter = "-7"
        dns3.param_val = "-7"
        dns3.active = True
        dns3.IDConfig = -7
        Try
            _ADO.GetParameter(dns3, "dns3")
        Catch ex As Exception
            ''TODO: Log
            applychanges = False
        End Try

        If applychanges Then
            'delete internet dns
            DeleteDNS(dnsloop.param_val, _Adaptername)
            'add loop dns as primary 
            AddDNS(dns2.param_val, _Adaptername, "")
            AddDNS(dns3.param_val, _Adaptername, "")
            Console.WriteLine("Internet Connected...")
        End If

    End Sub

    Private Sub DisconnecInternet()
        Dim applychanges As Boolean = True
        Dim dnsloop As New ConfigObj
        dnsloop.parameter = "-7"
        dnsloop.param_val = "-7"
        dnsloop.active = True
        dnsloop.IDConfig = -7
        Try
            _ADO.GetParameter(dnsloop, "dnsloop")
        Catch ex As Exception
            ''TODO: log
            applychanges = False
        End Try


        Dim dns1 As New ConfigObj
        dns1.parameter = "-7"
        dns1.param_val = "-7"
        dns1.active = True
        dns1.IDConfig = -7
        Try
            _ADO.GetParameter(dns1, "dns1")
        Catch ex As Exception
            ''TODO: log
            applychanges = False
        End Try


        Dim dns2 As New ConfigObj
        dns2.parameter = "-7"
        dns2.param_val = "-7"
        dns2.active = True
        dns2.IDConfig = -7
        Try
            _ADO.GetParameter(dns2, "dns2")
        Catch ex As Exception
            ''TODO: log
            applychanges = False
        End Try


        Dim dns3 As New ConfigObj
        dns3.parameter = "-7"
        dns3.param_val = "-7"
        dns3.active = True
        dns3.IDConfig = -7
        Try
            _ADO.GetParameter(dns3, "dns3")
        Catch ex As Exception
            ''TODO: log
            applychanges = False
        End Try

        If applychanges Then
            'add loop dns as primary 
            AddDNS(dnsloop.param_val, _Adaptername, "1")
            'delete internet dns
            DeleteDNS(dns2.param_val, _Adaptername)
            DeleteDNS(dns3.param_val, _Adaptername)
            Console.WriteLine("Internet Disconnected...")
        End If

    End Sub

    Private Sub AddDNS(dnsaddress As String, networkname As String, primary As String)
       
        Dim command As String = "interface ip add dns name=""" & networkname & """ address=""" & dnsaddress & """ " & primary
        Dim pr As New Process

        pr.StartInfo.FileName = "netsh.exe"
        pr.StartInfo.CreateNoWindow = True
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardInput = True
        pr.StartInfo.RedirectStandardOutput = True
        pr.Start()
        Dim wr As System.IO.StreamWriter = pr.StandardInput
        Dim rr As System.IO.StreamReader = pr.StandardOutput

        wr.WriteLine(command)
        wr.Flush()
        wr.Close()
        pr.WaitForExit()
        pr.Close()

    End Sub

    Private Sub DeleteDNS(dnsaddress As String, networkname As String)
        Dim command As String = "interface ip delete dns name=""" & networkname & """ address=""" & dnsaddress & """ "
        Dim pr As New Process
        pr.StartInfo.FileName = "netsh.exe"
        pr.StartInfo.UseShellExecute = False
        pr.StartInfo.RedirectStandardInput = True
        pr.StartInfo.RedirectStandardOutput = True
        pr.StartInfo.CreateNoWindow = True
        pr.Start()
        Dim wr As System.IO.StreamWriter = pr.StandardInput
        Dim rr As System.IO.StreamReader = pr.StandardOutput

        wr.WriteLine(command)
        wr.Flush()
        wr.Close()
        pr.WaitForExit()
        pr.Close()

    End Sub

#End Region


    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        _timer.Stop()
        _timer = Nothing
        _timerInternetConnection.Stop()
        _timerInternetConnection = Nothing
        fNetworkInterfaces = Nothing

    End Sub
End Class
