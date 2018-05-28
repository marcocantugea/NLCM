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
    Dim _IPServer As String = "127.0.0.1"

    ''added 5/24/2018 ports and server
    Dim _server As String = "127.0.0.1"
    Dim _port1 As Integer = 8000
    Dim _port2 As Integer = 8001
    Dim _port3 As Integer = 8002
    Dim _port4 As Integer = 8003

    'Added 26 may 2018
    ' Fix performance to check the interval on the Server
    Dim _intervaltoRecord As Long

    'Added 27 May 2018
    Private WithEvents _timer_maintenance As New Timer

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

    'Startup function
    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)
        If Not Me.IsHandleCreated Then
            Me.CreateHandle()
            value = False
        End If

        'Load server parameters and ports
        Me._server = System.Configuration.ConfigurationSettings.AppSettings("server")
        Try
            Me._port1 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("port-main"))
            Me._port2 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("port-1"))
            Me._port3 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("port-2"))
            Me._port4 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("port-3"))
        Catch ex As Exception

        End Try

        'set timer to minutes
        _timer.Interval = 1000
        _timerInternetConnection.Interval = 180000
        'initialize timers
        _timer.Start()
        _timerInternetConnection.Start()

        'initilizar monitor of network
        StartMonitor()

        'initiliza timer maintenance
        _timer_maintenance.Interval = 30000

        Dim returnmsg As String = ""
        Dim cmd As String = "STA@" & getIPAddr(), _port1, _server

        Try
            'Send to server alive wake up
            returnmsg = SendCommand(cmd)

            'if is maintenance mode set maintenance mode on the app
            If returnmsg.Contains("MAINTENANCEON!") Then
                SetMaintenanceMode()
            End If

            'Added 28 may 2018
            'added shutdown mode
            If returnmsg.Contains("SHUTDOWN!") Then
                SetShutDownMode()
            End If

        Catch ex As Exception

        End Try

        'Addded 26 may 2018
        ' set the interval to start
        'get the interval to save the record and convert into minutes
        _intervaltoRecord = GetIntervalRecord() * 60

        MyBase.SetVisibleCore(value)

    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
       

    End Sub


    'main time ticker
    Private Sub TimerTick(sender As Object, e As EventArgs) Handles _timer.Tick
        'get the interval to save the record and convert into minutes

        'checks if the timer meets the interval configured in the app.conf
        If _intervalTimerTick = _intervaltoRecord Then
            ''Label1.Text = "entro"
            'Set tie interval to 0
            _intervalTimerTick = 0
            'Check is the Monitoring is enable
            CheckMonitoringEngine()
            ' if is enable will record the data
            If _EnableMonitor Then
                CloseEverything(Nothing, Nothing)
            End If
            'get the interval 
            _intervaltoRecord = GetIntervalRecord() * 60

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
            ''modified 5 may 2018
            '' check on server the paramerter online
            'online = _ADO.GetOnlineParameter(getIPAddr)

            Dim cmd As String = "ONLINE?@" & getIPAddr()
            Dim returnmsg As String

            Try
                returnmsg = SendCommand(cmd)
                'if is maintenance mode set maintenance mode on the app
                If returnmsg.Contains("MAINTENANCEON!") Then
                    SetMaintenanceMode()
                End If

                'Added 28 may 2018
                'added shutdown mode
                If returnmsg.Contains("SHUTDOWN!") Then
                    SetShutDownMode()
                End If

                If Not returnmsg.Equals("") Then
                    If returnmsg.Contains("@") Then
                        Dim args() As String = returnmsg.Split("@")
                        If args.Length = 2 Then
                            Dim response As Integer = Integer.Parse(args(1))
                            online = response
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try


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

        'Modified 26 may 2018
        ' Remove consulting DB
        'get the configuration from the database
        'Dim _ADO_IntervalRec As Integer
        'Try
        '    _ADO_IntervalRec = _ADO.GetIntervalRec(getIPAddr)
        'Catch ex As Exception
        '    ''TODO: Log
        '    ''TODO : Remove hard corde
        '    _ADO_IntervalRec = 1
        'End Try

        'Added function to ask to the server the parameters
        Dim cmd As String = "INTERVAL?@" & getIPAddr()
        Dim returnmsg As String = ""


        Try
            returnmsg = SendCommand(cmd)

            'if is maintenance mode set maintenance mode on the app
            If returnmsg.Contains("MAINTENANCEON!") Then
                SetMaintenanceMode()
            End If

            'Added 28 may 2018
            'added shutdown mode
            If returnmsg.Contains("SHUTDOWN!") Then
                SetShutDownMode()
            End If

            If Not returnmsg.Equals("") Then
                If returnmsg.Contains("@") Then
                    Dim args() As String = returnmsg.Split("@")
                    If args.Length = 2 Then
                        Dim response As Integer = 0
                        Try
                            response = Integer.Parse(args(1))
                            IntervalRecord = response
                        Catch ex As Exception
                            Dim Setting_intervalrecord As String = System.Configuration.ConfigurationSettings.AppSettings("Interval-Rec")
                            If Not IsNothing(Setting_intervalrecord) Then
                                Try
                                    IntervalRecord = Long.Parse(Setting_intervalrecord)
                                Catch ex1 As Exception
                                    IntervalRecord = 5
                                End Try
                            End If
                        End Try
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

        'if the database configuration is set will set the return variable
        'If _ADO_IntervalRec > 0 Then
        '    IntervalRecord = _ADO_IntervalRec
        'Else
        'if the variable is not set will take the app.config to setup the interval record
        ' the configuration variable is Interval-Rec 

        'End If


        Return IntervalRecord
    End Function

    'function to check the parameter MONITOR in the database to record the consumption
    Private Sub CheckMonitoringEngine()
        'Modified remove the database check the server to get the parameter
        'get the ip adress for check the parameter monitor on the database

        'Dim checkmonitor As Boolean = _ADO.CheckmMonitorIsEnable(getIPAddr)
        'If Not checkmonitor Then
        '    _EnableMonitor = False
        'Else
        '    _EnableMonitor = True
        'End If

        Dim cmd As String = "MONITOR?@" & getIPAddr()
        Dim returnmsg As String = ""
        Try
            returnmsg = SendCommand(cmd)

            'if is maintenance mode set maintenance mode on the app
            If returnmsg.Contains("MAINTENANCEON!") Then
                SetMaintenanceMode()
            End If

            'Added 28 may 2018
            'added shutdown mode
            If returnmsg.Contains("SHUTDOWN!") Then
                SetShutDownMode()
            End If

            If Not returnmsg.Equals("") Then
                If returnmsg.Contains("@") Then
                    Dim args() As String = returnmsg.Split("@")
                    If args.Length = 2 Then
                        Dim enablemonitor As Boolean
                        enablemonitor = Integer.Parse(args(1))
                        _EnableMonitor = enablemonitor
                    End If
                End If
            End If
        Catch ex As Exception

        End Try


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

            'Old code replaced to connect to server and save the info
            'save record in database
            '_ADO.RecordIP(_PCInfo)

            'new code added 5/24/2018
            Dim dateforcmd As String = _PCInfo.DATE_RECORD.ToString("yyyyMMddhhmmss")
            Dim cmd_to_server As String = "MON@" & _PCInfo.IP_ADDRESS & "!" & dateforcmd & "!" & _PCInfo.MBCONSUMPTION_IN & "!" & _PCInfo.MBCONSUMPTION_OUT & "!" & _PCInfo.SESSIONUSERLOGED
            Dim retry As Integer = 0
            While retry <= 3
                Dim returnmsg As String = ""
                Try
                    returnmsg = SendCommand(cmd_to_server)
                    'if is maintenance mode set maintenance mode on the app
                    If returnmsg.Contains("MAINTENANCEON!") Then
                        SetMaintenanceMode()
                    End If

                    'Added 28 may 2018
                    'added shutdown mode
                    If returnmsg.Contains("SHUTDOWN!") Then
                        SetShutDownMode()
                    End If

                    retry = 4
                Catch ex As Exception
                    retry += 1
                End Try


                'Try
                '    SendCommandToServer(cmd_to_server, _port1, _server)
                '    retry = 4
                'Catch ex As Exception
                '    Try
                '        SendCommandToServer(cmd_to_server, _port2, _server)
                '        retry = 4
                '    Catch ex1 As Exception
                '        Try
                '            SendCommandToServer(cmd_to_server, _port3, _server)
                '            retry = 4
                '        Catch ex2 As Exception
                '            Try
                '                SendCommandToServer(cmd_to_server, _port4, _server)
                '                retry = 4
                '            Catch ex3 As Exception
                '                retry += 1
                '            End Try
                '        End Try
                '    End Try
                'End Try

            End While

            ''Modified 26 may 2018
            '' remove to add directly to the data base
            '' send command to the server
            'check if the ip address if is registered on the table
            ''Dim chek_ip As Boolean = _ADO.CheckIPExist(_PCInfo.IP_ADDRESS)
            Dim chek_ip As Boolean = False
            Dim cmd_checkip As String = "IP?@" & _PCInfo.IP_ADDRESS
            Dim return_msg As String = ""
            Try
                return_msg = SendCommand(cmd_checkip)
                'if is maintenance mode set maintenance mode on the app
                If return_msg.Contains("MAINTENANCEON!") Then
                    SetMaintenanceMode()
                End If

                'Added 28 may 2018
                'added shutdown mode
                If return_msg.Contains("SHUTDOWN!") Then
                    SetShutDownMode()
                End If

                If Not return_msg.Equals("") Then
                    If return_msg.Contains("@") Then
                        Dim args() As String = return_msg.Split("@")
                        If args.Length = 2 Then
                            chek_ip = Integer.Parse(args(1))
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try




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
                new_ipinfo.MACADDRESS = getMacAddress()

                ''Modified 26 may 2018
                '' remove to add directly to the data base
                '' send command to the server
                '_ADO.RecordIPInfo(new_ipinfo)
                Dim cmd As String
                Dim monitor As String = "0"
                If new_ipinfo.MONITOR Then
                    monitor = "1"

                End If

                Dim onlie As String = "0"
                If new_ipinfo.ONLINE Then
                    onlie = "1"
                End If

                Dim active As String = "0"
                If new_ipinfo.ACTIVE Then
                    active = "1"
                End If

                cmd = "DEVADD@" & new_ipinfo.IP_ADDRESS & "!" & new_ipinfo.LOCATION & "!" & monitor & "!" & onlie & "!" & active & "!" & new_ipinfo.Interval_REC.ToString & "!" & new_ipinfo.ADAPTERNAME & "!" & new_ipinfo.MACADDRESS
                Dim returnmsg As String = ""
                Try
                    returnmsg = SendCommand(cmd)
                    'if is maintenance mode set maintenance mode on the app
                    If return_msg.Contains("MAINTENANCEON!") Then
                        SetMaintenanceMode()
                    End If

                    'Added 28 may 2018
                    'added shutdown mode
                    If return_msg.Contains("SHUTDOWN!") Then
                        SetShutDownMode()
                    End If

                Catch ex As Exception

                End Try


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

    ''Added 22 May 2018
    ''Adding functionality to run as server client
    Public Function SendCommandToServer(cmd As String, port As Integer, server As String) As String
        Dim _returndata As String
        Dim tcpClient As New System.Net.Sockets.TcpClient()
        tcpClient.Connect(server, port)
        tcpClient.ReceiveBufferSize = 1024
        Dim networkStream As NetworkStream = tcpClient.GetStream()
        If networkStream.CanWrite And networkStream.CanRead Then

            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(cmd)
            networkStream.Write(sendBytes, 0, sendBytes.Length)


            Dim bytes(tcpClient.ReceiveBufferSize) As Byte
            networkStream.Read(bytes, 0, CInt(tcpClient.ReceiveBufferSize))

            _returndata = Encoding.ASCII.GetString(bytes)
            'Label1.Text = "Received  > " & _returndata & Environment.NewLine
            'If _returndata.Equals("ACK") Then
            tcpClient.Close()
            'End If
        Else
            tcpClient.Close()
        End If
        Return _returndata
    End Function

    ''Added 27 May 2018
    '' Send command in all ports
    Public Function SendCommand(cmd As String) As String
        Dim return_msg As String = ""
        Try
            return_msg = SendCommandToServer(cmd, _port1, _server)
        Catch ex As Exception
            Try
                return_msg = SendCommandToServer(cmd, _port2, _server)
            Catch ex1 As Exception
                Try
                    return_msg = SendCommandToServer(cmd, _port3, _server)
                Catch ex2 As Exception
                    Try
                        return_msg = SendCommandToServer(cmd, _port4, _server)
                    Catch ex3 As Exception
                        'TODO:buffer to saved for latter
                        Throw
                    End Try
                End Try
            End Try
        End Try
        Return return_msg
    End Function

    Function getMacAddress()
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
        Return nics(0).GetPhysicalAddress.ToString
    End Function


    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        _timer.Stop()
        _timer = Nothing
        _timerInternetConnection.Stop()
        _timerInternetConnection = Nothing
        fNetworkInterfaces = Nothing

        
    End Sub

    'Added 27 May 2018
    Private Sub SetMaintenanceMode()
        _timer.Stop()
        _timerInternetConnection.Stop()
        _timer_maintenance.Start()
    End Sub

    Private Sub Timer_MaintenanceMode(sender As Object, e As EventArgs) Handles _timer_maintenance.Tick

        Dim cmd As String = "STA@" & getIPAddr()
        Dim returnmsg As String = ""
        Try
            returnmsg = SendCommand(cmd)
            If returnmsg.Contains("ACK") Then
                _timer_maintenance.Stop()
                _timer.Start()
                _timerInternetConnection.Start()
            End If
        Catch ex As Exception

        End Try


    End Sub

    'Added 28 May 2018
    ' Function to shutdown mode
    Private Sub SetShutDownMode()
        Application.Exit()
        End

    End Sub


End Class
