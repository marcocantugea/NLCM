Imports System.Net.Sockets
Imports System.Text
Imports System.Collections

Module Main


    Dim _clientdatareceived As String
    Dim _portnumber1 As Integer = 8000
    Dim _portnumber2 As Integer = 8001
    Dim _portnumber3 As Integer = 8002
    Dim _portnumber4 As Integer = 8003

    'added 28 may 2018
    ' Variable IP Table for optimize performance
    ' Also variable to update the table every 30 seconds
    Dim _IPTable As New Dictionary(Of String, IPInfo)
    Dim _shutdownapp As Boolean = False

    Dim _th_runserver As New Threading.Thread(AddressOf runServer)
    Dim _th_runserver2 As New Threading.Thread(AddressOf runServer)
    Dim _th_runserver3 As New Threading.Thread(AddressOf runServer)
    Dim _th_runserver4 As New Threading.Thread(AddressOf runServer)
    Dim _th_UpdateIPTable As New Threading.Thread(AddressOf UpdateIPTable)
    Dim _th_shutdownapp As New Threading.Thread(AddressOf Shutdownapp)
    Sub Main()

        'Added 28 May 2018
        'Get The IP Table to memory
        Dim ado As New ADO
        ado.GetIPTable(_IPTable)


        'Get the port from the app.config
        Try
            _portnumber1 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-main"))
            _portnumber2 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-2"))
            _portnumber3 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-3"))
            _portnumber4 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-4"))
        Catch ex As Exception

        End Try

        'Run Server
        _th_runserver.Start()

        _th_runserver2.Start()

        _th_runserver3.Start()

        _th_runserver4.Start()

        _th_UpdateIPTable.Start()

        _th_shutdownapp.Start()
        
    End Sub

    Public Sub runServer()
        Try
            'Open port 1
            OpenConnectionService(_portnumber1)
        Catch ex As Exception
            Try
                'Open port 2
                OpenConnectionService(_portnumber2)
            Catch exs1 As Exception
                Try
                    'Open port 3
                    OpenConnectionService(_portnumber3)
                Catch exs2 As Exception
                    Try
                        'Open port 4
                        OpenConnectionService(_portnumber4)
                    Catch exs3 As Exception

                    End Try
                End Try
            End Try

        End Try
    End Sub

    Public Sub OpenConnectionService(portNumber As Integer)

        Dim closeconnection As Boolean = True
        While closeconnection

            ''Added 27 May 2018
            ''Adding the Maintenance Mode
            Dim maintenancemode As Boolean = False
            Try
                maintenancemode = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("MaintenanceMode"))
            Catch ex As Exception

            End Try

            ''Added 28 May 2018
            ''Adding shutdown mode for clients
            Dim shutdownmode As Boolean = False
            Try
                shutdownmode = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("ShutdownMode"))
            Catch ex As Exception

            End Try

            Dim tcpListener As New TcpListener(portNumber)
            tcpListener.Start()
            Console.WriteLine("Waiting for connection on port " & portNumber.ToString & "...")
            Try
                'Accept the pending client connection and return 
                'a TcpClient initialized for communication. 
                Dim tcpClient As TcpClient = tcpListener.AcceptTcpClient()
                Dim ipclient As System.Net.IPEndPoint = tcpClient.Client.RemoteEndPoint
                tcpClient.ReceiveBufferSize = 1024
                Console.WriteLine("Connection accepted from " & ipclient.Address.ToString & " Port: " & portNumber)

                ' Get the stream
                Dim networkStream As NetworkStream = tcpClient.GetStream()
                ' Read the stream into a byte array
                Dim bytes(tcpClient.ReceiveBufferSize) As Byte
                Dim myreadbuffer(1024) As Byte
                Dim numbersofbytesRead As Integer = 0
                numbersofbytesRead = networkStream.Read(myreadbuffer, 0, CInt(tcpClient.ReceiveBufferSize))

                ' Return the data received from the client to the console.
                Dim messagescut(numbersofbytesRead - 1) As Byte
                Dim index As Integer = 0
                For Each n As Object In myreadbuffer
                    If Not Integer.Parse(n).Equals(0) Then
                        messagescut.SetValue(n, index)
                        index += 1
                    End If
                Next
                '_clientdatareceived = Encoding.ASCII.GetString(myreadbuffer)
                Console.WriteLine("numbers of bytes read = " & numbersofbytesRead.ToString)
                Console.WriteLine("Data received: " & Encoding.ASCII.GetString(messagescut))

                ' formated message
                Dim mensaje As String = Encoding.ASCII.GetString(messagescut)

                'if the message is exitrun will close the program
                If mensaje.Equals("exitrun") Then
                    Console.WriteLine("Closing Port " & portNumber & "...")
                    closeconnection = False
                    _shutdownapp = True
                End If

                If _shutdownapp Then
                    closeconnection = False
                End If

                If maintenancemode Then
                    Dim responseString As String = "MAINTENANCEON!"
                    Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    Console.WriteLine("Message Sent /> : " & responseString)
                Else

                    If shutdownmode Then
                        Dim responseString As String = "SHUTDOWN!"
                        Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                        networkStream.Write(sendBytes, 0, sendBytes.Length)
                        Console.WriteLine("Message Sent /> : " & responseString)
                    Else
                        'Process if there is a command in the line
                        Dim th_processcommand As New Threading.Thread(AddressOf ProcessCMD)
                        th_processcommand.Start(mensaje)

                        ' Command STA
                        If mensaje.Contains("STA@") Then
                            Dim responseString As String = "ACK"
                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & responseString)
                        End If

                        ' Command MON
                        If mensaje.Contains("MON@") Then
                            Dim responseString As String = "ACKCHG"
                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & responseString)
                        End If

                        ' Command ONL
                        If mensaje.Contains("ONL@") Then
                            Dim responseString As String = "ACKCHG"
                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & responseString)
                        End If

                        ' Command MNT
                        If mensaje.Contains("MNT@") Then
                            Dim responseString As String = "ACKCHG"
                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & responseString)
                        End If

                        ' Command INT
                        If mensaje.Contains("INT@") Then
                            Dim responseString As String = "ACKCHG"
                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & responseString)
                        End If

                        ' Command INT
                        If mensaje.Contains("DEVADD@") Then
                            Dim responseString As String = "ACKCHG"
                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & responseString)
                        End If

                        ' Command ONLINE?
                        If mensaje.Contains("ONLINE?") Then
                            Dim cmd As String() = mensaje.Split("@")
                            Dim returnmessage As String
                            'validate the structure of the command
                            If cmd.Length = 2 Then
                                Dim online As Boolean = True
                                Try
                                    If _IPTable.Count > 0 Then
                                        Dim ipinfo As IPInfo = _IPTable(cmd(1))
                                        online = ipinfo.ONLINE
                                    Else
                                        Dim ADO As New ADO
                                        online = ADO.GetOnlineParameter(cmd(1))
                                    End If
                                    If online Then
                                        returnmessage = "ACK@1"
                                    Else
                                        returnmessage = "ACK@0"
                                    End If
                                Catch ex As Exception
                                End Try
                            End If

                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(returnmessage)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & returnmessage)

                        End If
                        ' Command INTERVAL?
                        If mensaje.Contains("INTERVAL?") Then
                            Dim cmd As String() = mensaje.Split("@")
                            Dim returnmessage As String
                            'validate the structure of the command
                            If cmd.Length = 2 Then
                                Dim interval As Integer = 1
                                Try
                                    If _IPTable.Count > 0 Then
                                        Dim ipinfo As IPInfo = _IPTable(cmd(1))
                                        interval = ipinfo.Interval_REC
                                    Else
                                        Dim ADO As New ADO
                                        interval = ADO.GetIntervalRec(cmd(1))

                                    End If
                                Catch ex As Exception
                                    'TODO Log of exception
                                End Try

                                returnmessage = "ACK@" & interval.ToString
                            End If

                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(returnmessage)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & returnmessage)

                        End If

                        ' Command MONITOR?
                        If mensaje.Contains("MONITOR?") Then
                            Dim cmd As String() = mensaje.Split("@")
                            Dim returnmessage As String
                            'validate the structure of the command
                            If cmd.Length = 2 Then

                                Dim monitor As Boolean = True
                                Try
                                    ''Added 28 may 2018
                                    '' added IP Table on memory to fast response
                                    If _IPTable.Count > 0 Then
                                        If _IPTable.ContainsKey(cmd(1)) Then
                                            Dim ipinfo As IPInfo = _IPTable(cmd(1))
                                            monitor = ipinfo.MONITOR
                                            If monitor Then
                                                returnmessage = "ACK@1"
                                            Else
                                                returnmessage = "ACK@0"
                                            End If
                                        Else
                                            Dim ADO As New ADO
                                            monitor = ADO.CheckmMonitorIsEnable(cmd(1))
                                            If monitor Then
                                                returnmessage = "ACK@1"
                                            Else
                                                returnmessage = "ACK@0"
                                            End If
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            End If

                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(returnmessage)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & returnmessage)

                        End If

                        ' Command MONITOR?
                        If mensaje.Contains("IP?") Then
                            Dim cmd As String() = mensaje.Split("@")
                            Dim returnmessage As String
                            'validate the structure of the command
                            If cmd.Length = 2 Then

                                Dim ip_exist As Boolean = True
                                Try
                                    ''Added 28 may 2018
                                    '' added IP Table on memory to fast response
                                    If _IPTable.Count > 0 Then
                                        If _IPTable.ContainsKey(cmd(1)) Then
                                            ip_exist = True
                                        Else
                                            ip_exist = False
                                        End If
                                    Else
                                        Dim ADO As New ADO
                                        ip_exist = ADO.CheckIPExist(cmd(1))
                                    End If

                                    If ip_exist Then
                                        returnmessage = "ACK@1"
                                    Else
                                        returnmessage = "ACK@0"
                                    End If
                                Catch ex As Exception

                                End Try
                            End If

                            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(returnmessage)
                            networkStream.Write(sendBytes, 0, sendBytes.Length)
                            Console.WriteLine("Message Sent /> : " & returnmessage)

                        End If
                    End If
                End If




                'Any communication with the remote client using the TcpClient can go here.
                'Close TcpListener and TcpClient.
                tcpClient.Close()
                tcpListener.Stop()
                tcpListener = Nothing

            Catch ex As Exception
                Console.WriteLine(ex.Message.ToString)
                Throw
            End Try

        End While

    End Sub

    Private Sub ProcessCMD(args As String)
        If args.Contains("@") Then
            Dim cmd As String() = args.Split("@")
            Select Case cmd(0)
                Case "MON"
                    'validate the structure of the command
                    If cmd.Length = 2 Then
                        Dim value_args As String = cmd(1)
                        If value_args.Contains("!") Then
                            'valid command
                            Dim values As String() = value_args.Split("!")
                            'checks the legnth of the args and all the values are set
                            If values.Length = 5 Then
                                ' will create a IPStatisticObj
                                Dim new_IPStatistics As New IPStatisticsObj
                                new_IPStatistics.IP_ADDRESS = values(0)
                                'get the date from the value(1) that cames like YYYYMMDDHHmmSS
                                Dim year As Integer = values(1).Substring(0, 4)
                                Dim month As Integer = values(1).Substring(4, 2)
                                Dim day As Integer = values(1).Substring(6, 2)
                                Dim hrs As Integer = values(1).Substring(8, 2)
                                Dim minutes As Integer = values(1).Substring(10, 2)
                                Dim seconds As Integer = values(1).Substring(12, 2)

                                Dim date_record As New DateTime(year, month, day, hrs, minutes, seconds)

                                new_IPStatistics.DATE_RECORD = date_record

                                new_IPStatistics.MBCONSUMPTION_IN = values(2)
                                new_IPStatistics.MBCONSUMPTION_OUT = values(3)
                                new_IPStatistics.SESSIONUSERLOGED = values(4)

                                Try
                                    Dim ADO As New ADO
                                    ADO.RecordIP(new_IPStatistics)
                                    Console.WriteLine("Record Added info-> IP:" & new_IPStatistics.IP_ADDRESS & "/DATE:" & new_IPStatistics.DATE_RECORD & "/MBIN:" & new_IPStatistics.MBCONSUMPTION_IN.ToString & "/MBOUT:" & new_IPStatistics.MBCONSUMPTION_OUT.ToString & "/USER:" & new_IPStatistics.SESSIONUSERLOGED)
                                Catch ex As Exception
                                    Console.WriteLine("!!!Fail to save the info-> IP:" & new_IPStatistics.IP_ADDRESS & "/DATE:" & new_IPStatistics.DATE_RECORD & "/MBIN:" & new_IPStatistics.MBCONSUMPTION_IN.ToString & "/MBOUT:" & new_IPStatistics.MBCONSUMPTION_OUT.ToString & "/USER:" & new_IPStatistics.SESSIONUSERLOGED)
                                End Try
                            Else
                                Console.WriteLine("!!!Monitor command set but Arguments are not valid ")
                            End If
                            'Console.WriteLine("Monitor command set. to register in the DB. " & values)
                        End If
                    End If
                Case "ONL"
                    'validate the structure of the command
                    If cmd.Length = 2 Then
                        Dim value_args As String = cmd(1)
                        'valid command
                        Dim values As String() = value_args.Split("!")
                        'checks the legnth of the args and all the values are set
                        If values.Length = 2 Then

                            Dim upd_IPInfo As New IPInfo
                            upd_IPInfo.IP_ADDRESS = values(0)
                            If values(1).Equals("ON") Then
                                upd_IPInfo.ONLINE = True
                            End If
                            If values(1).Equals("OFF") Then
                                upd_IPInfo.ONLINE = False
                            End If
                            Try
                                Dim ADO As New ADO
                                ADO.ChangeOnlineIP(upd_IPInfo)
                                Console.WriteLine("Record Changed info-> IP:" & upd_IPInfo.IP_ADDRESS & "/ONLIE:" & upd_IPInfo.ONLINE.ToString & "")
                            Catch ex As Exception
                                Console.WriteLine("!!!Fail to change Record info-> IP:" & upd_IPInfo.IP_ADDRESS & "/ONLIE:" & upd_IPInfo.ONLINE.ToString & "")
                            End Try
                        End If
                    End If
                Case "MNT"
                    'validate the structure of the command
                    If cmd.Length = 2 Then
                        Dim value_args As String = cmd(1)
                        'valid command
                        Dim values As String() = value_args.Split("!")
                        'checks the legnth of the args and all the values are set
                        If values.Length = 2 Then

                            Dim upd_IPInfo As New IPInfo
                            upd_IPInfo.IP_ADDRESS = values(0)
                            If values(1).Equals("ON") Then
                                upd_IPInfo.MONITOR = True
                            End If
                            If values(1).Equals("OFF") Then
                                upd_IPInfo.MONITOR = False
                            End If
                            Try
                                Dim ADO As New ADO
                                ADO.ChangeMonitorIP(upd_IPInfo)
                                Console.WriteLine("Record Changed info-> IP:" & upd_IPInfo.IP_ADDRESS & "/MONITOR:" & upd_IPInfo.MONITOR.ToString & "")
                            Catch ex As Exception
                                Console.WriteLine("!!!Fail to change Record info-> IP:" & upd_IPInfo.IP_ADDRESS & "/MONITOR:" & upd_IPInfo.MONITOR.ToString & "")
                            End Try
                        End If
                    End If
                Case "INT"
                    'validate the structure of the command
                    If cmd.Length = 2 Then
                        Dim value_args As String = cmd(1)
                        'valid command
                        Dim values As String() = value_args.Split("!")
                        'checks the legnth of the args and all the values are set
                        If values.Length = 2 Then
                            Try
                                Dim upd_IPInfo As New IPInfo
                                upd_IPInfo.IP_ADDRESS = values(0)
                                upd_IPInfo.Interval_REC = Integer.Parse(values(1))
                                Try
                                    Dim ADO As New ADO
                                    ADO.ChangeIntervalIP(upd_IPInfo)
                                    Console.WriteLine("Record Changed info-> IP:" & upd_IPInfo.IP_ADDRESS & "/INTERVAL:" & upd_IPInfo.Interval_REC.ToString & "")
                                Catch ex As Exception
                                    Console.WriteLine("!!!Fail to change Record info-> IP:" & upd_IPInfo.IP_ADDRESS & "/INTERVAL:" & upd_IPInfo.Interval_REC.ToString & "")
                                End Try
                            Catch ex As Exception
                                Console.WriteLine("!!!Interval command set but Arguments are not valid ")
                            End Try

                        End If
                    End If
                Case "DEVADD"
                    'validate the structure of the command
                    If cmd.Length = 2 Then
                        Dim value_args As String = cmd(1)
                        'valid arguments
                        Dim values As String() = value_args.Split("!")
                        If values.Length = 8 Then
                            Dim new_ipinfo As New IPInfo
                            new_ipinfo.IP_ADDRESS = values(0)
                            new_ipinfo.LOCATION = values(1)
                            new_ipinfo.MONITOR = values(2)
                            new_ipinfo.ONLINE = values(3)
                            new_ipinfo.ACTIVE = values(4)
                            new_ipinfo.Interval_REC = values(5)
                            new_ipinfo.ADAPTERNAME = values(6)
                            new_ipinfo.MACADDRESS = values(7)
                            Try
                                Dim ADO As New ADO
                                ADO.RecordIPInfo(new_ipinfo)
                            Catch ex As Exception
                                Console.WriteLine("!!!Fail to add Record info-> DEVADD Command")
                            End Try

                        End If
                    End If
            End Select
        End If
    End Sub


    Private Sub UpdateIPTable()
        Dim tick As Integer = 0
        Dim tick_triger As Integer = 2
        While tick <= tick_triger
            If tick = tick_triger Then
                _IPTable.Clear()
                'Added 28 May 2018
                'Get The IP Table to memory
                Dim ado As New ADO
                ado.GetIPTable(_IPTable)
                Console.WriteLine("IP Table Update it...")
                tick = 0
            Else
                Threading.Thread.Sleep(30000)
                tick += 1
            End If
            If _shutdownapp Then
                tick = 4
            End If
        End While

        
    End Sub

    Private Sub Shutdownapp()

        Dim closeapp As Boolean = False
        Dim endloop As Integer = False

        While Not endloop
            If _shutdownapp Then
                If Not _th_runserver.IsAlive Then
                    If Not _th_runserver2.IsAlive Then
                        If Not _th_runserver3.IsAlive Then
                            If Not _th_runserver4.IsAlive Then
                                If Not _th_UpdateIPTable.IsAlive Then
                                    closeapp = True
                                    endloop = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End While

        If closeapp Then
            Console.WriteLine("Shutting Down The Application...")
            Application.Exit()
            End
        End If

        '_th_runserver3.Abort()
        '_th_runserver4.Abort()
        '_th_UpdateIPTable.Abort()



    End Sub

End Module

Public Class CMDParameters
    Private _args As String
    Private _returnmessage As String

    Public Property returnmessage As String
        Get
            Return _returnmessage
        End Get
        Set(value As String)
            _returnmessage = value
        End Set
    End Property

    Public Property args As String
        Get
            Return _args
        End Get
        Set(value As String)
            _args = value
        End Set
    End Property

End Class
