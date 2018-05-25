Imports System.Net.Sockets
Imports System.Text


Module Main


    Dim _clientdatareceived As String
    Dim _portnumber1 As Integer = 8000
    Dim _portnumber2 As Integer = 8001
    Dim _portnumber3 As Integer = 8002
    Dim _portnumber4 As Integer = 8003
    Sub Main()
        'Get the port from the app.config
        Try
            _portnumber1 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-main"))
            _portnumber2 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-2"))
            _portnumber3 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-3"))
            _portnumber4 = Integer.Parse(System.Configuration.ConfigurationSettings.AppSettings("Port-4"))
        Catch ex As Exception

        End Try

        'Run Server
        Dim _th_runserver As New Threading.Thread(AddressOf runServer)
        _th_runserver.Start()

        Dim _th_runserver2 As New Threading.Thread(AddressOf runServer)
        _th_runserver2.Start()

        Dim _th_runserver3 As New Threading.Thread(AddressOf runServer)
        _th_runserver3.Start()

        Dim _th_runserver4 As New Threading.Thread(AddressOf runServer)
        _th_runserver4.Start()

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
            Dim tcpListener As New TcpListener(portNumber)
            tcpListener.Start()
            Console.WriteLine("Waiting for connection on port " & portNumber.ToString & "...")
            Try
                'Accept the pending client connection and return 
                'a TcpClient initialized for communication. 
                Dim tcpClient As TcpClient = tcpListener.AcceptTcpClient()
                tcpClient.ReceiveBufferSize = 1024
                Console.WriteLine("Connection accepted.")

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
                    Console.WriteLine("Closing Connection...")
                    closeconnection = False
                    'Application.Exit()
                    'Exit Sub
                End If

                'Process if there is a command in the line
                Dim th_processcommand As New Threading.Thread(AddressOf ProcessCMD)
                th_processcommand.Start(mensaje)

                If mensaje.Contains("STA@") Then
                    Dim responseString As String = "ACK"
                    Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    Console.WriteLine("Message Sent /> : " & responseString)
                End If

                If mensaje.Contains("MON@") Then
                    Dim responseString As String = "ACKCHG"
                    Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    Console.WriteLine("Message Sent /> : " & responseString)
                End If

                If mensaje.Contains("ONL@") Then
                    Dim responseString As String = "ACKCHG"
                    Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    Console.WriteLine("Message Sent /> : " & responseString)
                End If

                If mensaje.Contains("MNT@") Then
                    Dim responseString As String = "ACKCHG"
                    Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    Console.WriteLine("Message Sent /> : " & responseString)
                End If

                If mensaje.Contains("INT@") Then
                    Dim responseString As String = "ACKCHG"
                    Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(responseString)
                    networkStream.Write(sendBytes, 0, sendBytes.Length)
                    Console.WriteLine("Message Sent /> : " & responseString)
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
            End Select
        End If
    End Sub

End Module
