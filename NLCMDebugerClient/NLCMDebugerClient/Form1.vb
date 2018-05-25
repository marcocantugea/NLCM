Imports System.Net.Sockets
Imports System.Text

Public Class Form1

    Dim _returndata As String
    Dim WithEvents _timer As New Timer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _timer.Interval = 100
        TextBox2.Text = 100
    End Sub

    Private Sub Timer_tick(sender As Object, e As EventArgs) Handles _timer.Tick
        SendCommandToServer("MON@192.168.45.6!465646464!50!5!LM4PCVIT01")
    End Sub

    Public Sub SendCommandToServer(cmd As String)
        Dim tcpClient As New System.Net.Sockets.TcpClient()
        tcpClient.Connect("127.0.0.1", 8000)
        tcpClient.ReceiveBufferSize = 1024
        Dim networkStream As NetworkStream = tcpClient.GetStream()
        If networkStream.CanWrite And networkStream.CanRead Then

            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(cmd)
            networkStream.Write(sendBytes, 0, sendBytes.Length)


            Dim bytes(tcpClient.ReceiveBufferSize) As Byte
            networkStream.Read(bytes, 0, CInt(tcpClient.ReceiveBufferSize))

            _returndata = Encoding.ASCII.GetString(bytes)
            Label1.Text = "Received  > " & _returndata & Environment.NewLine
            'If _returndata.Equals("ACK") Then
            tcpClient.Close()
            'End If
        Else
            tcpClient.Close()
        End If
    End Sub

    Public Sub SendCommandToServer(cmd As String, port As Integer)
        Dim tcpClient As New System.Net.Sockets.TcpClient()
        tcpClient.Connect("127.0.0.1", port)
        tcpClient.ReceiveBufferSize = 1024
        Dim networkStream As NetworkStream = tcpClient.GetStream()
        If networkStream.CanWrite And networkStream.CanRead Then

            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(cmd)
            networkStream.Write(sendBytes, 0, sendBytes.Length)


            Dim bytes(tcpClient.ReceiveBufferSize) As Byte
            networkStream.Read(bytes, 0, CInt(tcpClient.ReceiveBufferSize))

            _returndata = Encoding.ASCII.GetString(bytes)
            Label1.Text = "Received  > " & _returndata & Environment.NewLine
            'If _returndata.Equals("ACK") Then
            tcpClient.Close()
            'End If
        Else
            tcpClient.Close()
        End If
    End Sub

    Public Sub SendCommandToServer(cmd As String, port As Integer, server As String)
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
            Label1.Text = "Received  > " & _returndata & Environment.NewLine
            'If _returndata.Equals("ACK") Then
            tcpClient.Close()
            'End If
        Else
            tcpClient.Close()
        End If
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim port As Integer = Integer.Parse(TextBox3.Text)
        Dim server As String = TextBox4.Text
        If Not TextBox1.Text.Equals("") Then
            Dim cmd As String = TextBox1.Text
            Try
                SendCommandToServer(cmd, port, server)
            Catch ex As Exception
                Label1.Text = "Can not connect to server."
            End Try

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        _timer.Interval = Integer.Parse(TextBox2.Text)
        _timer.Start()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        _timer.Stop()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        
    End Sub
End Class
