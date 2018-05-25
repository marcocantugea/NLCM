Imports System.IO
Public Class Form1


    Dim _folderserver As String
    Dim _localfolder As String
    Private WithEvents P As Process

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _folderserver = System.Configuration.ConfigurationSettings.AppSettings("folderserver")
        _localfolder = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & System.Configuration.ConfigurationSettings.AppSettings("localfolder")
        Start()
        Application.Exit()

    End Sub

    Public Sub Start()
        Dim PS() As Process = Process.GetProcessesByName("NLCM")
        If PS.Length = 0 Then
            Label1.Text = "Copying Files Services."
            CopyFilesToLocal()
            Label1.Text = "Done"
            P = Process.Start(_localfolder & "\" & "NLCM.exe")
            'P.EnableRaisingEvents = True
        Else
            Label1.Text = "Kill Services."
            For Each proc As Process In PS
                proc.Kill()
            Next
            System.Threading.Thread.Sleep(500)
            CopyFilesToLocal()
            Label1.Text = "Done"
            P = Process.Start(_localfolder & "\" & "NLCM.exe")
        End If

    End Sub

    Private Sub CopyFilesToLocal()
        If Not Directory.Exists(_localfolder) Then
            Directory.CreateDirectory(_localfolder)
            For Each file As IO.FileInfo In New IO.DirectoryInfo(_folderserver).GetFiles
                If file.Name <> "Thumbs.db" Then
                    file.CopyTo(_localfolder & "\" & file.Name)
                End If
            Next
        Else
            For Each file As IO.FileInfo In New IO.DirectoryInfo(_localfolder).GetFiles
                If file.Name <> "Thumbs.db" Then
                    System.IO.File.Delete(_localfolder & "\" & file.Name)
                End If
            Next
            For Each file As IO.FileInfo In New IO.DirectoryInfo(_folderserver).GetFiles
                If file.Name <> "Thumbs.db" Then
                    file.CopyTo(_localfolder & "\" & file.Name)
                End If
            Next
        End If

    End Sub

End Class
