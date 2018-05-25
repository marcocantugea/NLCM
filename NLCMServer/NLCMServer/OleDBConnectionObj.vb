Imports System.Data.OleDb
Imports System.Configuration
Imports System.Collections

Namespace com.data
    Public Class OleDBConnectionObj
        Protected connections As New Collection
        Protected connection As com.data.ConnectionsProperty

        Public Sub New()
            Dim cont As Boolean = False
            For Each s As String In System.Configuration.ConfigurationSettings.AppSettings
                If s.Contains("DB-") Then
                    cont = True
                    Dim con As New com.data.ConnectionsProperty
                    con.Name = s
                    con.ConectionString = System.Configuration.ConfigurationSettings.AppSettings(s)
                    connections.Add(con, con.Name)
                End If
            Next
            If Not cont Then
                Throw New Exception("There no any Database configure please configure at least 1 database.")
            End If
        End Sub

        Public Sub New(ByVal ConnectionProperty As com.data.ConnectionsProperty)
            connections.Add(ConnectionProperty, ConnectionProperty.Name)
        End Sub

        Protected Sub OpenDB(ByVal DB As String)
            Try
                connection = connections.Item(DB)
                connection.Connection.Open()
            Catch ex As Exception
                Throw
            End Try
        End Sub

        Protected Sub CreateConnection(ByVal ConnectionName As String, ByVal DatabaseToOpen As String)
            Dim con As New com.data.ConnectionsProperty
            con.Name = ConnectionName
            con.ConectionString = System.Configuration.ConfigurationSettings.AppSettings(DatabaseToOpen)
            connections.Add(con, con.Name)
        End Sub

        Protected Sub CloseDB()
            Try
                If Not IsNothing(connection.Adap) Then
                    connection.Adap.Dispose()
                End If
                If Not IsNothing(connection.Command) Then
                    connection.Command.Dispose()
                End If
                If Not IsNothing(connection.Connection) Then
                    connection.Connection.Close()
                End If
            Catch ex As Exception

            End Try
        End Sub

        Protected Sub CloseDB(ByVal DB As String)
            Try
                connection = connections.Item(DB)

                If Not IsNothing(connection.Adap) Then
                    connection.Adap.Dispose()
                End If
                If Not IsNothing(connection.Command) Then
                    connection.Command.Dispose()
                End If
                If Not IsNothing(connection.Connection) Then
                    connection.Connection.Close()
                End If
                connections.Remove(DB)
            Catch ex As Exception

            End Try
        End Sub

    End Class
End Namespace