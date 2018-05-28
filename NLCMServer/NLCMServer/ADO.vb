Imports System.Collections

Public Class ADO
    Inherits com.data.OleDBConnectionObj

    Public Sub RecordIP(record As IPStatisticsObj)
        Dim qbuilder As New QueryBuilder(Of IPStatisticsObj)
        qbuilder.TypeQuery = TypeQuery.Insert
        qbuilder.Entity = record
        qbuilder.BuildInsert("IPStatistics")
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
            connection.Command.ExecuteNonQuery()
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try
    End Sub

    Public Function CheckIPExist(IPAddres As String) As Boolean
        Dim result As Boolean = False
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand("SELECT IP_ADDRESS FROM IPTABLE WHERE IP_ADDRESS='" & IPAddres & "'", connection.Connection)
            Dim ip As String
            ip = connection.Command.ExecuteScalar
            If Not IsNothing(ip) Then
                result = True
            End If
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try

        Return result
    End Function

    Public Function CheckmMonitorIsEnable(IPAddres As String) As Boolean
        Dim result As Boolean = True
        Dim existipaddresontable As Boolean = CheckIPExist(IPAddres)
        If existipaddresontable Then
            Try
                OpenDB("DB-NLCMDB")
                connection.Command = New OleDb.OleDbCommand("SELECT MONITOR FROM IPTABLE WHERE IP_ADDRESS='" & IPAddres & "'", connection.Connection)
                If Not IsDBNull(connection.Command.ExecuteScalar) Then
                    result = connection.Command.ExecuteScalar
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End If
        Return result
    End Function
    Public Function GetIntervalRec(IPAddres As String) As Integer
        Dim result As Integer = 1
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand("SELECT Interval_REC FROM IPTABLE WHERE IP_ADDRESS='" & IPAddres & "'", connection.Connection)
            If Not IsDBNull(connection.Command.ExecuteScalar) Then
                result = connection.Command.ExecuteScalar
            End If
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try

        Return result
    End Function

    Public Function GetOnlineParameter(IPAddres As String) As Boolean
        Dim result As Boolean = False
        Dim existipaddresontable As Boolean = CheckIPExist(IPAddres)
        If existipaddresontable Then
            Try
                OpenDB("DB-NLCMDB")
                connection.Command = New OleDb.OleDbCommand("SELECT ONLINE FROM IPTABLE WHERE IP_ADDRESS='" & IPAddres & "'", connection.Connection)
                result = connection.Command.ExecuteScalar
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End If
        Return result
    End Function

    Public Sub GetParameter(Config As ConfigObj, parametertolook As String)
        Dim qbuilder As New QueryBuilder(Of ConfigObj)
        qbuilder.TypeQuery = TypeQuery.SelectInfo
        qbuilder.Entity = Config
        qbuilder.BuildSelect("ConfigParameters")
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand(qbuilder.Query & " WHERE parameter='" & parametertolook & "'", connection.Connection)
            connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
            Dim dts As New DataSet
            connection.Adap.Fill(dts)

            If dts.Tables.Count > 0 Then
                If dts.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In dts.Tables(0).Rows
                        'Dim o_ddr As New DDRControl
                        For Each member In Config.GetType.GetProperties
                            If member.CanWrite Then
                                If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                    If Not IsDBNull(row(member.Name)) Then
                                        member.SetValue(Config, row(member.Name), Nothing)
                                    End If
                                End If
                            End If
                        Next
                        'ddr.Add(o_ddr)
                    Next
                End If
            End If
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try
    End Sub

    Public Sub RecordIPInfo(record As IPInfo)
        Dim qbuilder As New QueryBuilder(Of IPInfo)
        qbuilder.TypeQuery = TypeQuery.Insert
        qbuilder.Entity = record
        qbuilder.BuildInsert("IPTable")
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
            connection.Command.ExecuteNonQuery()
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try
    End Sub

    Public Sub ChangeOnlineIP(IP As IPInfo)
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand("UPDATE IPTABLE SET ONLINE=" & IP.ONLINE & " WHERE IP_ADDRESS='" & IP.IP_ADDRESS & "'", connection.Connection)
            connection.Command.ExecuteNonQuery()
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try
    End Sub

    Public Sub ChangeMonitorIP(IP As IPInfo)
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand("UPDATE IPTABLE SET MONITOR=" & IP.MONITOR & " WHERE IP_ADDRESS='" & IP.IP_ADDRESS & "'", connection.Connection)
            connection.Command.ExecuteNonQuery()
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try
    End Sub

    Public Sub ChangeIntervalIP(IP As IPInfo)
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand("UPDATE IPTABLE SET Interval_REC=" & IP.Interval_REC.ToString & " WHERE IP_ADDRESS='" & IP.IP_ADDRESS & "'", connection.Connection)
            connection.Command.ExecuteNonQuery()
        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try
    End Sub

    'Added 27 May 2018
    'Get The IP Table on memory

    Public Sub GetIPTable(table As Dictionary(Of String, IPInfo))
        Try
            OpenDB("DB-NLCMDB")
            connection.Command = New OleDb.OleDbCommand("select * from IPTable", connection.Connection)
            connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
            Dim dts As New DataSet
            connection.Adap.Fill(dts)

            If dts.Tables.Count > 0 Then
                If dts.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In dts.Tables(0).Rows
                        Dim r_PCinfo As New IPInfo
                        r_PCinfo.IDIP = row(0)
                        r_PCinfo.IP_ADDRESS = row(1)
                        r_PCinfo.LOCATION = row(2)
                        r_PCinfo.ONLINE = row(3)
                        r_PCinfo.ACTIVE = row(4)
                        r_PCinfo.MONITOR = row(5)
                        r_PCinfo.Interval_REC = row(6)
                        r_PCinfo.ADAPTERNAME = row(7)
                        r_PCinfo.MACADDRESS = row(8)
                        Try
                            table.Add(row(1), r_PCinfo)
                        Catch ex As Exception

                        End Try

                    Next
                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            CloseDB()
        End Try
    End Sub


End Class
