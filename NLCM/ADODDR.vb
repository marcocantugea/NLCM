Imports DDRReportToolCore.com.entities

Namespace com.ADO
    Public Class ADODDR
        Inherits com.data.OleDBConnectionObj

        Public Sub SaveAllDDR(ByVal ddr As DDRControl)
            'Modificacion 18 Jul 2016
            ' Se modifico en validad que no se agrege un DDR con la misma fecha existente
            If ValidateDDRDate(ddr.ReportDate, ddr.Well) Then
                Try

                    If ddr.DDRID = -1 Then
                        SaveDDRControl(ddr)
                        ddr.DDRID = GetLastID("DDR_Control", "DDRID")
                        ddr.DDRReport.DDRID = ddr.DDRID
                    End If
                    If Not IsNothing(ddr.DDRReport) Then
                        UpdateDateAndReportNo(ddr.ReportNo, ddr.ReportDate, ddr.DDRID)
                        ddr.DDRReport.DDRID = ddr.DDRID
                        SaveDDRReport(ddr.DDRReport)


                        'Save DRR hrs
                        If Not IsNothing(ddr.DDRReport.DDRHrs) Then
                            For Each item As com.entities.DDRHrs In ddr.DDRReport.DDRHrs.Items
                                item.DDR_Report_ID = ddr.DDRID
                                SaveDDR_Hrs(item)
                            Next
                        End If

                        'Save BITS
                        If Not IsNothing(ddr.DDRReport.BITS) Then
                            For Each bit As com.entities.BITS In ddr.DDRReport.BITS.Items
                                bit.DDR_Report_ID = ddr.DDRID
                                SaveBits(bit)
                            Next
                        End If

                        'Save Drill String
                        If Not IsNothing(ddr.DDRReport.DrillString) Then
                            For Each bit As com.entities.DrillString In ddr.DDRReport.DrillString.Items
                                bit.DDR_Report_ID = ddr.DDRID
                                bit.DrillString_ID = -1
                                SaveDrillString(bit)
                            Next
                        End If

                        'Save Drill String survey
                        If Not IsNothing(ddr.DDRReport.DrillString_Survey) Then
                            For Each item As com.entities.DrillString_Survey In ddr.DDRReport.DrillString_Survey.Items
                                item.DDR_Report_ID = ddr.DDRID
                                item.Survey_ID = -1
                                SaveDrillString_survey(item)
                            Next
                        End If

                        'Save pumps
                        If Not IsNothing(ddr.DDRReport.Pumps) Then
                            For Each item As com.entities.Pumps In ddr.DDRReport.Pumps.Items
                                item.DDR_Report_ID = ddr.DDRID
                                item.PUMPS_ID = -1
                                SavePumps(item)
                            Next
                        End If

                        'Save shakers
                        If Not IsNothing(ddr.DDRReport.Shakers) Then
                            For Each item As com.entities.Shakers In ddr.DDRReport.Shakers.Items
                                item.DDR_Report_ID = ddr.DDRID
                                item.Shakers_ID = -1
                                SaveShakers(item)
                            Next
                        End If

                        'Save Mud
                        If Not IsNothing(ddr.DDRReport.Mud) Then
                            For Each item As com.entities.Mud In ddr.DDRReport.Mud.Items
                                item.DDR_Report_ID = ddr.DDRID
                                item.MUD_ID = -1
                                SaveMud(item)
                            Next
                        End If

                        'Save Marine Info
                        If Not IsNothing(ddr.DDRReport.MarineInfo) Then
                            ddr.DDRReport.MarineInfo.DDR_Report_ID = ddr.DDRID
                            SaveMarineInfo(ddr.DDRReport.MarineInfo)

                        End If

                        'Save POB
                        If Not IsNothing(ddr.DDRReport.POB) Then
                            ddr.DDRReport.POB.DDR_Report_ID = ddr.DDRID
                            SavePOB(ddr.DDRReport.POB)
                        End If

                        'Save Riser Profile
                        If Not IsNothing(ddr.DDRReport.RiserProfile) Then
                            For Each item As RiserProfile In ddr.DDRReport.RiserProfile.Items
                                item.DDR_Report_ID = ddr.DDRID
                                item.IDRiserProfile = -1
                                SaveRiserProfile(item)
                            Next
                        End If

                        'Save SOC
                        If Not IsNothing(ddr.DDRReport.SOC) Then
                            ddr.DDRReport.SOC.DDR_Report_ID = ddr.DDRID
                            SaveSOC(ddr.DDRReport.SOC)
                        End If

                        'Save Logistic Transit Log
                        If Not IsNothing(ddr.DDRReport.LogisticTransitLog) Then
                            For Each item As LogisticTransitLog In ddr.DDRReport.LogisticTransitLog.items
                                item.DDR_Report_ID = ddr.DDRID
                                item.LTID = -1
                                SaveLogisticTransitLog(item)
                            Next
                        End If

                        'Save Urgents MRs
                        If Not IsNothing(ddr.DDRReport.UrgentsMR) Then
                            For Each item As UrgentMRs In ddr.DDRReport.UrgentsMR.items
                                If item.MRUrgentID = -1 Then
                                    item.DDR_Report_ID = ddr.DDRID
                                    SaveUrgentMRs(item)
                                End If
                            Next
                        End If

                    End If
                    MsgBox("DDR Saved.")
                Catch ex As Exception

                    Throw
                End Try
            Else
                Throw New Exception("A DDR with the same date was found. Can not create a DDR with the same date.")
            End If

        End Sub

        
        Public Sub ModifyALLDDR(ByVal ddr As DDRControl)
            'DeleteDDR_Report(ddr.DDRID)
            'SaveAllDDR(ddr)

            UpdateDDRControl(ddr)

            If Not IsNothing(ddr.DDRReport) Then
                UpdateDDRReport(ddr.DDRReport)
            End If
            'Save DRR hrs
            If Not IsNothing(ddr.DDRReport.DDRHrs) Then
                For Each item As com.entities.DDRHrs In ddr.DDRReport.DDRHrs.Items
                    item.DDR_Report_ID = ddr.DDRID
                    If item.Detail_HR_ID = -1 Then
                        SaveDDR_Hrs(item)
                    Else
                        UpdateDDRHrs(item)
                    End If

                Next
            End If

            'Save BITS
            If Not IsNothing(ddr.DDRReport.BITS) Then
                For Each bit As com.entities.BITS In ddr.DDRReport.BITS.Items
                    bit.DDR_Report_ID = ddr.DDRID
                    If bit.BITS_ID = -1 Then
                        SaveBits(bit)
                    Else
                        UpdateBITS(bit)
                    End If

                Next
            End If

            'Save Drill String
            If Not IsNothing(ddr.DDRReport.DrillString) Then
                For Each bit As com.entities.DrillString In ddr.DDRReport.DrillString.Items
                    bit.DDR_Report_ID = ddr.DDRID
                    UpdateDrillString(bit)
                Next
            End If

            'Save Drill String survey
            If Not IsNothing(ddr.DDRReport.DrillString_Survey) Then
                For Each item As com.entities.DrillString_Survey In ddr.DDRReport.DrillString_Survey.Items
                    item.DDR_Report_ID = ddr.DDRID
                    If item.Survey_ID = -1 Then
                        SaveDrillString_survey(item)
                    Else
                        UpdateDrillStringSurvey(item)
                    End If

                Next
            End If

            'Save pumps
            If Not IsNothing(ddr.DDRReport.Pumps) Then
                For Each item As com.entities.Pumps In ddr.DDRReport.Pumps.Items
                    item.DDR_Report_ID = ddr.DDRID
                    UpdatePumps(item)
                Next
            End If

            'Save shakers
            If Not IsNothing(ddr.DDRReport.Shakers) Then
                For Each item As com.entities.Shakers In ddr.DDRReport.Shakers.Items
                    item.DDR_Report_ID = ddr.DDRID
                    If item.Shakers_ID = -1 Then
                        SaveShakers(item)
                    Else
                        UpdateShakers(item)
                    End If

                Next
            End If

            'Save Mud
            If Not IsNothing(ddr.DDRReport.Mud) Then
                For Each item As com.entities.Mud In ddr.DDRReport.Mud.Items
                    item.DDR_Report_ID = ddr.DDRID
                    If item.MUD_ID = -1 Then
                        SaveMud(item)
                    Else
                        UpdateMud(item)
                    End If

                Next
            End If

            'Save Marine Info
            If Not IsNothing(ddr.DDRReport.MarineInfo) Then
                ddr.DDRReport.MarineInfo.DDR_Report_ID = ddr.DDRID
                UpdateMarineinfo(ddr.DDRReport.MarineInfo)

            End If

            'Save POB
            If Not IsNothing(ddr.DDRReport.POB) Then
                ddr.DDRReport.POB.DDR_Report_ID = ddr.DDRID
                If ddr.DDRReport.POB.POB_ID = -1 Then
                    SavePOB(ddr.DDRReport.POB)
                Else
                    UpdatePOB(ddr.DDRReport.POB)
                End If

            End If

            'Save Riser Profile
            If Not IsNothing(ddr.DDRReport.RiserProfile) Then
                For Each item As RiserProfile In ddr.DDRReport.RiserProfile.Items
                    item.DDR_Report_ID = ddr.DDRID
                    If item.IDRiserProfile = -1 Then
                        SaveRiserProfile(item)
                    Else
                        UpdateRiserProfile(item)
                    End If

                Next
            End If

            'Save SOC
            If Not IsNothing(ddr.DDRReport.SOC) Then
                ddr.DDRReport.SOC.DDR_Report_ID = ddr.DDRID
                UpdateSOC(ddr.DDRReport.SOC)
            End If

            'Save Logistic Transit Log
            If Not IsNothing(ddr.DDRReport.LogisticTransitLog) Then
                For Each item As LogisticTransitLog In ddr.DDRReport.LogisticTransitLog.items
                    item.DDR_Report_ID = ddr.DDRID
                    If item.LTID = -1 Then
                        SaveLogisticTransitLog(item)
                    Else
                        UpdateLogisticTransitLog(item)
                    End If

                Next
            End If

        End Sub

        Public Function GetCompleteDDRReport(ByVal DDRID As Integer) As DDRControl
            Dim ddrc As New DDRControl
            ddrc.DDRID = DDRID
            ddrc = GetDDRControlHeader(DDRID)
            ddrc.DDRReport = GetDDRReport(DDRID)
            ddrc.DDRReport.DDRHrs = GetDDRHrs(DDRID)
            ddrc.DDRReport.BITS = GetDDRBits(DDRID)
            ddrc.DDRReport.DrillString = GetDrillString(DDRID)
            ddrc.DDRReport.DrillString_Survey = GetDrillStringSurvey(DDRID)
            ddrc.DDRReport.MarineInfo = GetMarineInfo(DDRID)
            ddrc.DDRReport.POB = GetPOB(DDRID)
            ddrc.DDRReport.Pumps = GetPumps(DDRID)
            ddrc.DDRReport.Shakers = GetShakers(DDRID)
            ddrc.DDRReport.Mud = GetMud(DDRID)
            ddrc.DDRReport.Activities = GetActivities(DDRID)
            ddrc.DDRReport.RiserProfile = GetRiserProfile(DDRID)
            ddrc.DDRReport.SOC = GetSOC(DDRID)
            ddrc.DDRReport.LogisticTransitLog = GetLogisticTransitLog(DDRID)
            ddrc.DDRReport.UrgentsMR = GetUrgentsMR(DDRID)
            ddrc.DDRReport.WorkOrders = GetWO(DDRID)
            ddrc.DDRReport.PUMR = GetPUMR(DDRID)
            Return ddrc
        End Function

#Region "Save Info Functions"

        Public Sub SaveDDRControl(ByVal ddrcontrol As DDRControl)
            Dim qbuilder As New QueryBuilder(Of DDRControl)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = ddrcontrol
            qbuilder.BuildInsert("DDR_Control")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try

        End Sub
        Public Sub SaveDDRReport(ByVal ddr_report As DDRReport)
            Dim qbuilder As New QueryBuilder(Of DDRReport)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = ddr_report
            qbuilder.BuildInsert("DDR_Report")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            ddr_report.DDR_Report_ID = GetLastID("DDR_Report", "DDR_Report_ID")
        End Sub
        Public Sub SaveDDR_Bits(ByVal ddr_bits As BITS)
            Dim qbuilder As New QueryBuilder(Of BITS)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = ddr_bits
            qbuilder.BuildInsert("DDR_BITS")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            ddr_bits.BITS_ID = GetLastID("DDR_BITS", "BITS_ID")
        End Sub
        Public Sub SaveDDR_Hrs(ByVal ddrhrs As com.entities.DDRHrs)
            Dim qbuilder As New QueryBuilder(Of DDRHrs)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = ddrhrs
            qbuilder.BuildInsert("DDR_Detail_Hrs")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            ddrhrs.Detail_HR_ID = GetLastID("DDR_Detail_Hrs", "Detail_Hr_ID")
        End Sub
        Public Sub SaveBits(ByVal bits As com.entities.BITS)
            Dim qbuilder As New QueryBuilder(Of BITS)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = bits
            qbuilder.BuildInsert("DDR_BITS")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            bits.BITS_ID = GetLastID("DDR_BITS", "BITS_ID")
        End Sub
        Public Sub SaveDrillString(ByVal drillstring As com.entities.DrillString)
            Dim qbuilder As New QueryBuilder(Of DrillString)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = drillstring
            qbuilder.BuildInsert("DDR_DrillString")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            drillstring.DrillString_ID = GetLastID("DDR_DrillString", "DrillString_ID")
        End Sub
        Public Sub SaveDrillString_survey(ByVal drillstring_survey As com.entities.DrillString_Survey)
            Dim qbuilder As New QueryBuilder(Of DrillString_Survey)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = drillstring_survey
            qbuilder.BuildInsert("DDR_DrillString_Surveys")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            drillstring_survey.Survey_ID = GetLastID("DDR_DrillString_Surveys", "Survey_ID")
        End Sub
        Public Sub SavePumps(ByVal pumps As com.entities.Pumps)
            Dim qbuilder As New QueryBuilder(Of Pumps)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = pumps
            qbuilder.BuildInsert("DDR_PUMPS")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            pumps.pumps_id = GetLastID("DDR_PUMPS", "PUMPS_ID")
        End Sub
        Public Sub SaveShakers(ByVal shakers As com.entities.Shakers)
            Dim qbuilder As New QueryBuilder(Of Shakers)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = shakers
            qbuilder.BuildInsert("DDR_Shakers")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            shakers.Shakers_ID = GetLastID("DDR_Shakers", "Shakers_ID")
        End Sub
        Public Sub SaveMud(ByVal muds As com.entities.Mud)
            Dim qbuilder As New QueryBuilder(Of Mud)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = muds
            qbuilder.BuildInsert("DDR_Mud")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            muds.MUD_ID = GetLastID("DDR_Mud", "MUD_ID")
        End Sub
        Public Sub SaveMarineInfo(ByVal marine As com.entities.MarineInfo)
            Dim qbuilder As New QueryBuilder(Of MarineInfo)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = marine
            qbuilder.BuildInsert("DDR_Marine")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            marine.Marine_ID = GetLastID("DDR_Marine", "Marine_ID")
        End Sub
        Public Sub SavePOB(ByVal pob As com.entities.POB)
            Dim qbuilder As New QueryBuilder(Of POB)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = pob
            qbuilder.BuildInsert("DDR_POB")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            pob.POB_ID = GetLastID("DDR_POB", "POB_ID")
        End Sub
        

        Public Sub SaveRiserProfile(ByVal riserprof As com.entities.RiserProfile)
            Dim qbuilder As New QueryBuilder(Of RiserProfile)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = riserprof
            qbuilder.BuildInsert("RiserProfile")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            riserprof.IDRiserProfile = GetLastID("RiserProfile", "IDRiserProfile")
        End Sub
        Public Sub SaveSOC(ByVal socdata As com.entities.SOC)
            Dim qbuilder As New QueryBuilder(Of SOC)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = socdata
            qbuilder.BuildInsert("DDR_SOC")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            socdata.SOCINFOID = GetLastID("DDR_SOC", "SOCINFOID")
        End Sub

        Public Sub SaveLogisticTransitLog(ByVal TransitLog As com.entities.LogisticTransitLog)
            Dim qbuilder As New QueryBuilder(Of LogisticTransitLog)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = TransitLog
            qbuilder.BuildInsert("DDR_LogisticTransitLog")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            TransitLog.LTID = GetLastID("DDR_LogisticTransitLog", "LTID")
        End Sub



        Public Function DebugQueryBuilder()
            Dim ddr1 As New DDRControl
            ddr1.ReportDate = Today()
            ddr1.Description = "-7"
            ddr1.ReportNo = "-7"
            ddr1.DDRID = -7

            Dim qbuilder As New QueryBuilder(Of DDRControl)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = ddr1

            'qbuilder.BuildInsert("DDR_Report")
            'qbuilder.BuildUpdate("DDR_Report", "DDR_Report_ID", "99")
            qbuilder.AddToQueryParameterForSelect("DDR_Report_ID=99")
            qbuilder.BuildSelect("DDR_Report")

            Return qbuilder.Query

        End Function
#End Region

        Private Sub DeleteDDR_Report(ByVal DDR_ID As Integer)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Report where DDRID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_BITS where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Detail_Hrs where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_DrillString where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_DrillString_Surveys where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Marine where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Mud where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_POB where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_PUMPS where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Shakers where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from RiserProfile where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_SOC where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_LogisticTransitLog where DDR_Report_ID=" & DDR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

#Region "Geters Info"

        'Modificado el 18 de Jul 2016
        ' Se agrego funcion de validar si la fecha del ddr que se desea grabar ya existe
        Public Function ValidateDDRDate(ByVal fechaddr As Date, ByVal well As String) As Boolean
            Dim lastddt As Integer = GetLastIDDDR(well)
            Dim ddrcontro As New DDRControl
            Dim validation As Boolean = True
            ddrcontro = GetDDRControlHeader(lastddt)

            If ddrcontro.ReportDate.ToString("MM/dd/yy") = fechaddr.ToString("MM/dd/yy") Then
                validation = False
            End If

            Return validation
        End Function

        'Modificado el 18 de Jul 2016
        ' se Agrego la funcion de buscar el ultimo reporte por pozo
        Public Function GetLastIDDDR(ByVal Well As String) As Integer
            Dim result As Integer = -1
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select max(DDRID) from DDR_CONTROL where Well='" & Well & "'", connection.Connection)
                If Not IsDBNull(connection.Command.ExecuteScalar()) Then
                    result = connection.Command.ExecuteScalar()
                Else
                    result = 1
                End If

            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try

            Return result
        End Function


        Public Function GetLastID(ByVal table As String, ByVal field As String) As Integer
            Dim result As Integer = -1
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select max(" & field & ") from " & table, connection.Connection)
                result = connection.Command.ExecuteScalar()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try

            Return result
        End Function
        Public Function GetLastDDRUpdate(ByVal ddrid As String) As Date
            Dim result As Date
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select lastupdate from DDR_Control where DDRID=" & ddrid, connection.Connection)
                result = connection.Command.ExecuteScalar()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try

            Return result
        End Function
        Public Sub GetDDRControlHeader(ByVal ddr As DDRControl_Collection)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_Control", connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim o_ddr As New DDRControl
                            For Each member In o_ddr.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            member.SetValue(o_ddr, row(member.Name), Nothing)
                                        End If
                                    End If
                                End If
                            Next
                            ddr.Add(o_ddr)
                        Next
                    End If
                End If

            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Sub GetWells(ByVal wells As Collection)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select distinct Well from DDR_Control", connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            If Not IsDBNull(row(0)) Then
                                wells.Add(row(0))
                            End If

                        Next
                    End If
                End If

            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Function GetDDRControlHeader(ByVal DDRID As Integer) As DDRControl
            Dim o_ddr As New DDRControl
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_Control where DDRID=" & DDRID.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            For Each member In o_ddr.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            member.SetValue(o_ddr, row(member.Name), Nothing)
                                        End If
                                    End If
                                End If
                            Next

                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return o_ddr
        End Function
        Public Function GetDDRReport(ByVal ddrid As Integer) As DDRReport
            Dim ddr_r As New DDRReport
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_Report where DDRID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows

                            For Each member In ddr_r.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(ddr_r, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(ddr_r, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "DateTime" Then
                                                member.SetValue(ddr_r, Date.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return ddr_r
        End Function
        Public Function GetDDRHrs(ByVal ddrid As Integer) As DDRHrs_Collection
            Dim ddrhrs_collected As New DDRHrs_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_Detail_Hrs where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim o_ddrhrs As New DDRHrs
                            For Each member In o_ddrhrs.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(o_ddrhrs, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(o_ddrhrs, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            ddrhrs_collected.Add(o_ddrhrs)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return ddrhrs_collected
        End Function
        Public Function GetDDRBits(ByVal ddrid As Integer) As BITS_Collection
            Dim bits_collected As New BITS_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_BITS where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim bits As New BITS
                            For Each member In bits.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(bits, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(bits, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            bits_collected.Add(bits)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return bits_collected
        End Function
        Public Function GetDrillString(ByVal ddrid As Integer) As DrillString_Collection
            Dim drillstring_collected As New DrillString_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_DrillString where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim drillstring As New DrillString
                            For Each member In drillstring.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(drillstring, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(drillstring, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            drillstring_collected.Add(drillstring)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return drillstring_collected
        End Function
        Public Function GetDrillStringSurvey(ByVal ddrid As Integer) As DrillString_Survey_Collection
            Dim drillstringsurvey_collected As New DrillString_Survey_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_DrillString_Surveys where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim drillstringsurvey As New DrillString_Survey
                            For Each member In drillstringsurvey.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(drillstringsurvey, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(drillstringsurvey, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            drillstringsurvey_collected.Add(drillstringsurvey)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return drillstringsurvey_collected
        End Function
        Public Function GetMarineInfo(ByVal ddrid As Integer) As MarineInfo
            Dim marineinfo As New MarineInfo
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_Marine where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            For Each member In marineinfo.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(marineinfo, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(marineinfo, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(marineinfo, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "DateTime" Then
                                                member.SetValue(marineinfo, Date.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return marineinfo
        End Function
        Public Function GetPOB(ByVal ddrid As Integer) As POB
            Dim POBC As New POB
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_POB where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            For Each member In POBC.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(POBC, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "DateTime" Then
                                                member.SetValue(POBC, row(member.Name), Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(POBC, row(member.Name), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return POBC
        End Function
        Public Function GetPumps(ByVal ddrid As Integer) As Pumps_Collection
            Dim pumps_c As New Pumps_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_PUMPS where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim pumps As New Pumps
                            For Each member In pumps.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(pumps, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(pumps, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            pumps_c.Add(pumps)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return pumps_c
        End Function
        Public Function GetShakers(ByVal ddrid As Integer) As Shakers_Collection
            Dim shakers As New Shakers_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_Shakers where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim shaker As New Shakers
                            For Each member In shaker.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(shaker, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(shaker, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            shakers.Add(shaker)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return shakers
        End Function
        Public Function GetMud(ByVal ddrid As Integer) As Mud_Collection
            Dim muds As New Mud_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_Mud where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim mud As New Mud
                            For Each member In mud.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(mud, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(mud, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            muds.Add(mud)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return muds
        End Function
        Public Function GetActivities(ByVal ddrid As Integer) As Activities_Collection
            Dim activities As New Activities_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from Activities_Details where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim activity As New Activities
                            For Each member In activity.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(activity, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(activity, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            activities.Add(activity)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return activities
        End Function
        Public Function GetRiserProfile(ByVal ddrid As Integer) As RiserProfileCollection
            Dim risersProfiles As New RiserProfileCollection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from RiserProfile where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim rp As New RiserProfile
                            For Each member In rp.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(rp, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(rp, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            risersProfiles.Add(rp)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return risersProfiles
        End Function
        Public Function GetSOC(ByVal ddrid As Integer) As SOC
            Dim socdata As New SOC
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_SOC where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)
                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            For Each member In socdata.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(socdata, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(socdata, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return socdata
        End Function
        Public Function GetLogisticTransitLog(ByVal ddrid As Integer) As LogisticTransitLogCollection
            Dim transitlog As New LogisticTransitLogCollection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_LogisticTransitLog where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim logtransit As New LogisticTransitLog
                            For Each member In logtransit.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(logtransit, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(logtransit, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "Boolean" Then
                                                member.SetValue(logtransit, Boolean.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            transitlog.Add(logtransit)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return transitlog
        End Function

        Public Function GetUrgentsMR(ByVal ddrid As Integer) As UrgentsMRsCollection
            Dim mrs As New UrgentsMRsCollection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from Activities_UrgentMRs where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim mr As New UrgentMRs
                            For Each member In mr.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(mr, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(mr, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            mrs.Add(mr)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return mrs
        End Function

        Public Function GetWO(ByVal ddrid As Integer) As WorkOrderCollection
            Dim wos As New WorkOrderCollection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from Activities_WorkOrders where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim wo As New WorkOrder
                            For Each member In wo.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(wo, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(wo, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "Boolean" Then
                                                member.SetValue(wo, Boolean.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            wos.Add(wo)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return wos
        End Function

        Public Function GetPUMR(ByVal ddrid As Integer) As PUMR_Collection
            Dim pumrs As New PUMR_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from DDR_PUMR where DDR_Report_ID=" & ddrid.ToString, connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim pumr As New PUMR
                            For Each member In pumr.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(pumr, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(pumr, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            pumrs.Add(pumr)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return pumrs
        End Function


#End Region

#Region "Modifiers"
        Public Sub UpdateDDRControl(ByVal DDRc As DDRControl)
            Dim qbuilder As New QueryBuilder(Of DDRControl)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = DDRc
            qbuilder.BuildUpdate("DDR_Control", "DDRID", DDRc.DDRID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateDDRReport(ByVal DDR As DDRReport)
            Dim qbuilder As New QueryBuilder(Of DDRReport)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = DDR
            qbuilder.BuildUpdate("DDR_Report", "DDR_Report_ID", DDR.DDR_Report_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateDDRHrs(ByVal ddrhrs As DDRHrs)
            Dim qbuilder As New QueryBuilder(Of DDRHrs)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = ddrhrs
            qbuilder.BuildUpdate("DDR_Detail_Hrs", "Detail_HR_ID", ddrhrs.Detail_HR_ID, True)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateBITS(ByVal bits As BITS)
            Dim qbuilder As New QueryBuilder(Of BITS)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = bits
            qbuilder.BuildUpdate("DDR_BITS", "BITS_ID", bits.BITS_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateDrillString(ByVal drilllstring As DrillString)
            Dim qbuilder As New QueryBuilder(Of DrillString)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = drilllstring
            qbuilder.BuildUpdate("DDR_DrillString", "DrillString_ID", drilllstring.DrillString_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateDrillStringSurvey(ByVal drilllstringsur As DrillString_Survey)
            Dim qbuilder As New QueryBuilder(Of DrillString_Survey)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = drilllstringsur
            qbuilder.BuildUpdate("DDR_DrillString_Surveys", "Survey_ID", drilllstringsur.Survey_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdatePumps(ByVal pump As Pumps)
            Dim qbuilder As New QueryBuilder(Of Pumps)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = pump
            qbuilder.BuildUpdate("DDR_PUMPS", "PUMPS_ID", pump.PUMPS_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateShakers(ByVal shaker As Shakers)
            Dim qbuilder As New QueryBuilder(Of Shakers)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = shaker
            qbuilder.BuildUpdate("DDR_Shakers", "Shakers_ID", shaker.Shakers_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateMud(ByVal mud As Mud)
            Dim qbuilder As New QueryBuilder(Of Mud)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = mud
            qbuilder.BuildUpdate("DDR_Mud", "MUD_ID", mud.MUD_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateMarineinfo(ByVal marine As MarineInfo)
            Dim qbuilder As New QueryBuilder(Of MarineInfo)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = marine
            qbuilder.BuildUpdate("DDR_Marine", "Marine_ID", marine.Marine_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdatePOB(ByVal pob As POB)
            Dim qbuilder As New QueryBuilder(Of POB)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = pob
            qbuilder.BuildUpdate("DDR_POB", "POB_ID", pob.POB_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateRiserProfile(ByVal rp As RiserProfile)
            Dim qbuilder As New QueryBuilder(Of RiserProfile)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = rp
            qbuilder.BuildUpdate("RiserProfile", "IDRiserProfile", rp.IDRiserProfile)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateSOC(ByVal soc As SOC)
            Dim qbuilder As New QueryBuilder(Of SOC)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = soc
            qbuilder.BuildUpdate("DDR_SOC", "SOCINFOID", soc.SOCINFOID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateLogisticTransitLog(ByVal logtranist As LogisticTransitLog)
            Dim qbuilder As New QueryBuilder(Of LogisticTransitLog)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = logtranist
            qbuilder.BuildUpdate("DDR_LogisticTransitLog", "LTID", logtranist.LTID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        'Modificacion 22-Sep-2017
        'Actualizar El nombre del supervisor en la base de datos para imprimirlo en 
        ' el reporte
        Public Sub UpdateF1SupervisorName(DDRReportID As Integer, Supervisorname As String)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Update DDR_Report set  F1SupervisorName='" & Supervisorname & "' where DDR_Report_ID=" & DDRReportID.ToString, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()

            End Try
        End Sub
        'Modificacion 22-Sep-2017
        'Actualizar El nombre del superintendente en la base de datos para imprimirlo en 
        ' el reporte
        Public Sub UpdateF1SuperintendentName(DDRReportID As Integer, name As String)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Update DDR_Report set  F1RigSuperintName='" & name & "' where DDR_Report_ID=" & DDRReportID.ToString, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()

            End Try
        End Sub

#End Region


        Public Sub LockReprot(ByVal DDRID As Integer)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("update ddr_Control set Lock=-1 where DDRID=" & DDRID.ToString, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UnlockReprot(ByVal DDRID As Integer)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("update ddr_Control set Lock=0 where DDRID=" & DDRID.ToString, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub GetUserGroup(ByVal suser As com.entities.SessionUser)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select UserGroup from users_group where username='" & suser.User & "'", connection.Connection)
                If Not IsDBNull(connection.Command.ExecuteScalar()) Then
                    suser.Group = connection.Command.ExecuteScalar()
                Else
                    suser.Group = "View"
                End If

            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub GetUserDeparmentID(ByVal suser As com.entities.SessionUser)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select DepartmentID from users_group where username='" & suser.User & "'", connection.Connection)
                suser.DepartmentId = connection.Command.ExecuteScalar()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub GetUseremail(ByVal suser As com.entities.SessionUser)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select email from users_group where username='" & suser.User & "'", connection.Connection)
                suser.email = connection.Command.ExecuteScalar()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Function GetDeparmentID(ByVal DeparmentName As String) As Integer
            Dim deparmentid As Integer
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select Deparment_ID from Activities_Deparments where description='" & DeparmentName & "'", connection.Connection)
                deparmentid = connection.Command.ExecuteScalar()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return deparmentid
        End Function
        Public Sub GetUserDeparmentName(ByVal suser As com.entities.SessionUser)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select Description from Activities_Deparments where Deparment_ID=" & suser.DepartmentId, connection.Connection)
                suser.DeparmentName = connection.Command.ExecuteScalar()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub PrepareNotification(ByVal emailcollection As com.Notifier.Email.EmailObjCollection, ByVal templatemessage As com.Notifier.Email.EmailObj, ByVal sender As String)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select email from Users_Group where notify=-1", connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)
                Dim messageem As com.Notifier.Email.EmailObj
                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            If Not IsDBNull(row(0)) Then
                                messageem = New com.Notifier.Email.EmailObj
                                messageem.Body = templatemessage.Body
                                messageem.eTo = row("email").ToString
                                messageem.From = sender
                                messageem.HTMLBody = templatemessage.HTMLBody
                                messageem.Subject = templatemessage.Subject
                                emailcollection.Add(messageem)
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub UpdateDateAndReportNo(ByVal reportno As Integer, ByVal reportdate As Date, ByVal ddrid As Integer)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("update DDR_Control set ReportDate=#" & reportdate.ToString("MM/dd/yyyy") & "#,ReportNo=" & reportno & " where DDRID=" & ddrid, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub


#Region "Activities"


        Public Sub SaveActivities(ByVal ddr As DDRControl)
            If Not IsNothing(ddr.DDRReport.Activities) Then
                For Each item As com.entities.Activities In ddr.DDRReport.Activities.Items
                    SaveActivitie(item)
                Next
            End If

        End Sub

        Public Sub SaveActivitie(ByVal act As com.entities.Activities)
            Dim qbuilder As New QueryBuilder(Of Activities)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = act
            qbuilder.BuildInsert("Activities_Details")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()

            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try

            act.Act_Detail_ID = GetLastID("Activities_Details", "Act_Detail_ID")
        End Sub

        Public Sub ModifyActivities(ByVal ddr As DDRControl)
            If Not IsNothing(ddr.DDRReport.Activities) Then
                For Each item As com.entities.Activities In ddr.DDRReport.Activities.Items
                    UpdateActivitie(item)
                Next
                'DeleteActivities(ddr.DDRID)
                'SaveActivities(ddr)
            End If
        End Sub

        Public Sub UpdateActivitie(ByVal activity As Activities)
            Dim qbuilder As New QueryBuilder(Of Activities)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = activity
            qbuilder.BuildUpdate("Activities_Details", "Act_Detail_ID", activity.Act_Detail_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Sub DeleteActivities(ByVal DDRID As Integer)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from Activities_Details where DDR_Report_ID=" & DDRID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()

            End Try
        End Sub
        Public Sub DeleteActivities(ByVal activity As Activities)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from Activities_Details where Act_Detail_ID=" & activity.Act_Detail_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()

            End Try
        End Sub
#End Region

#Region "MR Urgents"

        Public Sub UpdateUrgentMR(ByVal MR As UrgentMRs)
            Dim qbuilder As New QueryBuilder(Of UrgentMRs)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = MR
            qbuilder.BuildUpdate("Activities_UrgentMRs", "MRUrgentID", MR.MRUrgentID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Sub SaveUrgentMRs(ByVal MR As com.entities.UrgentMRs)
            Dim qbuilder As New QueryBuilder(Of UrgentMRs)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = MR
            qbuilder.BuildInsert("Activities_UrgentMRs")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            MR.MRUrgentID = GetLastID("Activities_UrgentMRs", "MRUrgentID")
        End Sub

        Public Sub DeleteUrgentMR(ByVal MR As UrgentMRs)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from Activities_UrgentMRs where MRUrgentID=" & MR.MRUrgentID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

#End Region

#Region "PEMEX MR Urgents"
        Public Sub SavePUMR(ByVal MR As com.entities.PUMR)
            Dim qbuilder As New QueryBuilder(Of PUMR)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = MR
            qbuilder.BuildInsert("DDR_PUMR")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            MR.PRUM_ID = GetLastID("DDR_PUMR", "PRUM_ID")
        End Sub
        Public Sub UpdatePUMR(ByVal MR As PUMR)
            Dim qbuilder As New QueryBuilder(Of PUMR)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = MR
            qbuilder.BuildUpdate("DDR_PUMR", "PRUM_ID", MR.PRUM_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub DeletePUMR(ByVal MR As PUMR)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_PUMR where PRUM_ID=" & MR.PRUM_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
#End Region

#Region "Work Orders"

        Public Sub UpdateWorkOrder(ByVal WO As WorkOrder)
            Dim qbuilder As New QueryBuilder(Of WorkOrder)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = WO
            qbuilder.BuildUpdate("Activities_WorkOrders", "WorkOrderID", WO.WorkOrderID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Sub SaveWorkOrder(ByVal WO As WorkOrder)
            Dim qbuilder As New QueryBuilder(Of WorkOrder)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = WO
            qbuilder.BuildInsert("Activities_WorkOrders")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            WO.WorkOrderID = GetLastID("Activities_WorkOrders", "WorkOrderID")
        End Sub

        Public Sub DeleteWorkOrder(ByVal WO As WorkOrder)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from Activities_WorkOrders where WorkOrderID=" & WO.WorkOrderID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()

            End Try
        End Sub

#End Region

        Public Sub DeleteDDHrs(ByVal ddrhrs As DDRHrs)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Detail_Hrs where Detail_Hr_ID=" & ddrhrs.Detail_HR_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub


        Public Sub DeleteBITS(ByVal Bits As BITS)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_BITS where BITS_ID=" & Bits.BITS_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub DeleteShaker(ByVal shaker As Shakers)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Shakers where Shakers_ID=" & shaker.Shakers_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub
        Public Sub DeleteMud(ByVal mud As Mud)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_Mud where MUD_ID=" & mud.MUD_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub


        Public Sub UpdateDDRLastUpdate(ByVal DDR As DDRControl)
            Dim qbuilder As New QueryBuilder(Of DDRControl)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = DDR
            qbuilder.BuildUpdate("DDR_Control", "DDRID", DDR.DDRID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Sub SaveLogOpenTab(ByVal SysOpenTab As SystemOpenedTab)
            Dim qbuilder As New QueryBuilder(Of SystemOpenedTab)
            qbuilder.TypeQuery = TypeQuery.Insert
            qbuilder.Entity = SysOpenTab
            qbuilder.BuildInsert("System_openedTabs")
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            SysOpenTab.OpenedTab_ID = GetLastID("System_openedTabs", "OpenedTab_ID")
        End Sub

        Public Sub UpdateLogOpenTab(ByVal SysOpenTab As SystemOpenedTab)
            Dim qbuilder As New QueryBuilder(Of SystemOpenedTab)
            qbuilder.TypeQuery = TypeQuery.Update
            qbuilder.Entity = SysOpenTab
            qbuilder.BuildUpdate("System_openedTabs", "OpenedTab_ID", SysOpenTab.OpenedTab_ID)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand(qbuilder.Query, connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Function GetTabSelected(ByVal obj_tosearch As SystemOpenedTab) As SystemOpenedTab
            Dim sysopentab As New SystemOpenedTab
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from  System_openedTabs where Tab_sel='" & obj_tosearch.Tab_sel & "' and Active=yes", connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)
                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            For Each member In sysopentab.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(sysopentab, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(sysopentab, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "DateTime" Then
                                                member.SetValue(sysopentab, row(member.Name), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return sysopentab
        End Function
        Public Function GetTabSelectedUser(ByVal obj_tosearch As SystemOpenedTab) As SystemOpenedTab
            Dim sysopentab As New SystemOpenedTab
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from  System_openedTabs where User_sess='" & obj_tosearch.User_sess & "' and Active=yes", connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)
                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            For Each member In sysopentab.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(sysopentab, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(sysopentab, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "DateTime" Then
                                                member.SetValue(sysopentab, row(member.Name), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return sysopentab
        End Function

        Public Function GetOpenedTabs() As System_OpenedTab_Collection
            Dim tabsopeneds As New System_OpenedTab_Collection
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("select * from System_openedTabs where Active=yes", connection.Connection)
                connection.Adap = New OleDb.OleDbDataAdapter(connection.Command)
                Dim dts As New DataSet
                connection.Adap.Fill(dts)

                If dts.Tables.Count > 0 Then
                    If dts.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In dts.Tables(0).Rows
                            Dim openedtab As New SystemOpenedTab
                            For Each member In openedtab.GetType.GetProperties
                                If member.CanWrite Then
                                    If member.PropertyType.Name = "String" Or member.PropertyType.Name = "Int32" Or member.PropertyType.Name = "DateTime" Or member.PropertyType.Name = "Boolean" Then
                                        If Not IsDBNull(row(member.Name)) Then
                                            If member.PropertyType.Name = "String" Then
                                                member.SetValue(openedtab, row(member.Name).ToString, Nothing)
                                            End If
                                            If member.PropertyType.Name = "Int32" Then
                                                member.SetValue(openedtab, Integer.Parse(row(member.Name)), Nothing)
                                            End If
                                            If member.PropertyType.Name = "DateTime" Then
                                                member.SetValue(openedtab, row(member.Name), Nothing)
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            tabsopeneds.Add(openedtab)
                        Next
                    End If
                End If
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
            Return tabsopeneds
        End Function

        Public Function HasOpenedTabs(ByVal opentab As entities.SystemOpenedTab) As Boolean
            Dim has As Boolean = False
            Dim result As Integer = -1
            If (opentab.User_sess <> "" And opentab.Tab_sel <> "") Or (Not IsNothing(opentab.User_sess) And Not IsNothing(opentab.Tab_sel)) Then
                Try
                    OpenDB("DB-DDR")
                    connection.Command = New OleDb.OleDbCommand("select count(1) from System_openedTabs where User='" & opentab.User_sess & "' and Tab='" & opentab.Tab_sel & "' and Active=yes", connection.Connection)
                    result = connection.Command.ExecuteScalar()
                Catch ex As Exception
                    Throw
                Finally
                    CloseDB()
                End Try
            End If
            Return has
        End Function

        Public Sub DeleteLogOpenTab(ByVal opentab As SystemOpenedTab)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from System_openedTabs where OpenedTab_ID=" & opentab.OpenedTab_ID.ToString & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Sub DeleteLogOpenTabs(ByVal opentab As SystemOpenedTab)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from System_openedTabs where User_sess='" & opentab.User_sess & "'", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

        Public Sub DeleteLogisticTransitLog(ByVal item As LogisticTransitLog)
            Try
                OpenDB("DB-DDR")
                connection.Command = New OleDb.OleDbCommand("Delete from DDR_LogisticTransitLog where LTID=" & item.LTID & "", connection.Connection)
                connection.Command.ExecuteNonQuery()
            Catch ex As Exception
                Throw
            Finally
                CloseDB()
            End Try
        End Sub

    End Class
End Namespace
