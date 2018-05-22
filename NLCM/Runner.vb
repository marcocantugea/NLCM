Public Class Runner

    Private WithEvents _timer As New Timer
    Dim _MainForm As MainForm
    Dim _ADO As New ADO
    Dim _testdays As Integer = -1
    Dim _currentdaysactive As Integer = 0

    Private Sub Runner_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ''time to run the form "MainForm" to 24 hrs
        ''_timer.Interval = 86400000
        '_timer.Interval = 1000
        ''_timer.Interval = 300000
        '_MainForm = New MainForm
        '_MainForm.Show()

        '_timer.Start()
    End Sub

    Private Sub TimerTicker(sender As Object, e As EventArgs) Handles _timer.Tick
        Dim Hour As String = Date.Now.ToString("HHmmss")

        Label1.Text = Hour
        If Hour.Equals("235859") Then
            _MainForm.Close()
            _MainForm.Dispose()
            _MainForm = Nothing
        End If
        If Hour.Equals("000000") Then
            _MainForm = New MainForm
            _MainForm.Show()
            _currentdaysactive += 1
        End If

        If _testdays > 0 Then
            If _currentdaysactive = _testdays Then
                _timer.Stop()
                _MainForm.Close()
                _MainForm.Dispose()
                _MainForm = Nothing
                Application.Exit()
                End
            End If
        End If

    End Sub

    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)
        If Not Me.IsHandleCreated Then
            Me.CreateHandle()
            value = False
        End If

        'time to run the form "MainForm" to make close at 23:58:29 and opens at 00:00:00 of the next day
        _timer.Interval = 1000
        _MainForm = New MainForm
        _MainForm.Show()
        _timer.Start()

        'get test days from database
        ' if test days is -1 will run the program normal
        ' if is limited will run only the days on the parameter
        Dim testdaysparam As New ConfigObj
        testdaysparam.parameter = "-7"
        testdaysparam.param_val = "-7"
        testdaysparam.active = True
        testdaysparam.IDConfig = -7
        Try
            _ADO.GetParameter(testdaysparam, "testdays")
            _testdays = Integer.Parse(testdaysparam.param_val)
        Catch ex As Exception
            _testdays = -1
        End Try


        MyBase.SetVisibleCore(value)
    End Sub
End Class