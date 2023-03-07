Imports TVCS.MySharedFunctions
Imports TVCS.MySharedFunctionsWzConn
Public Class TVCS
    Private iCT As Byte
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        CloseAllConn()
    End Sub
        Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Byte, tmpLstRun As Date, tmpString As String = ""

        pstrPrg = "TVS"
        MyCity(0) = "SGN"
        MyCity(1) = "HAN"
        arrTvcsDb(0) = "tvcs"
        arrTvcsDb(1) = "tvcshan"
        arrRasDb(0) = "ras12"
        arrRasDb(1) = "ras12han"
        connStrRAS(0) = "server=118.69.81.103;uid=user_ras;pwd=VietHealthy@170172#;database=" & arrRasDb(0)
        connStrRAS(1) = "server=118.69.81.103;uid=user_rashan;pwd=VietHealthy@170172#;database=" & arrRasDb(1)
        RPT_Sync_Running(0) = "update fox_MISC set Dat1=Getdate() where cat='TTGT' and VAL='ISYNCSGN' "
        RPT_Sync_Running(1) = "update fox_MISC set Dat1=Getdate() where cat='TTGT' and VAL='ISYNCHAN' "

        '"select top 1 LstRun from Misc where cat='" & pCat & "'"
        conn(0) = New SqlClient.SqlConnection
        conn(0).ConnectionString = connStrRAS(0)
        conn(1) = New SqlClient.SqlConnection
        conn(1).ConnectionString = connStrRAS(1)

        conn_Web.ConnectionString = "server=118.69.81.103;uid=user_ft;pwd=VietHealthy@170172#;database=FT"
        Me.Show()
        On Error Resume Next
        conn_Web.Open()
         On Error GoTo QuitDueTimeOut
        For iCT = 0 To 1
            conn(iCT).Open()
        Next
        'pblnTestInv = True
        'AutoE_Invoice("TS0622002882", True)
        On Error GoTo 0
        Me.Timer1.Enabled = True
        Me.TrayIcon.Visible = True
        Exit Sub
QuitDueTimeOut:
        Me.Timer1.Enabled = True
        On Error GoTo 0
        Me.TrayIcon.Visible = True
    End Sub
    Private Sub CloseAllConn()
        On Error Resume Next
        conn_Web.Close()
        conn_Web.Dispose()
        For iCT = 0 To 1
            conn(iCT).Close()
            conn(iCT).Dispose()
        Next
        End
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Timer1.Enabled = False
        On Error GoTo QuitDueTimeOut
        If Not conn_Web.State = ConnectionState.Open Then conn_Web.Open()
        'On Error Resume Next
        If Not conn(0).State = ConnectionState.Open Then conn(0).Open()
        If Not conn(1).State = ConnectionState.Open Then conn(1).Open()
        'If Me.Visible Then Me.Hide()
        If Not conn(0).State = ConnectionState.Open And Not conn(1).State = ConnectionState.Open Then
            GoTo QuitDueTimeOut
        End If

        intCounter = intCounter + 1
        Reset() 'Close all file opened using fileOpen

        Me.TrayIcon.Text = intCounter.ToString & "/" & Now

        'phong truong hop gio PC bi sai
        Dim intRetryCount As Int16
CheckPcTime:
        If SysDateIsWrong(conn(0)) Then
            intRetryCount = intRetryCount + 1
            If intRetryCount = 3 Then
                conn(0).Close()
                conn(1).Close()
                Application.Exit()
            Else
                Threading.Thread.Sleep(15000)
                GoTo CheckPcTime
            End If
        Else
            intRetryCount = 0
        End If

        For iCT = 0 To 1
            Me.TxtFeedBack.Text = Now & vbTab & "Start sync_2 " & iCT.ToString & vbCrLf & Me.TxtFeedBack.Text
            Call Sync_2(-2, iCT)

            AutoGetCapturedTKT2RAS_Master()
        Next

        If intCounter Mod 5 = 0 Then
            For iCT = 0 To 1
                Me.TxtFeedBack.Text = Now & vbTab & "Start sync_5 " & iCT.ToString & vbCrLf & Me.TxtFeedBack.Text
                Call Sync_5(iCT)
            Next

        End If
        If intCounter Mod 7 = 0 Then
            For iCT = 0 To 1
                Me.TxtFeedBack.Text = Now & vbTab & "Start sync_7 " & iCT.ToString & vbCrLf & Me.TxtFeedBack.Text
                Call Sync_7(iCT)
                Call Sync_2(-2, iCT)
            Next
        End If
        If intCounter = 53 Then
            intCounter = 1
            For iCT = 0 To 1
                Me.TxtFeedBack.Text = Now & vbTab & "Start sync_daily " & iCT.ToString & vbCrLf & Me.TxtFeedBack.Text
                Call Sync_Daily(iCT)
                Me.TxtFeedBack.Text = Now & vbTab & "Start Refresh BG_OVerDue " & iCT.ToString & vbCrLf & Me.TxtFeedBack.Text
                Call CheckOverDue_RefreshBGValidity(1, iCT)
                Call Sync_2(-2, iCT)
            Next
            Insert_Update_WebTable(conn_Web, "update tkt_1a set status='RE' where status='OK' and SRV='V' and creditamt=0 and doi <dateadd (dd,-1,getdate())")
        End If
QuitDueTimeOut:
        Me.Timer1.Enabled = True
        On Error GoTo 0
    End Sub
    Private Sub TrayIcon_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TrayIcon.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then Me.Hide()
    End Sub
    Private Sub Form1_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        Me.TrayIcon.Visible = Not Me.Visible
    End Sub


End Class
