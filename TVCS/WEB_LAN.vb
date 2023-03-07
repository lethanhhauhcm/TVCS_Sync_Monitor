Imports TVCS.MySharedFunctions
Imports TVCS.MySharedFunctionsWzConn
Imports System.Text
Imports System.Runtime.InteropServices
Module WEB_LAN
    Dim dTable As DataTable
    Public Sub Sync_Daily(ByVal i As Byte)
        Try
            If conn(i).State = ConnectionState.Closed Then conn(i).Open()
            If conn_Web.State = ConnectionState.Closed Then conn_Web.Open()
            If conn(i).State = ConnectionState.Closed Or conn_Web.State = ConnectionState.Closed Then Exit Sub

            DaNhapRasToWeb(0, i, -16)

            UpdateVoidRefundRQ(i)
        Catch ex As Exception
        End Try
    End Sub
    Public Sub Sync_7(ByVal intConxIndex As Integer)
        Dim BackDate8 As Date
        Dim strBackDate As String
        Try
            If Not conn(intConxIndex).State = ConnectionState.Open Then conn(intConxIndex).Open()
            If Not conn_Web.State = ConnectionState.Open Then conn_Web.Open()
            If conn(intConxIndex).State = ConnectionState.Closed Or conn_Web.State = ConnectionState.Closed Then Exit Sub
            Dim cmd As SqlClient.SqlCommand = conn(intConxIndex).CreateCommand
            Dim cmd_Web As SqlClient.SqlCommand = conn_Web.CreateCommand

            Insert_Update_WebTable(conn_Web, RPT_Sync_Running(intConxIndex))

            BackDate8 = ScalarToDate("select dateadd(d,-8,getdate())", conn_Web)
            'strBackDate = Format(BackDate8, "dd-MMM-yy")  '^_^20221028 mark by 7643
            strBackDate = Format(BackDate8, "dd-MMM-yyyy")  '^_^20221028 modi by 7643
            strSQL = "update " & arrTvcsDb(intConxIndex) & ".dbo.TKT_1A set Status='RE' where status='OK' and recid in" _
                & " (select RecID from  " & arrTvcsDb(intConxIndex) & ".dbo.TKT_1A a " _
                & " Left Join (select t.tkno,t.srv, r.counter from tkt t " _
                & " left join Rcp r on r.recid=t.rcpid and r.srv=t.srv" _
                & " where t.Status <> 'XX' and r.Status <> 'XX' and t.doi > '" & strBackDate & "' ) b " _
                & " on a.tkno=b.tkno and a.srv=b.srv and a.counter=b.counter" _
                & " where a.status='OK' and a.doi>'" & strBackDate & "'  and b.tkno is not null)"

            'strSQL = strSQL & " tvcs.dbo.funcTKT_1A_updated2RAS_wzparam ('" & Format(BackDate8, "dd-MMM-yy") & "')) "
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()

            strSQL = "update " & arrTvcsDb(intConxIndex) & ".dbo.TKT_1A set Status='RE' where status='OK' and SRV='S' and "
            '^_^20221028 mark by 7643 -b-
            'strSQL = strSQL & " recid in (select RecID from " & arrTvcsDb(intConxIndex) & ".dbo.funcTKT_1ANotUpdated2RAS_wzparam('" & Format(BackDate8, "dd-MMM-yy") & "')) and "
            'strSQL = strSQL & " TKNO in (select TKNO from " & arrTvcsDb(intConxIndex) & ".dbo.funcTKT_1A_Updated2RAS_wzparam ('" & Format(BackDate8, "dd-MMM-yy") & "') where srv='V')"
            '^_^20221028 mark by 7643 -e-
            '^_^20221028 modi by 7643 -b-
            strSQL = strSQL & " recid in (select RecID from " & arrTvcsDb(intConxIndex) & ".dbo.funcTKT_1ANotUpdated2RAS_wzparam('" & Format(BackDate8, "dd-MMM-yyyy") & "')) and "
            strSQL = strSQL & " TKNO in (select TKNO from " & arrTvcsDb(intConxIndex) & ".dbo.funcTKT_1A_Updated2RAS_wzparam ('" & Format(BackDate8, "dd-MMM-yyyy") & "') where srv='V')"
            '^_^20221028 modi by 7643 -e-
            cmd.CommandText = strSQL
            cmd.ExecuteNonQuery()

            dTable = GetDataTable("select LocID, PSW, QNumber, Status from fox_iAmchick where left(status,1) <>'X' and city='" & MyCity(intConxIndex) & "'", conn_Web)
            For r As Int16 = 0 To dTable.Rows.Count - 1
                cmd.CommandText = "update " & arrTvcsDb(intConxIndex) & ".dbo.fox_iAmChick set PSW='" & dTable.Rows(r)("PSW") & "', QNumber='" &
                    dTable.Rows(r)("QNumber") & "', Status='" & dTable.Rows(r)("Status") & "' where recid=" & dTable.Rows(r)("LocID") &
                    " and left(status,1) <>'X'"
                cmd.ExecuteNonQuery()
            Next
            If strSQL.Length > 2 Then Insert_Update_WebTable(conn_Web, strSQL.Substring(1))

            Call Add_Delete_Chick(intConxIndex)

            Insert_Update_WebTable(conn_Web, RPT_Sync_Running(intConxIndex))

        Catch ex As Exception

        End Try
    End Sub

    Public Sub Sync_5(ByVal i As Byte)
        If conn(i).State = ConnectionState.Closed Then conn(i).Open()
        If Not conn_Web.State = ConnectionState.Open Then conn_Web.Open()
        If conn(i).State = ConnectionState.Closed Or conn_Web.State = ConnectionState.Closed Then Exit Sub

        DaNhapRasToWeb(0, i, -4)
    End Sub

    Public Sub Sync_2(ByVal pInterval As Int16, ByVal ct As Byte)
        Dim tmpLocID As Integer, tblCust As DataTable, LstSync As Integer
        Dim FrmDate As Date = DateAdd(DateInterval.Hour, pInterval, Now)
        Try

            If conn(ct).State <> ConnectionState.Open Then conn(ct).Open()
            If conn_Web.State <> ConnectionState.Open Then conn_Web.Open()
            If conn(ct).State <> ConnectionState.Open Or conn_Web.State <> ConnectionState.Open Then Exit Sub
            Dim cmd As SqlClient.SqlCommand = conn(ct).CreateCommand
            Dim cmd_web As SqlClient.SqlCommand = conn_Web.CreateCommand

            DaNhapRasToWeb(0, ct, -2)
            '^_^20221028 mark by 7643 -b-
            'tblCust = GetDataTable("select distinct CustID from CC_BLC where FstUpdate >'" &
            '                       Format(FrmDate, "dd-MMM-yy HH:mm") & "'", conn(ct))
            '^_^20221028 mark by 7643 -e-
            '^_^20221028 modi by 7643 -b-
            tblCust = GetDataTable("select distinct CustID from CC_BLC where FstUpdate >'" &
                                   Format(FrmDate, "dd-MMM-yyyy HH:mm") & "'", conn(ct))
            '^_^20221028 modi by 7643 -e-
            strSQL = ""
            For r As Int16 = 0 To tblCust.Rows.Count - 1
                dTable = GetDataTable("select top 1 * from CC_BLC where custID=" &
                                      tblCust.Rows(r)("CustID") & " order by recID desc", conn(ct))
                cmd_web.CommandText = "select RecID from cc_blc where custid=" & dTable.Rows(0)("CustID") &
                    " and Locid=" & dTable.Rows(0)("recID")
                tmpLocID = cmd_web.ExecuteScalar
                If tmpLocID = 0 Then
                    strSQL = strSQL & "; insert cc_BLC (FstUser, CustID, BG, CR, PPD_Depo, CRCoef, PSP_UnUsed, PPD_Used, PSP_UnPaid,"
                    strSQL = strSQL & " PPCoef, PSP_UnInv, FT_Sale, SabreSale, city, CustShortName, LOCID,RMK ) values ('"
                    strSQL = strSQL & dTable.Rows(0)("FstUser") & "',"
                    strSQL = strSQL & dTable.Rows(0)("CustID") & ","
                    strSQL = strSQL & dTable.Rows(0)("BG") & ","
                    strSQL = strSQL & dTable.Rows(0)("CR") & ","
                    strSQL = strSQL & dTable.Rows(0)("PPD_Depo") & ","
                    strSQL = strSQL & dTable.Rows(0)("CRCoef") & ","
                    strSQL = strSQL & dTable.Rows(0)("PSP_Unused") & ","
                    strSQL = strSQL & dTable.Rows(0)("PPD_Used") & ","
                    strSQL = strSQL & dTable.Rows(0)("PSP_Unpaid") & ","
                    strSQL = strSQL & dTable.Rows(0)("PPCoef") & ","
                    strSQL = strSQL & dTable.Rows(0)("PSP_Uninv") & ","
                    strSQL = strSQL & dTable.Rows(0)("FT_Sale") & ","
                    strSQL = strSQL & dTable.Rows(0)("SabreSale") & ",'"
                    strSQL = strSQL & MyCity(ct) & "','"
                    strSQL = strSQL & dTable.Rows(0)("CustShortName") & "',"
                    strSQL = strSQL & dTable.Rows(0)("RecID") & ",'"
                    strSQL = strSQL & dTable.Rows(0)("RMK") & "')"

                End If
            Next
            If strSQL.Length > 2 Then
                Insert_Update_WebTable(conn_Web, strSQL.Substring(1))
            End If


            'Lay tkt tu web ve local
            cmd.CommandText = "select top 1 WebID from " & arrTvcsDb(ct) & ".dbo.tkt_1a order by WebID desc"
            LstSync = cmd.ExecuteScalar
            cmd.CommandText = "exec " & arrTvcsDb(ct) & ".dbo.Web_TKT1a_toLocal " & LstSync & "," & MyCity(ct)
            cmd.ExecuteNonQuery()
            '=====

            Insert_Update_WebTable(conn_Web, RPT_Sync_Running(ct))
            '^_^20221028 mark by 7643 -b-
            'cmd.CommandText = "insert MISC (CAT,VAL) select distinct 'BSPSTOCK', substring(tkno,5,4) " &
            '    "from " & arrTvcsDb(ct) & ".dbo.tkt_1A where prg in ('TTQ', 'TTP') and len(TKNO )>8 and doi >'" & Format(FrmDate, "dd-MMM-yy") & "' and " &
            '    "substring(tkno,5,4) not in (select val from misc where cat='BSPSTOCK')"
            '^_^20221028 mark by 7643 -e-
            '^_^20221028 modi by 7643 -b-
            cmd.CommandText = "insert MISC (CAT,VAL) select distinct 'BSPSTOCK', substring(tkno,5,4) " &
                "from " & arrTvcsDb(ct) & ".dbo.tkt_1A where prg in ('TTQ', 'TTP') and len(TKNO )>8 and doi >'" & Format(FrmDate, "dd-MMM-yyyy") & "' and " &
                "substring(tkno,5,4) not in (select val from misc where cat='BSPSTOCK')"
            '^_^20221028 modi by 7643 -e-
            cmd.ExecuteNonQuery()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub DaNhapRasToWeb(ByVal pCustID As Integer, ByVal pct As Byte, ByVal pBackDate As Int16)
        Dim conn1 As New SqlClient.SqlConnection, BackDate16 As Date = DateAdd(DateInterval.Day, pBackDate, Now.Date)
        Dim dTable As DataTable
        Try
            conn1.ConnectionString = connStrRAS(pct)
            conn1.Open()
            '^_^20221028 mark by 7643 -b-
            'dTable = GetDataTable("select WebID from " & arrTvcsDb(pct) & ".dbo.funcTKT_1A_Updated2RAS_wzparam('" &
            '                      Format(BackDate16, "dd-MMM-yy") & "') where WebID<>0 ", conn1)
            '^_^20221028 mark by 7643 -e-
            '^_^20221028 modi by 7643 -b-
            dTable = GetDataTable("select WebID from " & arrTvcsDb(pct) & ".dbo.funcTKT_1A_Updated2RAS_wzparam('" &
                                  Format(BackDate16, "dd-MMM-yyyy") & "') where WebID<>0 ", conn1)
            '^_^20221028 modi by 7643 -e-
            For i As Int16 = 0 To dTable.Rows.Count - 1
                Insert_Update_WebTable(conn_Web, "update TKT_1A set Status='RE' , tracking =tracking + '|RE ' + convert(varchar(24),GETDATE(), 120) where status='OK' and RecID = " & dTable.Rows(i)("WebID"))
            Next
        Catch ex As Exception
        End Try
        conn1.Close()
        conn1.Dispose()
    End Sub
    Private Sub UpdateVoidRefundRQ(ByVal pCT As Byte)
        Dim tmp1AWebID As Integer
        Try
            Dim cmd_w As SqlClient.SqlCommand = conn_Web.CreateCommand
            Dim cmd As SqlClient.SqlCommand = conn(pCT).CreateCommand
            cmd_w.CommandText = "select top 1 RecID from tkt_1a where srv+Status='XOK'"
            tmp1AWebID = cmd_w.ExecuteScalar
            If tmp1AWebID > 0 Then
                cmd.CommandText = "update tkt_1A set status='XX' where webID=" & tmp1AWebID
                cmd.ExecuteNonQuery()
                Insert_Update_WebTable(conn_Web, "update TKT_1A set Status='RE' where RecID=" & tmp1AWebID)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Sub Add_Delete_Chick(ByVal bytConxId As Byte)
        Try
            dTable = GetDataTable("select RecID, LstUpdate from " & arrTvcsDb(bytConxId) & ".dbo.fox_iAmChick where status='XX' and lstupdate >'" &
                DateAdd(DateInterval.Day, -8, Now.Date) & "'", conn(bytConxId))
            strSQL = ""
            For r As Int16 = 0 To dTable.Rows.Count - 1
                'strSQL = strSQL & "; Update fox_iamchick set status='XX', LstUpdate='" & Format(dTable.Rows(r)("LstUpdate"), "dd-MMM-yy HH:mm")  '^_^20221028 mark by 7643
                strSQL = strSQL & "; Update fox_iamchick set status='XX', LstUpdate='" & Format(dTable.Rows(r)("LstUpdate"), "dd-MMM-yyyy HH:mm")  '^_^20221028 modi by 7643
                strSQL = strSQL & "' where status<>'XX' and city='" & MyCity(bytConxId) & "' and  LocID =" & dTable.Rows(r)("RecID")
            Next
            If strSQL.Length > 1 Then Insert_Update_WebTable(conn_Web, strSQL.Substring(1))

            dTable = GetDataTable("select LOCID, RecID from ft.dbo.fox_iAmChick where status<>'XX'", conn_Web)
            Dim ChickList As New Hashtable
            For r As Int16 = 0 To dTable.Rows.Count - 1
                ChickList.Add(dTable.Rows(r)("LocID"), dTable.Rows(r)("recID"))
            Next
            dTable = GetDataTable("select * from " & arrTvcsDb(bytConxId) & ".dbo.fox_iAmChick where status <>'XX' and fstupdate >dateadd(d,-4,getdate())", conn(bytConxId))
            strSQL = ""
            For r As Int16 = 0 To dTable.Rows.Count - 1
                If Not ChickList.ContainsKey(dTable.Rows(r)("recID")) Then
                    strSQL = strSQL & "; insert fox_iamchick (CustID, SI, FstUpdate, LstUpdate, FstUser, Status, PSW, Mobile, LocID, QNumber, City)"
                    strSQL = strSQL & " values (" & dTable.Rows(r)("CustID") & ",'" & dTable.Rows(r)("SI") & "','"
                    'strSQL = strSQL & Format(dTable.Rows(r)("FstUpdate"), "dd-MMM-yy HH:mm") & "','" & Format(dTable.Rows(r)("LstUpdate"), "dd-MMM-yy HH:mm")  '^_^20221028 mark by 7643
                    strSQL = strSQL & Format(dTable.Rows(r)("FstUpdate"), "dd-MMM-yyyy HH:mm") & "','" & Format(dTable.Rows(r)("LstUpdate"), "dd-MMM-yyyy HH:mm")  '^_^20221028 modi by 7643
                    strSQL = strSQL & "','" & dTable.Rows(r)("FstUser") & "','" & dTable.Rows(r)("Status") & "','"
                    strSQL = strSQL & dTable.Rows(r)("PSW") & "','" & dTable.Rows(r)("Mobile") & "'," & dTable.Rows(r)("recID")
                    strSQL = strSQL & ",'" & dTable.Rows(r)("QNumber") & "','" & MyCity(bytConxId) & "')"
                End If
            Next
            If strSQL.Length > 2 Then Insert_Update_WebTable(conn_Web, strSQL.Substring(1))
        Catch ex As Exception
        End Try
    End Sub

End Module
