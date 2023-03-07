Imports TVCS.MySharedFunctions
Imports TVCS.Crd_Ctrl
Imports TVCS.MySharedFunctionsWzConn
Module LCL_only
    Private Function DefineBG_TinChap(ByVal pCust As Integer, ByVal BG_CR As String, ByVal ppConn As SqlClient.SqlConnection) As Decimal
        Dim KQ As Decimal
        Dim cmd As SqlClient.SqlCommand = ppConn.CreateCommand
        strSQL = "select isnull(sum(BGAmount),0) as TTL from bg where status='OK' and custid=" & pCust & _
            " and getdate() between BGValidFrm and BGExpireDate "
        If BG_CR = "BG" Then
            strSQL = strSQL & " and Bank not in ('CRD','CSH')"
        Else
            strSQL = strSQL & " and Bank in ('CRD','CSH')"
        End If
        cmd.CommandText = strSQL
        Try
            KQ = cmd.ExecuteScalar
        Catch ex As Exception
            KQ = -2
        End Try
        Return KQ
    End Function
    Public Sub CheckOverDue_RefreshBGValidity(ByVal pCheckRunAlready As Int16, ByVal intConxIndex As Integer)
        Try
            If Not conn(intConxIndex).State = ConnectionState.Open Then conn(intConxIndex).Open()
            Dim dTable As DataTable, tmpCustID As Integer
            Dim cmd As SqlClient.SqlCommand = conn(intConxIndex).CreateCommand
            Dim tmpDOI As Date = DateAdd(DateInterval.Day, -2, Now.Date)
            strSQL = "select CustID from cc_setting where status='OK' and crCoef >0 " &
                " and custID in (select custID from cust_Detail where cat='Channel' and status='OK' and val in('TA','TO'))" &
                " order by custID "
            dTable = GetDataTable(strSQL, conn(intConxIndex))
            For i As Int16 = 0 To dTable.Rows.Count - 1
                tmpCustID = dTable.Rows(i)("CustID")

                CheckOverDue(tmpCustID, conn(intConxIndex))
                refreshValidity_BG_Credit_12(tmpCustID, pCheckRunAlready, conn(intConxIndex), intConxIndex)
            Next
            Insert_Update_WebTable(conn_Web, "update MISC SET details='" & Environment.MachineName.ToString & "', LstRun='" &
                    Now & "' where CAT='CHKOVERDUE" & MyCity(intConxIndex) & "'")
            cmd.CommandText = "insert MISC (CAT,VAL) select distinct 'BSPSTOCK', substring(tkno,5,4) " &
                "from " & arrTvcsDb(intConxIndex) & ".dbo.tkt_1A where prg in ('TTQ', 'TTP') and len(TKNO )>8 and doi >'" & tmpDOI & "' and " &
                "substring(tkno,5,4) not in (select val from misc where cat='BSPSTOCK')"
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub refreshValidity_BG_Credit_12(ByVal pCustID As Integer, ByVal pCheckRunAlready As Int16, ByVal myConn As SqlClient.SqlConnection, ByVal pICT As Byte)
        Dim cmd As SqlClient.SqlCommand = myConn.CreateCommand
        Dim cmd_web As SqlClient.SqlCommand = conn_Web.CreateCommand
        Dim BG As Decimal, CR As Decimal, XpireDate As Date
        Dim tmpRecID As Integer
        BG = DefineBG_TinChap(pCustID, "BG", myConn)
        CR = DefineBG_TinChap(pCustID, "CR", myConn)
        If BG = -2 Or CR = -2 Then Exit Sub

        Try
            If pCheckRunAlready = 1 Then
                strSQL = "select top 1 BGExpireDate from BG where status='OK' and custid=" & pCustID
                strSQL = strSQL & " and bank not in ('***','CSH','CRD') and BGExpireDate < DateAdd(dd, 14, getdate()) and BGExpireDate > getdate() "
                strSQL = strSQL & " and renewinproc=0 order by BGExpireDate"
                cmd.CommandText = strSQL
                XpireDate = cmd.ExecuteScalar
                If XpireDate > Now.Date Then
                    Dim strCustShortName As String = String.Empty
                    strCustShortName = cmd.ExecuteScalar("select top 1 CustShortName from CustomerList where RecId=" & pCustID)

                    cmd_web.CommandText = "select top 1 recid from emaillog where subj='BG Expire Advice'" &
                        " and month(fstupdate)=month(getdate()) and year(fstupdate)=year(getdate())" &
                        " and (custid in(-129," & pCustID & ") or  (msg like '" & pCustID.ToString & "%') OR (msg like '" & strCustShortName & "%'))"
                    tmpRecID = cmd_web.ExecuteScalar
                    '^_^20221028 mark by 7643 -b-
                    'If tmpRecID = 0 Then
                    '    strSQL = "Insert FT.dbo.Emaillog (CustID, Subj, Msg, Frm, city, Dept) values (-129"
                    '    strSQL = strSQL & ",'BG Expire Advice'," & strCustShortName & "' BG Will Expire On " & Format(XpireDate, "dd-MMM-yy")
                    '    strSQL = strSQL & ". Plz Renew it ASAP. Tks','SYS','" & MyCity(pICT) & "','ACC')"
                    '    Insert_Update_WebTable(conn_Web, strSQL)
                    'End If
                    '^_^20221028 mark by 7643 -e-
                    '^_^20221028 modi by 7643 -b-
                    If tmpRecID = 0 Then
                        strSQL = "Insert FT.dbo.Emaillog (CustID, Subj, Msg, Frm, city, Dept) values (-129"
                        strSQL = strSQL & ",'BG Expire Advice'," & strCustShortName & "' BG Will Expire On " & Format(XpireDate, "dd-MMM-yyyy")
                        strSQL = strSQL & ". Plz Renew it ASAP. Tks','SYS','" & MyCity(pICT) & "','ACC')"
                        Insert_Update_WebTable(conn_Web, strSQL)
                    End If
                    '^_^20221028 modi by 7643 -e-
                End If
            End If
            cmd.CommandText = "select top 1 recID from CC_BLC where custid=" & pCustID & " order by recid desc"
            tmpRecID = cmd.ExecuteScalar
            cmd.CommandText = "update CC_BLC set BG=" & BG & " where BG <>" & BG & " and Recid=" & tmpRecID
            cmd.ExecuteNonQuery()
            cmd.CommandText = "update CC_BLC set CR=" & CR & " where CR <>" & CR & " and Recid=" & tmpRecID
            cmd.ExecuteNonQuery()
        Catch ex As Exception

        End Try
    End Sub
End Module

