Imports System.IO

Module mdlSubAndFunction
    Public conn(2) As SqlClient.SqlConnection
    Public connStrRAS(2) As String
    Public strSQL As String
    Public intCounter As Int16
    Public MyCity(2) As String
    Public conn_Web As New SqlClient.SqlConnection
    Public KQ_DirectMail As String
    Public RPT_Sync_Running(2) As String
    Public arrTvcsDb(2) As String
    Public arrRasDb(2) As String

    Public Function GenRCPNo_ID(ByVal pTRXCode As String, ByVal parPOS As String, PTKT1AID As Integer _
                                , ByRef conn As SqlClient.SqlConnection) As String()
        Dim NewRCPno As String, i As Int32
        Dim ThangNam As String = Format(Now.Date, "MMyy")
        Dim cmd As SqlClient.SqlCommand = conn.CreateCommand
        ThangNam = ThangNam & parPOS
        cmd.CommandText = "select top 1 RCPNO from RCP where left(RCPno,7)='" & pTRXCode & ThangNam & "' order by RCPNO desc"
        NewRCPno = cmd.ExecuteScalar
        If NewRCPno = "" Then
            NewRCPno = pTRXCode & ThangNam & "00001"
        Else
            i = CInt(NewRCPno.Substring(7)) + 1
            NewRCPno = pTRXCode & ThangNam & Format(i, "00000")
        End If
        Dim RCPID As Integer = Insert_RCP(NewRCPno, pTRXCode, PTKT1AID, conn)
        Dim KQ(2) As String
        KQ(0) = RCPID
        KQ(1) = NewRCPno
        Return KQ
    End Function

    Public Function Insert_RCP(ByVal pRCPNO As String, ByVal pAL As String, ptkt1AID As Integer _
                               , ByRef conn As SqlClient.SqlConnection) As Integer
        Dim cmd As SqlClient.SqlCommand = conn.CreateCommand

        Try
            cmd.CommandText = "insert RCP (RCPNO, AL, SBU, rmk, FstUser) values ('" & pRCPNO & "','" & pAL & "'," &
                 "'TVS'," & ptkt1AID & ",'AUT') ; SELECT SCOPE_IDENTITY() AS [RecID]"
            Return cmd.ExecuteScalar
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Public Function InvalidTourCode(ByVal pTC As String, ByVal pCustID As Integer, pSRV As String, pTKNO As String, pIsNew As Boolean, pDOI As Date, Optional pExDoc As String = "") As Boolean
        Dim RecNo As Integer, RCPID As Integer, S_TCode As String, vWhat As String = IIf(pSRV = "R", pTKNO, pExDoc)
        Dim strDK As String, cmd As SqlClient.SqlCommand = conn(0).CreateCommand
        strDK = " custid=" & pCustID & " and TCode='" & pTC & "' and BillingBy in ('Event','Bundle') and status not in ('XX','RR')"
        If pSRV = "R" Or pExDoc <> "" Then
            cmd.CommandText = "select RCPID from tkt where TKNO='" & vWhat & "' and srv='S' and status<>'XX'"
            RCPID = cmd.ExecuteScalar
            cmd.CommandText = "select top 1 Document from fop where RCPID=" & RCPID & " and status='OK'"
            S_TCode = cmd.ExecuteScalar
            If pTC <> S_TCode Then Return True
            cmd.CommandText = "select recID from dutoan_Tour where " & strDK
            RecNo = cmd.ExecuteScalar
        ElseIf pSRV = "S" And pExDoc = "" Then
            If pIsNew Then strDK = strDK & " and edate >='" & pDOI & "'"
            cmd.CommandText = "select recID from dutoan_Tour where " & strDK
            RecNo = cmd.ExecuteScalar
        End If
        If RecNo = 0 Then
            cmd.CommandText = "select count (*) from TourInfo where TourCode='" & pTC & "'"
            RecNo = cmd.ExecuteScalar
        End If
        Return (RecNo = 0)
    End Function


    Public Function Insert_INV(ByVal SQL_Exec As String, ByVal pINVNo As String, ByVal pAL As String, ByVal pRCPID As Integer, Optional pNgayTaoDautien As Date = Nothing) As String
        Dim KQ As Integer, cmd As SqlClient.SqlCommand = conn(0).CreateCommand
        If pNgayTaoDautien = Nothing Then pNgayTaoDautien = Now
        '^_^20221028 mark by 7643 -b-
        'strSQL = "insert into INV (InvNo,AL, RCPID, city, FstUser, fstUpdate) values ('" & pINVNo & "','" & pAL &
        '    "'," & pRCPID & ",'SGN','AUT','" & Format(pNgayTaoDautien, "dd-MMM-yy HH:mm") & "')"
        '^_^20221028 mark by 7643 -e-
        '^_^20221028 modi by 7643 -b-
        strSQL = "insert into INV (InvNo,AL, RCPID, city, FstUser, fstUpdate) values ('" & pINVNo & "','" & pAL &
            "'," & pRCPID & ",'SGN','AUT','" & Format(pNgayTaoDautien, "dd-MMM-yyyy HH:mm") & "')"
        '^_^20221028 modi by 7643 -e-
        If SQL_Exec = "S" Then
            Return strSQL
        Else
            Cmd.CommandText = strSQL & "; SELECT SCOPE_IDENTITY() AS [RecID]"
            KQ = Cmd.ExecuteScalar
            Return KQ.ToString
        End If
    End Function


    Public Function ForEX_12(ByVal DOS As Date, ByVal pCurr As String, ByVal pType As String, ByVal pAL As String, Optional ByVal parQuay As String = "**") As Decimal
        Dim KQ As Decimal, surCharge As Decimal
        Dim dTable As DataTable
        dTable = GetDataTable("select * from ForEx where Currency='" & pCurr & "' and Status='OK' order by EffectDate DESC, recid desc ", conn(0))
        For i As Int16 = 0 To dTable.Rows.Count - 1
            If dTable.Rows(i)("EffectDate") <= DOS And
                (dTable.Rows(i)("ApplyROEto") = "YY" Or InStr(dTable.Rows(i)("ApplyROEto"), pAL) > 0 Or
                pAL = "YY" Or InStr(dTable.Rows(i)("ApplySCto"), pAL) > 0) Then
                KQ = dTable.Rows(i)(pType)
                If pType = "RECID" Then Exit For
                surCharge = dTable.Rows(i)("SurCharge")
                If pType = "BBR" Then surCharge = -surCharge
                If (parQuay <> "**" And InStr(dTable.Rows(i)("ApplySCto"), parQuay) > 0) Or
                    InStr(dTable.Rows(i)("ApplySCto"), pAL) > 0 Then
                    KQ = KQ + surCharge
                End If
                Exit For
            End If
        Next
        Return KQ
    End Function

    Public Function GenInvNo_QD153(ByVal pRCP As String, ByVal pKyHieu As String) As String
        Dim KQ As String = "", strDK As String
        Dim strPrefix As String = pRCP.Substring(0, 2) + pKyHieu + pRCP.Substring(4, 2) + "0" 'AL yy POS
        Dim cmd As SqlClient.SqlCommand = conn(0).CreateCommand
        strDK = " left(invno,7)='" & strPrefix & "' order by substring(invno,3,10) desc" 'Thay ngay 1JUL11 xu ly tr.hop UA/CO dung chung so HD
        cmd.CommandText = "select top 1 INVNO from inv where " & strDK
        KQ = cmd.ExecuteScalar
        If KQ <> "" Then
            KQ = strPrefix & Format(CInt(Strings.Right(KQ, 5)) + 1, "00000")
        Else
            KQ = strPrefix & "00001"
        End If
        Return KQ
    End Function
    Public Sub TaoBanGhiTKTNO_INVNO_Standard(pRCPID As Integer, pINVNO As String, pINVID As Integer)
        Dim cmd As SqlClient.SqlCommand = conn(0).CreateCommand
        cmd.CommandText = "select ROE from rcp where recID=" & pRCPID
        Dim ROE As Decimal = cmd.ExecuteScalar
        cmd.CommandText = "insert TKTNO_INVNO (INVNO, INVID, FstUser, RCPID, TKNO, F_VND, T_VND, C_VND)" &
            " select '" & pINVNO & "'," & pINVID & ", FstUser, RCPID, TKNO, " & "Fare*qty*" & ROE & ", Tax*Qty*" & ROE & ", Charge*" & ROE &
            " from TKT where status='OK' and RCPID=" & pRCPID
        cmd.ExecuteNonQuery()
    End Sub
    Public Function GetDataTable(ByVal pStrCmd As String, ByVal pConn As SqlClient.SqlConnection) As DataTable
        Dim tblResults As New DataTable
        Dim adapter As New SqlClient.SqlDataAdapter(pStrCmd, pConn)
        adapter.Fill(tblResults)
        Return tblResults
    End Function
    Public Function GetVALFromMisc(ByVal pCat As String) As String
        Dim KQ As String = "", cmd_Web As SqlClient.SqlCommand = conn_Web.CreateCommand
        On Error GoTo errHandeler
        If Not conn_Web.State = ConnectionState.Open Then conn_Web.Open()
        If Not conn_Web.State = ConnectionState.Open Then GoTo errHandeler
        cmd_Web.CommandText = "select VAL from Misc where cat='" & pCat & "'"
        KQ = cmd_Web.ExecuteScalar
errHandeler:
        Return KQ
    End Function
    Public Function ScalarToInt(ByVal pSQL As String) As Integer
        Dim KQ As Integer, cmd_Web As SqlClient.SqlCommand = conn_Web.CreateCommand
        On Error GoTo errHandeler
        If Not conn_Web.State = ConnectionState.Open Then conn_Web.Open()
        If Not conn_Web.State = ConnectionState.Open Then GoTo errHandeler
        cmd_Web.CommandText = pSQL
        KQ = cmd_Web.ExecuteScalar
errHandeler:
        Return KQ
    End Function

    Public Function ScalarToDate(ByVal pSQL As String, objSql As SqlClient.SqlConnection) As DateTime
        Dim KQ As DateTime, cmd As SqlClient.SqlCommand = objSql.CreateCommand
        On Error GoTo errHandeler
        If Not objSql.State = ConnectionState.Open Then objSql.Open()
        If Not objSql.State = ConnectionState.Open Then GoTo errHandeler
        cmd.CommandText = pSQL
        KQ = cmd.ExecuteScalar
errHandeler:
        Return KQ
    End Function
    Public Function GetLstRunFromMisc(ByVal pCat As String) As Date
        Dim KQ As Date, cmd_Web As SqlClient.SqlCommand = conn_Web.CreateCommand
        On Error GoTo errHandeler
        If Not conn_Web.State = ConnectionState.Open Then conn_Web.Open()
        If Not conn_Web.State = ConnectionState.Open Then GoTo errHandeler
        cmd_Web.CommandText = "select top 1 LstRun from Misc where cat='" & pCat & "'"
        KQ = cmd_Web.ExecuteScalar
        Return KQ
errHandeler:
    End Function
    '^_^20221017 mark by 7643 -b-
    '    Public Function CreateAopQueueAirCTS(strRcp As String) As Boolean

    '        Dim strQuerry As String
    '        Dim lstQueueRecIds As New List(Of String)
    '        Dim tblInvoice As DataTable
    '        Dim tblBill As DataTable
    '        Dim strBu As String = String.Empty
    '        Dim intResult As Integer
    '        Dim arrQueueRecIds As String()
    '        Dim strMemo As String

    '        'Tao Invoice
    '        strQuerry = "select (case when m.RecID is null then 1 else 2 end) as InvCount, r.CustId, r.RcpNo,R.Srv" _
    '                & ",r.TtlDue as InvAmt, 0 as SvcFee" _
    '                & ",r.Charge as MerchantFee" _
    '                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
    '                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
    '                & ",l.CustShortName,AOPListID" _
    '                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS Tkno" _
    '                & " from Rcp r" _
    '                & " left join CustomerList l on l.Recid=r.CustId" _
    '                & " left join Misc m on r.CustId=m.intVal and m.Cat='CustNameInGroup' and m.VAL='2 INVOICES CUS' and m.Status='OK'" _
    '                & " where r.status='OK' and r.Counter='CWT' and r.RcpNo='" & strRcp & "'"

    '        tblInvoice = GetDataTable(strQuerry, conn(0))

    '        If tblInvoice.Rows.Count = 0 Then
    '            MsgBox("Invalide Rcp:" & strRcp)
    '            Return False
    '        End If
    '        If tblInvoice.Rows(0)("AOPListID") = "" Then
    '            MsgBox("You must ask PQT to update AOPListID for the following Customers: " & tblInvoice.Rows(0)("CustShortName"))
    '            Return False
    '        End If



    '        For Each objRow As DataRow In tblInvoice.Rows
    '            strBu = "CTS-AIR"

    '            If objRow("Tkno").ToString.EndsWith("/") Then
    '                strMemo = Mid(objRow("Tkno"), 1, Len(objRow("Tkno")) - 1)
    '            Else
    '                strMemo = objRow("Tkno")
    '            End If
    '            If objRow("InvCount") = 1 Then
    '                If objRow("SRV") = "S" Then
    '                    intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), objRow("MerchantFee"), objRow("AOPListid") _
    '                                                      , objRow("TrxDate"), objRow("RefNumber"), strMemo)
    '                    If intResult = 0 Then
    '                        Return False
    '                    Else
    '                        lstQueueRecIds.Add(intResult.ToString)
    '                    End If
    '                Else
    '                    intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, "CUSTOMER RECEIVABLE (VND)", objRow("TrxDate") _
    '                                                         , objRow("RefNumber"), strMemo, objRow("InvAmt"), objRow("MerchantFee"), strMemo, strRcp, "Air")
    '                    If intResult = 0 Then
    '                        Return False
    '                    Else
    '                        lstQueueRecIds.Add(intResult.ToString)
    '                    End If
    '                End If
    '            ElseIf objRow("InvCount") = 2 Then
    '                objRow("SvcFee") = ScalarToDec("tkt", "sum(ChargeTV)", "Status<>'xx' and Rcpno='" & objRow("RcpNo") & "'")
    '                objRow("InvAmt") = objRow("InvAmt") - objRow("SvcFee")

    '                If objRow("SRV") = "S" Then
    '                    Dim decMerchantFee As Decimal
    '                    intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), objRow("MerchantFee"), objRow("AOPListid") _
    '                                                      , objRow("TrxDate"), objRow("RefNumber"), strMemo)
    '                    If intResult = 0 Then
    '                        Return False
    '                    Else
    '                        lstQueueRecIds.Add(intResult.ToString)
    '                    End If

    '                    Select Case objRow("CustShortName")
    '                        Case "PG VIETNAM", "PG INDOCHINA"
    '                            decMerchantFee = 0
    '                        Case Else
    '                            decMerchantFee = objRow("MerchantFee")
    '                    End Select
    '                    If objRow("SvcFee") > 0 Then
    '                        intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("SvcFee"), 0, objRow("AOPListid") _
    '                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo)
    '                        If intResult = 0 Then
    '                            Return False
    '                        Else
    '                            lstQueueRecIds.Add(intResult.ToString)
    '                        End If
    '                    End If

    '                Else
    '                    intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, "CUSTOMER RECEIVABLE (VND)", objRow("TrxDate") _
    '                                                         , objRow("RefNumber"), objRow("Tkno"), objRow("InvAmt"), objRow("MerchantFee"), strMemo, strRcp, "Air")
    '                    If intResult = 0 Then
    '                        Return False
    '                    Else
    '                        lstQueueRecIds.Add(intResult.ToString)
    '                    End If
    '                    intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, "CUSTOMER RECEIVABLE (VND)", objRow("TrxDate") _
    '                                                         , objRow("RefNumber"), objRow("Tkno"), objRow("SvcFee"), objRow("MerchantFee"), strMemo, strRcp, "Air")
    '                    If intResult = 0 Then
    '                        Return False
    '                    Else
    '                        lstQueueRecIds.Add(intResult.ToString)
    '                    End If
    '                End If

    '            End If
    '        Next

    '        'Tao bill
    '        strQuerry = "select r.Vendor,c.CustShortName, r.RcpNo,R.Srv" _
    '                & " ,(select top 1 DOI from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0) AS DOI" _
    '                & " ,(select sum((NetToAL+Tax)+Charge*t.Qty) from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0) AS BillAmt" _
    '                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
    '                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
    '                & " ,(select distinct t.DocType+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS DocType" _
    '                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS Tkno" _
    '                & ",v.AOPListID as VendorAopId,c.AOPListID as CustAopId " _
    '                & " from Rcp r" _
    '                & " left join CustomerList c on c.Recid=r.CustId" _
    '                & " left join Vendor v on v.Recid=r.VendorId" _
    '                & " where r.status='OK' and r.Srv<>'V' and r.Counter='CWT' and r.rcpno='" & strRcp & "'" _
    '                & " and r.RecId not in (Select RcpId from tkt where DocType='AHC')" _
    '                & " order by r.FstUpdate"

    '        tblBill = GetDataTable(strQuerry, conn(0))

    '        For Each objRow As DataRow In tblBill.Rows
    '            If objRow("Vendor") = "" Then
    '                MsgBox("You must ask CTS to update Vendor for " & strRcp)
    '                Return False
    '            ElseIf objRow("VendorAopId") = "" Then
    '                MsgBox("You must ask PQT to update AOPListID for the following Vendor:  " & objRow("Vendor"))
    '                Return False
    '            End If
    '        Next

    '        For Each objRow As DataRow In tblBill.Rows
    '            If objRow("Tkno").ToString.EndsWith("/") Then
    '                strMemo = Mid(objRow("Tkno"), 1, Len(objRow("Tkno")) - 1)
    '            Else
    '                strMemo = objRow("Tkno")
    '            End If

    '            If objRow("SRV") = "R" Then
    '                intResult = CreateAopQueueVendorCredit(objRow("VendorAopId"), objRow("CustAopId"), strBu, "VENDOR PAYABLE (VND)" _
    '                              , objRow("TrxDate"), objRow("RefNumber"), "COST", objRow("BillAmt"), 0, strMemo, "Air", strRcp)
    '                If intResult = 0 Then
    '                    Return False
    '                Else
    '                    lstQueueRecIds.Add(intResult.ToString)
    '                End If
    '            Else
    '                Dim strDueDate As String = String.Empty
    '                Select Case objRow("Vendor")
    '                    Case "VN", "BSP", "VN DEB"
    '                        If objRow("DocType").ToString.Contains("ETK") Or objRow("DocType").ToString.Contains("EMD") _
    '                            Or objRow("DocType").ToString.Contains("MCO") Then
    '                            If objRow("Vendor") = "BSP" Then
    '                                strDueDate = GetDueDate4AopBsp(objRow("DOI"))
    '                            Else
    '                                strDueDate = GetDueDate4AopNonBsp(objRow("DOI"))
    '                            End If
    '                        End If
    '                    Case "QH TK"
    '                        strDueDate = Format(objRow("DOI"), "yyyy-MM-dd")
    '                End Select

    '                intResult = CreateAopQueueBill(objRow("VendorAopId"), objRow("CustAOPid"), strBu, objRow("TrxDate") _
    '                                                         , objRow("RefNumber"), "COST", objRow("BillAmt"), 0 _
    '                                                         , strMemo, "Air", strRcp, "VND", strDueDate)
    '                If intResult = 0 Then
    '                    Return False
    '                Else
    '                    lstQueueRecIds.Add(intResult.ToString)
    '                End If
    '            End If
    '        Next
    '        ReDim arrQueueRecIds(0 To lstQueueRecIds.Count - 1)
    '        lstQueueRecIds.CopyTo(arrQueueRecIds)
    '        Return ExecuteNonQuerry("update AopQueue Set Status='OK' where LinkId in (" & Strings.Join(arrQueueRecIds, ",") & ")", conn(0))
    '    End Function
    '    Public Function CreateAopQueueAirTVS(strRcp As String, dteTrxDate As Date) As Boolean
    '        Dim decAopUsdRoe As Decimal = ScalarToDec("Forex", "TOP 1 BSR", "Status='OK' and Currency='USD' and ApplyROETo='AOP'" _
    '                                                  & " and EffectDate <='" & CreateFromDate(dteTrxDate) _
    '                                                  & "' order by EffectDate")
    '        Dim tblInvoice As DataTable
    '        Dim tblBill As DataTable
    '        Dim objRow As DataRow
    '        Dim lstQueueRecIds As New List(Of String)
    '        Dim intResult As Integer
    '        Dim strBu As String = "TVS"
    '        Dim strMemo As String
    '        Dim arrQueueRecIds As String()

    '        Dim strQuerry As String = "select r.CustId, r.RcpNo,R.Srv" _
    '                & " ,(case (select count (*) from tkt t where t.Status<>'XX' and t.RCPID=r.RecID and T.DocType in ('GRP','MCO')) when 0 then 'FIT' else 'GRP' end) AS GRP" _
    '                & ",'' as AopRecord" _
    '                & ",r.Roe*r.TtlDue as InvAmt, 0 as SvcFee" _
    '                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
    '                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
    '                & ",l.CustShortName,'' as TourCode" _
    '                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID  ORDER BY TKNO For Xml path('') ) AS Tkno,'' as DepDate" _
    '                & ",'' as Account,'' as AccountName" _
    '                & ",'' as OriDeposits" _
    '                & " ,(select distinct t.DocType+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS DocTypes,AOPListID" _
    '                & ",r.Vendor,r.VendorId" _
    '                & ", (case r.CustShortName when 'VYM' then " & decAopUsdRoe & " else 1 end)  as ROE" _
    '                & " from Rcp r" _
    '                & " left join CustomerList l on l.Recid=r.CustId" _
    '                & " where r.status='OK' and r.Counter='TVS' and RcpNo='" & strRcp _
    '                & "' order by r.FstUpdate"

    '        tblInvoice = GetDataTable(strQuerry, conn(0))

    '        If tblInvoice.Rows.Count = 0 Then
    '            Append2TextFile("Invalid RCP " & strRcp)
    '            Return False
    '        Else
    '            objRow = tblInvoice.Rows(0)
    '        End If

    '        If objRow("AOPListID") = "" Then
    '            MsgBox("Missing AOPListID for " & objRow("CustShortName"))
    '            Return False
    '        ElseIf objRow("AopRecord") = "Deposit" Then
    '            'MsgBox("You must update account for " & objRow("RcpNo"))
    '            Return False
    '        End If

    '        objRow("tkno") = Mid(objRow("tkno"), 1, Len(objRow("Tkno")) - 1)
    '        strMemo = objRow("tkno")


    '        'import Invoice
    '        Select Case objRow("GRP")
    '            Case "GRP"
    '                'Bo khong import tu dong

    '            Case "FIT"
    '                Select Case objRow("CustShortName")

    '                    Case "TVHAN"
    '                        'Bo qua ko nhap

    '                    Case "TVSGN", "GDSSGN"
    '                        strMemo = strMemo & GetColumnValuesAsString("FOP", "Document", "WHERE Status='OK' and RcpNO='" & strRcp & "' and FOP<>'EXC'", "|")
    '                        Dim decExchangeAmt As Decimal = ScalarToDec("FOP", "Amount", "Status='OK' and FOP='EXC' and RcpNo='" & strRcp & "'")

    '                        objRow("InvAmt") = objRow("InvAmt") - decExchangeAmt

    '                        If objRow("SRV") = "S" Then
    '                            intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), 0, objRow("AOPListid") _
    '                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo,, "ODBC")
    '                            If intResult = 0 Then
    '                                Return False
    '                            Else
    '                                lstQueueRecIds.Add(intResult.ToString)
    '                            End If

    '                        Else    'refund
    '                            intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, "CUSTOMER RECEIVABLE (VND)", objRow("TrxDate") _
    '                                                             , objRow("RefNumber"), objRow("Tkno"), objRow("InvAmt"), 0, strMemo, strRcp, "Air")
    '                            If intResult = 0 Then
    '                                Return False
    '                            Else
    '                                lstQueueRecIds.Add(intResult.ToString)
    '                            End If
    '                        End If
    '                    Case "VYM" 'VAYMA (TRAVIX)
    '                        strMemo = objRow("Tkno") & " (" & Format(objRow("InvAmt"), "#,##0") & "/" & Format(objRow("ROE"), "#,##0") & ")"
    '                        Dim decInvAmt As Decimal = Math.Round(objRow("InvAmt") / objRow("ROE"), 2)
    '                        If objRow("SRV") = "S" Then
    '                            intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, decInvAmt, 0, objRow("AOPListid") _
    '                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo, "USD")
    '                            If intResult = 0 Then
    '                                Return False
    '                            Else
    '                                lstQueueRecIds.Add(intResult.ToString)
    '                            End If

    '                        Else
    '                            intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, "CUSTOMER RECEIVABLE (USD)", objRow("TrxDate") _
    '                                                             , objRow("RefNumber"), objRow("Tkno"), decInvAmt, 0, strMemo, strRcp, "Air")
    '                            If intResult = 0 Then
    '                                Return False
    '                            Else
    '                                lstQueueRecIds.Add(intResult.ToString)
    '                            End If

    '                        End If

    '                    Case Else
    '                        If objRow("SRV") = "S" Then
    '                            intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), 0, objRow("AOPListid") _
    '                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo)
    '                            If intResult = 0 Then
    '                                Return False
    '                            Else
    '                                lstQueueRecIds.Add(intResult.ToString)
    '                            End If

    '                        Else    'refund
    '                            intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, "CUSTOMER RECEIVABLE (VND)", objRow("TrxDate") _
    '                                                             , objRow("RefNumber"), objRow("Tkno"), objRow("InvAmt"), 0, strMemo, strRcp, "Air")
    '                            If intResult = 0 Then
    '                                Return False
    '                            Else
    '                                lstQueueRecIds.Add(intResult.ToString)
    '                            End If
    '                        End If
    '                End Select
    '        End Select

    '        ' Import for TSP
    '        Select Case objRow("CustShortName")
    '            Case "TVHAN"
    '                Dim tblTourCodes As DataTable = GetTourCodeTableByRcp(strRcp)
    '                Dim strJournalCreditLineEntityRefFullName As String = objRow("Vendor") & " [" & objRow("VendorId") & "]"

    '                If tblTourCodes.Rows.Count > 0 Then
    '                    strMemo = objRow("Tkno")
    '                    For Each objToourCodeRow As DataRow In tblTourCodes.Rows
    '                        strMemo = strMemo & " " & objToourCodeRow("TourCode")
    '                    Next
    '                End If

    '                If objRow("SRV") = "S" Then
    '                    intResult = CreateAopQueueJournalEntry(strBu, strRcp, "Air", objRow("TrxDate"), objRow("RefNumber"), "Vietnamese Dong" _
    '                                                               , "VENDOR PAYABLE (VND)", objRow("InvAmt"), strMemo, strJournalCreditLineEntityRefFullName _
    '                                                               , "INTERNAL DEBT:AIR HAN", "TVHAN [68612]")
    '                    If intResult = 0 Then
    '                        Return False
    '                    Else
    '                        lstQueueRecIds.Add(intResult.ToString)
    '                    End If
    '                Else
    '                    intResult = CreateAopQueueJournalEntry(strBu, strRcp, "Air", objRow("TrxDate"), objRow("RefNumber"), "Vietnamese Dong" _
    '                                                               , "INTERNAL DEBT:AIR HAN", objRow("InvAmt"), strMemo, "TVHAN [68612]" _
    '                                                               , "VENDOR PAYABLE (VND)", strJournalCreditLineEntityRefFullName)
    '                    If intResult = 0 Then
    '                        Return False
    '                    Else
    '                        lstQueueRecIds.Add(intResult.ToString)
    '                    End If
    '                End If

    '            Case "TVSGN", "GDSSGN"
    '                Dim tblTourCodes As DataTable = GetTourCodeTableByRcp(strRcp)

    '                If tblTourCodes.Rows.Count > 0 Then
    '                    strMemo = objRow("Tkno")
    '                    For Each objToourCodeRow As DataRow In tblTourCodes.Rows
    '                        strMemo = strMemo & " " & objToourCodeRow("TourCode")
    '                    Next

    '                    If objRow("SRV") = "S" Then
    '                        intResult = CreateAopQueueCheck(strBu, strRcp, "Air", "CASH AIR", "AL TRANSVIET 2020 [10015]", objRow("RefNumber"), objRow("TrxDate") _
    '                                            , objRow("InvAmt"), "Vietnamese Dong", strMemo, "VENDOR PAYABLE (VND)")
    '                        If intResult = 0 Then
    '                            Return False
    '                        Else
    '                            lstQueueRecIds.Add(intResult.ToString)
    '                        End If
    '                    Else
    '                        intResult = CreateAopQueueJournalEntry(strBu, strRcp, "Air", objRow("TrxDate"), objRow("RefNumber"), "Vietnamese Dong" _
    '                                                               , "VENDOR PAYABLE (VND)", objRow("InvAmt"), strMemo, "AL TRANSVIET 2020 [10015]" _
    '                                                               , "CASH AIR", "AL TRANSVIET 2020 [10015]")
    '                        If intResult = 0 Then
    '                            Return False
    '                        Else
    '                            lstQueueRecIds.Add(intResult.ToString)
    '                        End If
    '                    End If
    '                End If
    '        End Select

    '        'IMPORT BILL
    '        If objRow("CustShortName") <> "TVHAN" Then
    '            strQuerry = "select r.Vendor,c.CustShortName, r.RcpNo,R.Srv" _
    '                & " ,(select top 1 DOI from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0) AS DOI" _
    '                & " ,((select sum((NetToAL+Tax)+Charge*t.Qty) from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0)" _
    '                & " - (select isnull(sum(Amount),0) from FOP f where f.Status='OK' and f.RCPID=r.RecID and f.FOP='EXC'))*r.Roe  AS BillAmt" _
    '                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
    '                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
    '                & " ,(select distinct t.DocType+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS DocType" _
    '                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS Tkno" _
    '                & ",'' AS TourCode,'' as DepDate,'' as TspCust,'' as TspCustAopId,'' as TspClass" _
    '                & ",v.AOPListID as VendorAopId,c.AOPListID as CustAopId,r.CustId,r.Currency as RcpCur,v.Cur as VenCur,r.TtlDue " _
    '                & " from Rcp r" _
    '                & " left join CustomerList c on c.Recid=r.CustId" _
    '                & " left join Vendor v on v.Recid=r.VendorId" _
    '                & " where r.status='OK' and r.Srv<>'V' and r.Counter='TVS' and r.RcpNo='" & strRcp & "'" _
    '                & " and r.RecId not in (Select RcpId from tkt where DocType='AHC')"

    '            tblBill = GetDataTable(strQuerry, conn(0))
    '            objRow = tblBill.Rows(0)

    '            If IsDBNull(objRow("VendorAopId")) Then
    '                MsgBox("You must ask PQT to update AOPListID for the following Vendor: " & objRow("Vendor"))
    '                Return False
    '            End If
    '            If objRow("CustAopId") = "" Then
    '                MsgBox("You must ask PQT to update AOPListID for the following Customer: " & objRow("CustShortName"))
    '                Return False
    '            End If

    '            objRow("Tkno") = Mid(objRow("tkno"), 1, Len(objRow("Tkno")) - 1)
    '            strMemo = objRow("Tkno")

    '            If IsDBNull(objRow("BillAmt")) Then
    '                GoTo BillNotRequired    've XX doi voi Hang
    '            End If
    '            Dim strBillCur As String = "VND"
    '            Dim decBillAmt As Decimal = objRow("BillAmt")
    '            If objRow("Vendor") = "AK TK" Then
    '                strBillCur = "USD"
    '                decBillAmt = objRow("TtlDue")
    '            End If


    '            If objRow("SRV") = "R" Then
    '                intResult = CreateAopQueueVendorCredit(objRow("VendorAopId"), objRow("CustAopId"), strBu, "VENDOR PAYABLE (" & strBillCur & ")" _
    '                                  , objRow("TrxDate"), objRow("RefNumber"), "COST", decBillAmt, 0, strMemo, "Air", strRcp)
    '                If intResult = 0 Then
    '                    Return False
    '                Else
    '                    lstQueueRecIds.Add(intResult.ToString)
    '                End If

    '            ElseIf objRow("BillAmt") > 0 Then       'bo qua ve zero value
    '                Dim strDueDate As String = String.Empty
    '                Select Case objRow("Vendor")
    '                    Case "BSP", "BSP AERTICKET-37314944"
    '                        If objRow("DocType").ToString.Contains("ETK") Or objRow("DocType").ToString.Contains("EMD") _
    '                                                Or objRow("DocType").ToString.Contains("MCO") Then
    '                            strDueDate = GetDueDate4AopBsp(objRow("DOI"))
    '                        End If
    '                    Case "VN", "VN DEB"
    '                        If objRow("DocType").ToString.Contains("ETK") Or objRow("DocType").ToString.Contains("EMD") _
    '                                                Or objRow("DocType").ToString.Contains("MCO") Then
    '                            strDueDate = GetDueDate4AopNonBsp(objRow("DOI"))
    '                        End If

    '                    Case "QH TK"
    '                        strDueDate = Format(objRow("DOI"), "yyyy-MM-dd")
    '                End Select

    '                intResult = CreateAopQueueBill(objRow("VendorAopId"), objRow("CustAOPid"), strBu, objRow("TrxDate") _
    '                                                             , objRow("RefNumber"), "COST", decBillAmt, 0 _
    '                                                             , strMemo, "Air", strRcp, strBillCur, strDueDate)
    '                If intResult = 0 Then
    '                    Return False
    '                Else
    '                    lstQueueRecIds.Add(intResult.ToString)
    '                End If

    '            End If
    '        End If


    'BillNotRequired:
    '        ReDim arrQueueRecIds(0 To lstQueueRecIds.Count - 1)
    '        lstQueueRecIds.CopyTo(arrQueueRecIds)
    '        Return ExecuteNonQuerry("update AopQueue Set Status='OK' where LinkId in (" & Strings.Join(arrQueueRecIds, ",") & ")", conn(0))

    '    End Function
    'Private Function CreateAopQueueCheck(strBU As String, strTrxCode As String, strProd As String _
    '                                , strAccountRefFullName As String, strPayeeEntityRefFullName As String _
    '                                , strRefNumber As String, strTrxDate As String, decAmount As Decimal _
    '                                , strCurrencyRefFullName As String, strMemo As String _
    '                                , strExpenseLineAccountRefFullName As String) As Integer
    '    Dim intQueueRecId As Integer
    '    Dim intNewRecId As Integer
    '    Dim strQuerry As String

    '    If Not strTrxDate.Contains("{d") Then
    '        strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
    '    End If

    '    strQuerry = "INSERT INTO [CheckExpenseLine] (AccountRefFullName,PayeeEntityRefFullName,RefNumber" _
    '                    & ",TxnDate,Memo,ExpenseLineAccountRefFullName,ExpenseLineAmount" _
    '                    & " ,ExpenseLineMemo,ExpenseLineCustomerRefFullName,IsToBePrinted,FQSaveToCache) VALUES ('" _
    '                    & strAccountRefFullName & "','" & strPayeeEntityRefFullName & "','" & strRefNumber & "'," _
    '                    & strTrxDate & ",'" & strMemo & "','" & strExpenseLineAccountRefFullName _
    '                    & "'," & decAmount & ",'" & strMemo & "','" & strPayeeEntityRefFullName & "',0,1)"


    '    intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, "ODBC")
    '    If intQueueRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP CheckExpenseLine cho TrxCode " & strTrxCode)
    '        Return 0
    '    ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn(0)) Then
    '        MsgBox("Không tạo được bản ghi AOP Check cho TrxCode " & strTrxCode)
    '        Return 0
    '    End If

    '    strQuerry = "INSERT INTO [Check] (AccountRefFullName,PayeeEntityRefFullName,RefNumber" _
    '                    & ",TxnDate, Memo,IsToBePrinted) VALUES ('" & strAccountRefFullName & "','" & strPayeeEntityRefFullName & "','" & strRefNumber & "'," _
    '                    & strTrxDate & ",'" & strMemo & "',0)"

    '    intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, "ODBC")
    '    If intNewRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP Check cho TrxCode " & strTrxCode)
    '        Return 0
    '    End If


    '    Return intQueueRecId
    'End Function
    'Private Function CreateAopQueueCreditMemo(strAOPListid As String, strBU As String, strAccountName As String, strInvDate As String _
    '                               , strRefNumber As String, strDesc As String, decInvAmt As Decimal, decMerchantFee As Decimal _
    '                               , strMemo As String, strTrxCode As String, strProd As String) As Integer
    '    Dim intQueueRecId As Integer
    '    Dim intNewRecId As Integer
    '    Dim strQuerry As String
    '    Dim strItemName As String

    '    If strProd = "Air" Then
    '        strItemName = "REVENUE"
    '    Else
    '        strItemName = strDesc
    '    End If
    '    If Not strInvDate.Contains("{d") Then
    '        strInvDate = "{d'" & Format(CDate(strInvDate), "yyyy-MM-dd") & "'}"
    '    End If
    '    strQuerry = "INSERT INTO CreditMemoLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
    '                    & ",TxnDate,RefNumber,CreditMemoLineItemRefFullName, CreditMemoLineDesc" _
    '                    & ",CreditMemoLineAmount,memo,FQSaveToCache) VALUES ('" & strAOPListid & "','" & strBU & "','" & strAccountName & "'," _
    '                    & strInvDate & ",'" & strRefNumber & "','" & strItemName & "','" & strDesc _
    '                    & "'," & decInvAmt & ",'" & strMemo & "',1)"

    '    intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, "ODBC")

    '    If intQueueRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP CreditMemo cho TrxCode " & strTrxCode)
    '        Return 0
    '    ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn(0)) Then
    '        MsgBox("Không tạo được bản ghi AOP CreditMemo cho Tcode " & strTrxCode)
    '        Return 0
    '    End If

    '    If decMerchantFee <> 0 Then
    '        strQuerry = "INSERT INTO CreditMemoLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
    '                    & ",TxnDate,RefNumber,CreditMemoLineItemRefFullName, CreditMemoLineDesc" _
    '                    & ",CreditMemoLineAmount,memo,FQSaveToCache) VALUES ('" & strAOPListid & "','" & strBU & "','" & strAccountName & "'," _
    '                    & strInvDate & ",'" & strRefNumber & "','" & strItemName & "','" & strDesc _
    '                    & "'," & 0 - decMerchantFee & ",'" & strMemo & "',1)"
    '        intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, "ODBC")
    '        If intNewRecId = 0 Then
    '            MsgBox("Không tạo được bản ghi AOP CreditMemo Phí cà thẻ cho TrxCode" & strTrxCode)
    '            Return 0
    '        End If
    '    End If

    '    strQuerry = "INSERT INTO CreditMemo (CustomerRefListID,ClassRefFullName, ARAccountRefFullName, TxnDate, RefNumber" _
    '                        & ", Memo) VALUES ('" & strAOPListid & "','" & strBU & "','" & strAccountName & "'," _
    '                        & strInvDate & ",'" & strRefNumber & "','" & strMemo & "')"

    '    intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, "ODBC")
    '    If intNewRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP CreditMemo cho TrxCode " & strTrxCode)
    '        Return 0
    '    End If
    '    Return intQueueRecId
    'End Function
    'Public Function CreateAopQueueRecord(intLinkId As Integer, strBillInvoice As String, strProd As String, strBU As String, strTrxCode As String _
    '                                     , strQuerry As String, blnGetRecId As Boolean, Optional strQuerryType As String = "ODBC") As Integer
    '    Dim cmd As SqlClient.SqlCommand = conn(0).CreateCommand

    '    cmd.CommandText = "insert Into AopQueue (LinkId,B_I, Prod,BU, TrxCode, Querry, FstUser,QuerryType) values (" & intLinkId & ",'" & strBillInvoice _
    '                    & "','" & strProd & "','" & strBU & "','" & strTrxCode & "',@Querry,'AUT','" & strQuerryType & "')"

    '    If blnGetRecId Then
    '        cmd.CommandText = cmd.CommandText & ";SELECT SCOPE_IDENTITY() AS [RecID]"
    '    End If
    '    cmd.Parameters.Add("@Querry", SqlDbType.NVarChar).Value = strQuerry
    '    Try
    '        If blnGetRecId Then
    '            Return cmd.ExecuteScalar
    '        Else
    '            Return cmd.ExecuteNonQuery
    '        End If
    '    Catch ex As Exception
    '        Append2TextFile("SQL error:" & ex.Message & vbCrLf & cmd.CommandText & vbCrLf & strQuerry)
    '        Return 0
    '    End Try
    'End Function
    'Public Function CreateAopQueueInvoice(intLinkId As Integer, strBillInvoice As String, strProd As String, strBU As String, strTrxCode As String _
    '                                     , decInvAmout As Decimal, decMerchantFee As Decimal _
    '                                     , strCustIdAop As String, strTrxDate As String, strRefNumber As String, strMemo As String _
    '                                     , Optional strCur As String = "", Optional strQuerryType As String = "ODBC") As Integer
    '    Dim intQueueRecId As Integer
    '    Dim strQuerry As String

    '    If strCur = "" Then
    '        strCur = "VND"
    '    End If
    '    If Not strTrxDate.Contains("{d") Then
    '        strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
    '    End If

    '    Select Case strProd
    '        Case "NonAir"
    '            strQuerry = "INSERT INTO InvoiceLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
    '                & ",TxnDate,RefNumber,InvoiceLineItemRefFullName, InvoiceLineDesc" _
    '                & ",InvoiceLineAmount,memo,FQSaveToCache) VALUES ('" & strCustIdAop & "','" & strBU & "','CUSTOMER RECEIVABLE (" & strCur & ")'," _
    '                & strTrxDate & ",'" & strRefNumber & "','" & strTrxCode & "','" & strTrxCode _
    '                & "'," & decInvAmout & ",'" & strMemo & "',1)"
    '        Case "Air"
    '            strQuerry = "INSERT INTO InvoiceLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
    '                & ",TxnDate,RefNumber,InvoiceLineItemRefFullName, InvoiceLineDesc" _
    '                & ",InvoiceLineAmount,memo,FQSaveToCache) VALUES ('" & strCustIdAop & "','" & strBU & "','CUSTOMER RECEIVABLE (" & strCur & ")'," _
    '                & strTrxDate & ",'" & strRefNumber & "','REVENUE','" & strMemo _
    '                & "'," & decInvAmout & ",'" & strMemo & "',1)"

    '    End Select

    '    intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, strQuerryType)

    '    If intQueueRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP Invoice cho TrxCode " & strTrxCode)
    '        Return 0
    '    ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn(0)) Then
    '        MsgBox("Không tạo được bản ghi AOP Invoice cho Tcode " & strTrxCode)
    '        Return 0
    '    End If

    '    If decMerchantFee <> 0 Then
    '        strQuerry = "INSERT INTO InvoiceLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
    '            & ",TxnDate,RefNumber,InvoiceLineItemRefFullName, InvoiceLineDesc" _
    '            & ",InvoiceLineAmount,memo,FQSaveToCache) VALUES ('" & strCustIdAop & "','" & strBU & "','CUSTOMER RECEIVABLE (" & strCur & ")'," _
    '            & strTrxDate & ",'" & strRefNumber & "','REVENUE','" & strMemo _
    '            & "'," & 0 - decMerchantFee & ",'" & strMemo & "',1)"
    '        If CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, False) = 0 Then
    '            MsgBox("Không tạo được bản ghi AOP Invoice cho TrxCode " & strTrxCode)
    '            Return 0
    '        End If
    '    End If

    '    strQuerry = "INSERT INTO Invoice (CustomerRefListID,ClassRefFullName, ARAccountRefFullName, TxnDate, RefNumber" _
    '            & ", Memo) VALUES ('" & strCustIdAop & "','" & strBU & "','CUSTOMER RECEIVABLE (" & strCur & ")'," _
    '            & strTrxDate & ",'" & strRefNumber & "','" & strMemo & "')"

    '    If CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, False, strQuerryType) = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP Invoice cho TrxCode " & strTrxCode)
    '        Return 0
    '    End If
    '    Return intQueueRecId
    'End Function
    'Private Function CreateAopQueueJournalEntry(strBU As String, strTrxCode As String, strProd As String _
    '                                , strTrxDate As String, strRefNumber As String, strCurrencyRefFullName As String _
    '                                , strJournalCreditLineAccountRefFullName As String, decAmount As Decimal _
    '                                , strMemo As String, strJournalCreditLineEntityRefFullName As String _
    '                                , strJournalDebitLineAccountRefFullName As String _
    '                                , strJournalDebitLineEntityRefFullName As String) As Integer
    '    Dim intQueueRecId As Integer
    '    Dim intNewRecId As Integer
    '    Dim strQuerry As String

    '    If Not strTrxDate.Contains("{d") Then
    '        strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
    '    End If

    '    strQuerry = "INSERT INTO JournalEntryCreditLine (TxnDate,RefNumber,CurrencyRefFullName" _
    '                    & ",JournalCreditLineAccountRefFullName,JournalCreditLineAmount,JournalCreditLineMemo, JournalCreditLineEntityRefFullName,FQSaveToCache) VALUES (" _
    '                    & strTrxDate & ",'" & strRefNumber & "','" & strCurrencyRefFullName & "','" & strJournalCreditLineAccountRefFullName _
    '                    & "'," & decAmount & ",'" & strMemo & "','" & strJournalCreditLineEntityRefFullName & "',1)"


    '    intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, "ODBC")

    '    If intQueueRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP JournalEntryCreditLine cho TrxCode " & strTrxCode)
    '        Return 0
    '    ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn(0)) Then
    '        MsgBox("Không tạo được bản ghi AOP JournalEntryCreditLine cho TrxCode " & strTrxCode)
    '        Return 0
    '    End If

    '    strQuerry = "INSERT INTO JournalEntryDebitLine (TxnDate,RefNumber,CurrencyRefFullName" _
    '                    & ",JournalDebitLineAccountRefFullName,JournalDebitLineAmount,JournalDebitLineMemo, JournalDebitLineEntityRefFullName,FQSaveToCache) VALUES (" _
    '                    & strTrxDate & ",'" & strRefNumber & "','" & strCurrencyRefFullName & "','" & strJournalDebitLineAccountRefFullName _
    '                    & "'," & decAmount & ",'" & strMemo & "','" & strJournalDebitLineEntityRefFullName & "',0)"

    '    intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, "ODBC")
    '    If intNewRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP JournalEntryDebitLine cho TrxCode " & strTrxCode)
    '        Return 0
    '    End If


    '    Return intQueueRecId
    'End Function
    'Private Function CreateAopQueueVendorCredit(strVendorAOPListid As String, strCustAopId As String, strBu As String _
    '                                , strAccountName As String, strTrxDate As String _
    '                               , strRefNumber As String, strItemName As String, decInvAmt As Decimal, decMerchantFee As Decimal _
    '                               , strMemo As String, strProd As String, strTrxCode As String) As Integer
    '    Dim strQuerry As String
    '    Dim intQueueRecId As Integer
    '    Dim intNewRecId As Integer

    '    If Not strTrxDate.Contains("{d") Then
    '        strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
    '    End If

    '    strQuerry = "INSERT INTO VendorCreditItemLine (VendorRefListID,APAccountRefFullName" _
    '                    & ",TxnDate,RefNumber,Memo,ItemLineItemRefFullName,ItemLineDesc" _
    '                    & ",ItemLineAmount,ItemLineCustomerRefListID,ItemLineClassRefFullName,ItemLineBillableStatus,FQSaveToCache) VALUES ('" _
    '                    & strVendorAOPListid & "','" & strAccountName & "'," & strTrxDate & ",'" & strRefNumber & "','" & strMemo _
    '                    & "','" & strItemName & "','" & strMemo _
    '                    & "'," & decInvAmt & ",'" & strCustAopId & "','" & strBu & "','Billable',1)"

    '    intQueueRecId = CreateAopQueueRecord(0, "B", strProd, strBu, strTrxCode, strQuerry, True, "ODBC")

    '    If intQueueRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP VendorCredit cho TrxCode " & strTrxCode)
    '        Return 0
    '    ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn(0)) Then
    '        MsgBox("Không tạo được bản ghi AOP VendorCredit cho TrxCode " & strTrxCode)
    '        Return 0
    '    End If

    '    strQuerry = "INSERT INTO VendorCredit (VendorRefListID,APAccountRefFullName, TxnDate, RefNumber" _
    '                        & ", Memo) VALUES ('" & strVendorAOPListid & "','" & strAccountName & "'," _
    '                        & strTrxDate & ",'" & strRefNumber & "','" & strMemo & "')"

    '    intNewRecId = CreateAopQueueRecord(intQueueRecId, "B", strProd, strBu, strTrxCode, strQuerry, True, "ODBC")

    '    If intNewRecId = 0 Then
    '        Return 0
    '    End If

    '    Return intQueueRecId

    'End Function
    'Private Function CreateAopQueueBill(strVendorAOPListid As String, strCustAopId As String, strBu As String _
    '                                , strTrxDate As String _
    '                               , strRefNumber As String, strItemName As String, decBillAmt As Decimal, decMerchantFee As Decimal _
    '                               , strMemo As String, strProd As String, strTrxCode As String, strCur As String, strDueDate As String) As Integer

    '    Dim strQuerry As String
    '    Dim intQueueRecId As Integer
    '    Dim intNewRecId As Integer

    '    If Not strTrxDate.Contains("{d") Then
    '        strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
    '    End If

    '    If strDueDate = "" Then
    '        strQuerry = "INSERT INTO BillItemLine (ItemLineBillableStatus, ItemLineCustomerRefListID" _
    '            & ",VendorRefListID,ItemLineClassRefFullName,APAccountRefFullName" _
    '            & ",TxnDate,RefNumber,ItemLineItemRefFullName, ItemLineDesc" _
    '            & ",ItemLineAmount,memo,FQSaveToCache) VALUES ('Billable','" & strCustAopId & "','" & strVendorAOPListid _
    '            & "','" & strBu & "','VENDOR PAYABLE (" & strCur & ")'," _
    '            & strTrxDate & ",'" & strRefNumber & "','" & strItemName & "','" & strMemo _
    '            & "'," & decBillAmt & ",'" & strMemo & "',1)"
    '    Else
    '        If Not strDueDate.Contains("{d") Then
    '            strDueDate = "{d'" & Format(CDate(strDueDate), "yyyy-MM-dd") & "'}"
    '        End If
    '        strQuerry = "INSERT INTO BillItemLine (ItemLineBillableStatus, ItemLineCustomerRefListID" _
    '            & ",VendorRefListID,ItemLineClassRefFullName,APAccountRefFullName" _
    '            & ",TxnDate,RefNumber,ItemLineItemRefFullName, ItemLineDesc" _
    '            & ",ItemLineAmount,memo,DueDate,FQSaveToCache) VALUES ('Billable','" & strCustAopId & "','" & strVendorAOPListid _
    '            & "','" & strBu & "','VENDOR PAYABLE (" & strCur & ")'," _
    '            & strTrxDate & ",'" & strRefNumber & "','" & strItemName & "','" & strMemo _
    '            & "'," & decBillAmt & ",'" & strMemo & "'," & strDueDate & ",1)"
    '    End If

    '    intQueueRecId = CreateAopQueueRecord(intQueueRecId, "B", strProd, strBu, strTrxCode, strQuerry, True, "ODBC")

    '    If intQueueRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP Bill cho Trxcode " & strMemo)
    '        Return 0
    '    ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn(0)) Then
    '        MsgBox("Không tạo được bản ghi AOP Bill cho TrxCode " & strMemo)
    '        Return 0
    '    End If

    '    strQuerry = "INSERT INTO Bill (VendorRefListID, APAccountRefFullName, TxnDate, RefNumber" _
    '            & ", Memo) VALUES ('" & strVendorAOPListid & "','VENDOR PAYABLE (" & strCur & ")'," _
    '            & strTrxDate & ",'" & strRefNumber & "','" & strMemo & "')"

    '    intNewRecId = CreateAopQueueRecord(intQueueRecId, "B", strProd, strBu, strTrxCode, strQuerry, True, "ODBC")

    '    If intNewRecId = 0 Then
    '        MsgBox("Không tạo được bản ghi AOP Bill cho TrxCode " & strMemo)
    '        Return 0
    '    End If

    '    Return intQueueRecId
    'End Function
    '^_^20221017 mark by 7643 -e-
    '^_^20221017 modi by 7643 -b-
    Public Function CreateAopQueueAirCTS(strRcp As String, ByRef conn As SqlClient.SqlConnection) As Boolean

        Dim strQuerry As String
        Dim lstQueueRecIds As New List(Of String)
        Dim tblInvoice As DataTable
        Dim tblBill As DataTable
        Dim strBu As String = String.Empty
        Dim intResult As Integer
        Dim arrQueueRecIds As String()
        Dim strMemo As String
        Dim strARAccountRefFullName As String = ""
        Dim strAccountName As String = ""
        Dim strItemLineItemRefFullName As String = ""

        'Tao Invoice
        strQuerry = "select (case when m.RecID is null then 1 else 2 end) as InvCount, r.CustId, r.RcpNo,R.Srv" _
                & ",r.TtlDue as InvAmt, 0 as SvcFee" _
                & ",r.Charge as MerchantFee" _
                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
                & ",l.CustShortName,AOPListID" _
                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS Tkno" _
                & " from Rcp r" _
                & " left join CustomerList l on l.Recid=r.CustId" _
                & " left join Misc m on r.CustId=m.intVal and m.Cat='CustNameInGroup' and m.VAL='2 INVOICES CUS' and m.Status='OK'" _
                & " where r.status='OK' and r.Counter='CWT' and r.RcpNo='" & strRcp & "'"

        tblInvoice = GetDataTable(strQuerry, conn)

        If tblInvoice.Rows.Count = 0 Then
            MsgBox("Invalide Rcp:" & strRcp)
            Return False
        End If
        If tblInvoice.Rows(0)("AOPListID") = "" Then
            MsgBox("You must ask PQT to update AOPListID for the following Customers: " & tblInvoice.Rows(0)("CustShortName"))
            Return False
        End If

        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
            strARAccountRefFullName = "CUSTOMER RECEIVABLE (VND)"
            strAccountName = "VENDOR PAYABLE (VND)"
            strItemLineItemRefFullName = "COST"
        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
            strARAccountRefFullName = "PHAI THU KHACH HANG"
            strAccountName = "PHAI TRA NGUOI BAN"
            strItemLineItemRefFullName = "COGS:Cost"
        End If

        For Each objRow As DataRow In tblInvoice.Rows
            strBu = "CTS-AIR"

            If objRow("Tkno").ToString.EndsWith("/") Then
                strMemo = Mid(objRow("Tkno"), 1, Len(objRow("Tkno")) - 1)
            Else
                strMemo = objRow("Tkno")
            End If
            If objRow("InvCount") = 1 Then
                If objRow("SRV") = "S" Then
                    intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), objRow("MerchantFee"), objRow("AOPListid") _
                                                      , objRow("TrxDate"), objRow("RefNumber"), strMemo, conn)
                    If intResult = 0 Then
                        Return False
                    Else
                        lstQueueRecIds.Add(intResult.ToString)
                    End If
                Else
                    intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, strARAccountRefFullName, objRow("TrxDate") _
                                                         , objRow("RefNumber"), strMemo, objRow("InvAmt"), objRow("MerchantFee"), strMemo, strRcp, "Air", conn)
                    If intResult = 0 Then
                        Return False
                    Else
                        lstQueueRecIds.Add(intResult.ToString)
                    End If
                End If
            ElseIf objRow("InvCount") = 2 Then
                objRow("SvcFee") = ScalarToDec("tkt", "sum(ChargeTV)", "Status<>'xx' and Rcpno='" & objRow("RcpNo") & "'")
                objRow("InvAmt") = objRow("InvAmt") - objRow("SvcFee")

                If objRow("SRV") = "S" Then
                    Dim decMerchantFee As Decimal
                    intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), objRow("MerchantFee"), objRow("AOPListid") _
                                                      , objRow("TrxDate"), objRow("RefNumber"), strMemo, conn)
                    If intResult = 0 Then
                        Return False
                    Else
                        lstQueueRecIds.Add(intResult.ToString)
                    End If

                    Select Case objRow("CustShortName")
                        Case "PG VIETNAM", "PG INDOCHINA"
                            decMerchantFee = 0
                        Case Else
                            decMerchantFee = objRow("MerchantFee")
                    End Select
                    If objRow("SvcFee") > 0 Then
                        intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("SvcFee"), 0, objRow("AOPListid") _
                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo, conn)
                        If intResult = 0 Then
                            Return False
                        Else
                            lstQueueRecIds.Add(intResult.ToString)
                        End If
                    End If

                Else
                    intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, strARAccountRefFullName, objRow("TrxDate") _
                                                         , objRow("RefNumber"), objRow("Tkno"), objRow("InvAmt"), objRow("MerchantFee"), strMemo, strRcp, "Air", conn)
                    If intResult = 0 Then
                        Return False
                    Else
                        lstQueueRecIds.Add(intResult.ToString)
                    End If
                    intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, strARAccountRefFullName, objRow("TrxDate") _
                                                         , objRow("RefNumber"), objRow("Tkno"), objRow("SvcFee"), objRow("MerchantFee"), strMemo, strRcp, "Air", conn)
                    If intResult = 0 Then
                        Return False
                    Else
                        lstQueueRecIds.Add(intResult.ToString)
                    End If
                End If

            End If
        Next

        'Tao bill
        strQuerry = "select r.Vendor,c.CustShortName, r.RcpNo,R.Srv" _
                & " ,(select top 1 DOI from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0) AS DOI" _
                & " ,(select sum((NetToAL+Tax)+Charge*t.Qty) from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0) AS BillAmt" _
                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
                & " ,(select distinct t.DocType+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS DocType" _
                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS Tkno" _
                & ",v.AOPListID as VendorAopId,c.AOPListID as CustAopId " _
                & " from Rcp r" _
                & " left join CustomerList c on c.Recid=r.CustId" _
                & " left join Vendor v on v.Recid=r.VendorId" _
                & " where r.status='OK' and r.Srv<>'V' and r.Counter='CWT' and r.rcpno='" & strRcp & "'" _
                & " and r.RecId not in (Select RcpId from tkt where DocType='AHC')" _
                & " order by r.FstUpdate"

        tblBill = GetDataTable(strQuerry, conn)

        For Each objRow As DataRow In tblBill.Rows
            If objRow("Vendor") = "" Then
                MsgBox("You must ask CTS to update Vendor for " & strRcp)
                Return False
            ElseIf objRow("VendorAopId") = "" Then
                MsgBox("You must ask PQT to update AOPListID for the following Vendor:  " & objRow("Vendor"))
                Return False
            End If
        Next

        For Each objRow As DataRow In tblBill.Rows
            If objRow("Tkno").ToString.EndsWith("/") Then
                strMemo = Mid(objRow("Tkno"), 1, Len(objRow("Tkno")) - 1)
            Else
                strMemo = objRow("Tkno")
            End If

            If objRow("SRV") = "R" Then
                intResult = CreateAopQueueVendorCredit(objRow("VendorAopId"), objRow("CustAopId"), strBu, strAccountName _
                              , objRow("TrxDate"), objRow("RefNumber"), strItemLineItemRefFullName, objRow("BillAmt"), 0, strMemo, "Air", strRcp, conn)
                If intResult = 0 Then
                    Return False
                Else
                    lstQueueRecIds.Add(intResult.ToString)
                End If
            Else
                Dim strDueDate As String = String.Empty
                Select Case objRow("Vendor")
                    Case "VN", "BSP", "VN DEB"
                        If objRow("DocType").ToString.Contains("ETK") Or objRow("DocType").ToString.Contains("EMD") _
                            Or objRow("DocType").ToString.Contains("MCO") Then
                            If objRow("Vendor") = "BSP" Then
                                strDueDate = GetDueDate4AopBsp(objRow("DOI"))
                            Else
                                strDueDate = GetDueDate4AopNonBsp(objRow("DOI"))
                            End If
                        End If
                    Case "QH TK"
                        strDueDate = Format(objRow("DOI"), "yyyy-MM-dd")
                End Select

                intResult = CreateAopQueueBill(objRow("VendorAopId"), objRow("CustAOPid"), strBu, objRow("TrxDate") _
                                                         , objRow("RefNumber"), strItemLineItemRefFullName, objRow("BillAmt"), 0 _
                                                         , strMemo, "Air", strRcp, "VND", strDueDate, conn)
                If intResult = 0 Then
                    Return False
                Else
                    lstQueueRecIds.Add(intResult.ToString)
                End If
            End If
        Next
        ReDim arrQueueRecIds(0 To lstQueueRecIds.Count - 1)
        lstQueueRecIds.CopyTo(arrQueueRecIds)
        Return ExecuteNonQuerry("update AopQueue Set Status='OK' where LinkId in (" & Strings.Join(arrQueueRecIds, ",") & ")", conn)
    End Function
    Public Function CreateAopQueueAirTVS(strRcp As String, dteTrxDate As Date, ByRef conn As SqlClient.SqlConnection) As Boolean
        Dim decAopUsdRoe As Decimal = ScalarToDec("Forex", "TOP 1 BSR", "Status='OK' and Currency='USD' and ApplyROETo='AOP'" _
                                                  & " and EffectDate <='" & CreateFromDate(dteTrxDate) _
                                                  & "' order by EffectDate")
        Dim tblInvoice As DataTable
        Dim tblBill As DataTable
        Dim objRow As DataRow
        Dim lstQueueRecIds As New List(Of String)
        Dim intResult As Integer
        Dim strBu As String = "TVS"
        Dim strMemo As String
        Dim arrQueueRecIds As String()
        Dim strARAccountRefFullName As String = ""
        Dim strAccountName As String = ""
        Dim strItemLineItemRefFullName As String = ""

        '^_^20221017 add by 7643 -b-
        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
            strBu = "TVS"
        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
            strBu = "Travel:TVS"
        End If
        '^_^20221017 add by 7643 -e-

        Dim strQuerry As String = "select r.CustId, r.RcpNo,R.Srv" _
                & " ,(case (select count (*) from tkt t where t.Status<>'XX' and t.RCPID=r.RecID and T.DocType in ('GRP','MCO')) when 0 then 'FIT' else 'GRP' end) AS GRP" _
                & ",'' as AopRecord" _
                & ",r.Roe*r.TtlDue as InvAmt, 0 as SvcFee" _
                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
                & ",l.CustShortName,'' as TourCode" _
                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID  ORDER BY TKNO For Xml path('') ) AS Tkno,'' as DepDate" _
                & ",'' as Account,'' as AccountName" _
                & ",'' as OriDeposits" _
                & " ,(select distinct t.DocType+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS DocTypes,AOPListID" _
                & ",r.Vendor,r.VendorId" _
                & ", (case r.CustShortName when 'VYM' then " & decAopUsdRoe & " else 1 end)  as ROE" _
                & " from Rcp r" _
                & " left join CustomerList l on l.Recid=r.CustId" _
                & " where r.status='OK' and r.Counter='TVS' and RcpNo='" & strRcp _
                & "' order by r.FstUpdate"

        tblInvoice = GetDataTable(strQuerry, conn)

        If tblInvoice.Rows.Count = 0 Then
            Append2TextFile("Invalid RCP " & strRcp)
            Return False
        Else
            objRow = tblInvoice.Rows(0)
        End If

        If objRow("AOPListID") = "" Then
            MsgBox("Missing AOPListID for " & objRow("CustShortName"))
            Return False
        ElseIf objRow("AopRecord") = "Deposit" Then
            'MsgBox("You must update account for " & objRow("RcpNo"))
            Return False
        End If

        objRow("tkno") = Mid(objRow("tkno"), 1, Len(objRow("Tkno")) - 1)
        strMemo = objRow("tkno")


        'import Invoice
        Select Case objRow("GRP")
            Case "GRP"
                'Bo khong import tu dong

            Case "FIT"
                Select Case objRow("CustShortName")

                    Case "TVHAN"
                        'Bo qua ko nhap

                    Case "TVSGN", "GDSSGN"
                        strMemo = strMemo & GetColumnValuesAsString("FOP", "Document", "WHERE Status='OK' and RcpNO='" & strRcp & "' and FOP<>'EXC'", "|")
                        Dim decExchangeAmt As Decimal = ScalarToDec("FOP", "Amount", "Status='OK' and FOP='EXC' and RcpNo='" & strRcp & "'")

                        objRow("InvAmt") = objRow("InvAmt") - decExchangeAmt

                        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
                            strARAccountRefFullName = "CUSTOMER RECEIVABLE (VND)"
                        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
                            strARAccountRefFullName = "PHAI THU KHACH HANG"
                        End If

                        If objRow("SRV") = "S" Then
                            intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), 0, objRow("AOPListid") _
                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo, conn,, "ODBC")
                            If intResult = 0 Then
                                Return False
                            Else
                                lstQueueRecIds.Add(intResult.ToString)
                            End If

                        Else    'refund
                            intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, strARAccountRefFullName, objRow("TrxDate") _
                                                             , objRow("RefNumber"), objRow("Tkno"), objRow("InvAmt"), 0, strMemo, strRcp, "Air", conn)
                            If intResult = 0 Then
                                Return False
                            Else
                                lstQueueRecIds.Add(intResult.ToString)
                            End If
                        End If
                    Case "VYM" 'VAYMA (TRAVIX)
                        strMemo = objRow("Tkno") & " (" & Format(objRow("InvAmt"), "#,##0") & "/" & Format(objRow("ROE"), "#,##0") & ")"
                        Dim decInvAmt As Decimal = Math.Round(objRow("InvAmt") / objRow("ROE"), 2)

                        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
                            strARAccountRefFullName = "CUSTOMER RECEIVABLE (USD)"
                        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
                            strARAccountRefFullName = "PHAI THU KHACH HANG"
                        End If

                        If objRow("SRV") = "S" Then
                            intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, decInvAmt, 0, objRow("AOPListid") _
                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo, conn, "USD")
                            If intResult = 0 Then
                                Return False
                            Else
                                lstQueueRecIds.Add(intResult.ToString)
                            End If

                        Else
                            intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, strARAccountRefFullName, objRow("TrxDate") _
                                                             , objRow("RefNumber"), objRow("Tkno"), decInvAmt, 0, strMemo, strRcp, "Air", conn)
                            If intResult = 0 Then
                                Return False
                            Else
                                lstQueueRecIds.Add(intResult.ToString)
                            End If

                        End If

                    Case Else
                        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
                            strARAccountRefFullName = "CUSTOMER RECEIVABLE (VND)"
                        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
                            strARAccountRefFullName = "PHAI THU KHACH HANG"
                        End If

                        If objRow("SRV") = "S" Then
                            intResult = CreateAopQueueInvoice(0, "I", "Air", strBu, strRcp, objRow("InvAmt"), 0, objRow("AOPListid") _
                                                          , objRow("TrxDate"), objRow("RefNumber"), strMemo, conn)
                            If intResult = 0 Then
                                Return False
                            Else
                                lstQueueRecIds.Add(intResult.ToString)
                            End If

                        Else    'refund
                            intResult = CreateAopQueueCreditMemo(objRow("AOPListid"), strBu, strARAccountRefFullName, objRow("TrxDate") _
                                                             , objRow("RefNumber"), objRow("Tkno"), objRow("InvAmt"), 0, strMemo, strRcp, "Air", conn)
                            If intResult = 0 Then
                                Return False
                            Else
                                lstQueueRecIds.Add(intResult.ToString)
                            End If
                        End If
                End Select
        End Select

        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
            strAccountName = "VENDOR PAYABLE (VND)"
            strItemLineItemRefFullName = "COST"
        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
            strAccountName = "PHAI TRA NGUOI BAN"
            strItemLineItemRefFullName = "COGS:Cost"
        End If

        ' Import for TSP
        Select Case objRow("CustShortName")
            Case "TVHAN"
                Dim tblTourCodes As DataTable = GetTourCodeTableByRcp(strRcp, conn)
                Dim strJournalCreditLineEntityRefFullName As String = objRow("Vendor") & " [" & objRow("VendorId") & "]"

                If tblTourCodes.Rows.Count > 0 Then
                    strMemo = objRow("Tkno")
                    For Each objToourCodeRow As DataRow In tblTourCodes.Rows
                        strMemo = strMemo & " " & objToourCodeRow("TourCode")
                    Next
                End If

                If objRow("SRV") = "S" Then
                    intResult = CreateAopQueueJournalEntry(strBu, strRcp, "Air", objRow("TrxDate"), objRow("RefNumber"), "Vietnamese Dong" _
                                                               , strAccountName, objRow("InvAmt"), strMemo, strJournalCreditLineEntityRefFullName _
                                                               , "INTERNAL DEBT:AIR HAN", "TVHAN [68612]", conn)
                    If intResult = 0 Then
                        Return False
                    Else
                        lstQueueRecIds.Add(intResult.ToString)
                    End If
                Else
                    intResult = CreateAopQueueJournalEntry(strBu, strRcp, "Air", objRow("TrxDate"), objRow("RefNumber"), "Vietnamese Dong" _
                                                               , "INTERNAL DEBT:AIR HAN", objRow("InvAmt"), strMemo, "TVHAN [68612]" _
                                                               , strAccountName, strJournalCreditLineEntityRefFullName, conn)
                    If intResult = 0 Then
                        Return False
                    Else
                        lstQueueRecIds.Add(intResult.ToString)
                    End If
                End If

            Case "TVSGN", "GDSSGN"
                Dim tblTourCodes As DataTable = GetTourCodeTableByRcp(strRcp, conn)

                If tblTourCodes.Rows.Count > 0 Then
                    strMemo = objRow("Tkno")
                    For Each objToourCodeRow As DataRow In tblTourCodes.Rows
                        strMemo = strMemo & " " & objToourCodeRow("TourCode")
                    Next

                    If objRow("SRV") = "S" Then
                        intResult = CreateAopQueueCheck(strBu, strRcp, "Air", "CASH AIR", "AL TRANSVIET 2020 [10015]", objRow("RefNumber"), objRow("TrxDate") _
                                            , objRow("InvAmt"), "Vietnamese Dong", strMemo, strAccountName, conn)
                        If intResult = 0 Then
                            Return False
                        Else
                            lstQueueRecIds.Add(intResult.ToString)
                        End If
                    Else
                        intResult = CreateAopQueueJournalEntry(strBu, strRcp, "Air", objRow("TrxDate"), objRow("RefNumber"), "Vietnamese Dong" _
                                                               , strAccountName, objRow("InvAmt"), strMemo, "AL TRANSVIET 2020 [10015]" _
                                                               , "CASH AIR", "AL TRANSVIET 2020 [10015]", conn)
                        If intResult = 0 Then
                            Return False
                        Else
                            lstQueueRecIds.Add(intResult.ToString)
                        End If
                    End If
                End If
        End Select

        'IMPORT BILL
        If objRow("CustShortName") <> "TVHAN" Then
            strQuerry = "select r.Vendor,c.CustShortName, r.RcpNo,R.Srv" _
                & " ,(select top 1 DOI from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0) AS DOI" _
                & " ,((select sum((NetToAL+Tax)+Charge*t.Qty) from tkt t where t.Status<>'XX' and t.StatusAL<>'XX' and t.RCPID=r.RecID and t.Qty<>0)" _
                & " - (select isnull(sum(Amount),0) from FOP f where f.Status='OK' and f.RCPID=r.RecID and f.FOP='EXC'))*r.Roe  AS BillAmt" _
                & ",substring(r.RcpNo,1,6)+ substring(r.RcpNo,9,4) as RefNumber" _
                & ",CONVERT(VARCHAR,r.FstUpdate,23) as TrxDate" _
                & " ,(select distinct t.DocType+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS DocType" _
                & " ,(select t.Tkno+'/' from tkt t where t.Status<>'XX' and t.RCPID=r.RecID For Xml path('') ) AS Tkno" _
                & ",'' AS TourCode,'' as DepDate,'' as TspCust,'' as TspCustAopId,'' as TspClass" _
                & ",v.AOPListID as VendorAopId,c.AOPListID as CustAopId,r.CustId,r.Currency as RcpCur,v.Cur as VenCur,r.TtlDue " _
                & " from Rcp r" _
                & " left join CustomerList c on c.Recid=r.CustId" _
                & " left join Vendor v on v.Recid=r.VendorId" _
                & " where r.status='OK' and r.Srv<>'V' and r.Counter='TVS' and r.RcpNo='" & strRcp & "'" _
                & " and r.RecId not in (Select RcpId from tkt where DocType='AHC')"

            tblBill = GetDataTable(strQuerry, conn)
            objRow = tblBill.Rows(0)

            If IsDBNull(objRow("VendorAopId")) Then
                MsgBox("You must ask PQT to update AOPListID for the following Vendor: " & objRow("Vendor"))
                Return False
            End If
            If objRow("CustAopId") = "" Then
                MsgBox("You must ask PQT to update AOPListID for the following Customer: " & objRow("CustShortName"))
                Return False
            End If

            objRow("Tkno") = Mid(objRow("tkno"), 1, Len(objRow("Tkno")) - 1)
            strMemo = objRow("Tkno")

            If IsDBNull(objRow("BillAmt")) Then
                GoTo BillNotRequired    've XX doi voi Hang
            End If
            Dim strBillCur As String = "VND"
            Dim decBillAmt As Decimal = objRow("BillAmt")
            If objRow("Vendor") = "AK TK" Then
                strBillCur = "USD"
                decBillAmt = objRow("TtlDue")
            End If

            If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
                strAccountName = "VENDOR PAYABLE (" & strBillCur & ")"
            ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
                strAccountName = "PHAI TRA NGUOI BAN"
            End If

            If objRow("SRV") = "R" Then
                intResult = CreateAopQueueVendorCredit(objRow("VendorAopId"), objRow("CustAopId"), strBu, strAccountName _
                                  , objRow("TrxDate"), objRow("RefNumber"), strItemLineItemRefFullName, decBillAmt, 0, strMemo, "Air", strRcp, conn)
                If intResult = 0 Then
                    Return False
                Else
                    lstQueueRecIds.Add(intResult.ToString)
                End If

            ElseIf objRow("BillAmt") > 0 Then       'bo qua ve zero value
                Dim strDueDate As String = String.Empty
                Select Case objRow("Vendor")
                    Case "BSP", "BSP AERTICKET-37314944"
                        If objRow("DocType").ToString.Contains("ETK") Or objRow("DocType").ToString.Contains("EMD") _
                                                Or objRow("DocType").ToString.Contains("MCO") Then
                            strDueDate = GetDueDate4AopBsp(objRow("DOI"))
                        End If
                    Case "VN", "VN DEB"
                        If objRow("DocType").ToString.Contains("ETK") Or objRow("DocType").ToString.Contains("EMD") _
                                                Or objRow("DocType").ToString.Contains("MCO") Then
                            strDueDate = GetDueDate4AopNonBsp(objRow("DOI"))
                        End If

                    Case "QH TK"
                        strDueDate = Format(objRow("DOI"), "yyyy-MM-dd")
                End Select

                intResult = CreateAopQueueBill(objRow("VendorAopId"), objRow("CustAOPid"), strBu, objRow("TrxDate") _
                                                             , objRow("RefNumber"), strItemLineItemRefFullName, decBillAmt, 0 _
                                                             , strMemo, "Air", strRcp, strBillCur, strDueDate, conn)
                If intResult = 0 Then
                    Return False
                Else
                    lstQueueRecIds.Add(intResult.ToString)
                End If

            End If
        End If


BillNotRequired:
        ReDim arrQueueRecIds(0 To lstQueueRecIds.Count - 1)
        lstQueueRecIds.CopyTo(arrQueueRecIds)
        Return ExecuteNonQuerry("update AopQueue Set Status='OK' where LinkId in (" & Strings.Join(arrQueueRecIds, ",") & ")", conn)

    End Function
    Private Function CreateAopQueueCheck(strBU As String, strTrxCode As String, strProd As String _
                                    , strAccountRefFullName As String, strPayeeEntityRefFullName As String _
                                    , strRefNumber As String, strTrxDate As String, decAmount As Decimal _
                                    , strCurrencyRefFullName As String, strMemo As String _
                                    , strExpenseLineAccountRefFullName As String, ByRef conn As SqlClient.SqlConnection) As Integer
        Dim intQueueRecId As Integer
        Dim intNewRecId As Integer
        Dim strQuerry As String

        If Not strTrxDate.Contains("{d") Then
            strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
        End If

        strQuerry = "INSERT INTO [CheckExpenseLine] (AccountRefFullName,PayeeEntityRefFullName,RefNumber" _
                        & ",TxnDate,Memo,ExpenseLineAccountRefFullName,ExpenseLineAmount" _
                        & " ,ExpenseLineMemo,ExpenseLineCustomerRefFullName,IsToBePrinted,FQSaveToCache) VALUES ('" _
                        & strAccountRefFullName & "','" & strPayeeEntityRefFullName & "','" & strRefNumber & "'," _
                        & strTrxDate & ",'" & strMemo & "','" & strExpenseLineAccountRefFullName _
                        & "'," & decAmount & ",'" & strMemo & "','" & strPayeeEntityRefFullName & "',0,1)"


        intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, "ODBC")
        If intQueueRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP CheckExpenseLine cho TrxCode " & strTrxCode)
            Return 0
        ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn) Then
            MsgBox("Không tạo được bản ghi AOP Check cho TrxCode " & strTrxCode)
            Return 0
        End If

        strQuerry = "INSERT INTO [Check] (AccountRefFullName,PayeeEntityRefFullName,RefNumber" _
                        & ",TxnDate, Memo,IsToBePrinted) VALUES ('" & strAccountRefFullName & "','" & strPayeeEntityRefFullName & "','" & strRefNumber & "'," _
                        & strTrxDate & ",'" & strMemo & "',0)"

        intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, "ODBC")
        If intNewRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP Check cho TrxCode " & strTrxCode)
            Return 0
        End If


        Return intQueueRecId
    End Function
    Private Function CreateAopQueueCreditMemo(strAOPListid As String, strBU As String, strAccountName As String, strInvDate As String _
                                   , strRefNumber As String, strDesc As String, decInvAmt As Decimal, decMerchantFee As Decimal _
                                   , strMemo As String, strTrxCode As String, strProd As String, ByRef conn As SqlClient.SqlConnection) As Integer
        Dim intQueueRecId As Integer
        Dim intNewRecId As Integer
        Dim strQuerry As String
        Dim strItemName As String

        If strProd = "Air" Then
            strItemName = "REVENUE"
        Else
            strItemName = strDesc
        End If
        If Not strInvDate.Contains("{d") Then
            strInvDate = "{d'" & Format(CDate(strInvDate), "yyyy-MM-dd") & "'}"
        End If
        strQuerry = "INSERT INTO CreditMemoLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
                        & ",TxnDate,RefNumber,CreditMemoLineItemRefFullName, CreditMemoLineDesc" _
                        & ",CreditMemoLineAmount,memo,FQSaveToCache) VALUES ('" & strAOPListid & "','" & strBU & "','" & strAccountName & "'," _
                        & strInvDate & ",'" & strRefNumber & "','" & strItemName & "','" & strDesc _
                        & "'," & decInvAmt & ",'" & strMemo & "',1)"

        intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, "ODBC")

        If intQueueRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP CreditMemo cho TrxCode " & strTrxCode)
            Return 0
        ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn) Then
            MsgBox("Không tạo được bản ghi AOP CreditMemo cho Tcode " & strTrxCode)
            Return 0
        End If

        If decMerchantFee <> 0 Then
            strQuerry = "INSERT INTO CreditMemoLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
                        & ",TxnDate,RefNumber,CreditMemoLineItemRefFullName, CreditMemoLineDesc" _
                        & ",CreditMemoLineAmount,memo,FQSaveToCache) VALUES ('" & strAOPListid & "','" & strBU & "','" & strAccountName & "'," _
                        & strInvDate & ",'" & strRefNumber & "','" & strItemName & "','" & strDesc _
                        & "'," & 0 - decMerchantFee & ",'" & strMemo & "',1)"
            intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, "ODBC")
            If intNewRecId = 0 Then
                MsgBox("Không tạo được bản ghi AOP CreditMemo Phí cà thẻ cho TrxCode" & strTrxCode)
                Return 0
            End If
        End If

        strQuerry = "INSERT INTO CreditMemo (CustomerRefListID,ClassRefFullName, ARAccountRefFullName, TxnDate, RefNumber" _
                            & ", Memo) VALUES ('" & strAOPListid & "','" & strBU & "','" & strAccountName & "'," _
                            & strInvDate & ",'" & strRefNumber & "','" & strMemo & "')"

        intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, "ODBC")
        If intNewRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP CreditMemo cho TrxCode " & strTrxCode)
            Return 0
        End If
        Return intQueueRecId
    End Function
    Public Function CreateAopQueueRecord(intLinkId As Integer, strBillInvoice As String, strProd As String, strBU As String, strTrxCode As String _
                                         , strQuerry As String, blnGetRecId As Boolean, ByRef conn As SqlClient.SqlConnection, Optional strQuerryType As String = "ODBC") As Integer
        Dim cmd As SqlClient.SqlCommand = conn.CreateCommand

        cmd.CommandText = "insert Into AopQueue (LinkId,B_I, Prod,BU, TrxCode, Querry, FstUser,QuerryType) values (" & intLinkId & ",'" & strBillInvoice _
                        & "','" & strProd & "','" & strBU & "','" & strTrxCode & "',@Querry,'AUT','" & strQuerryType & "')"

        If blnGetRecId Then
            cmd.CommandText = cmd.CommandText & ";SELECT SCOPE_IDENTITY() AS [RecID]"
        End If
        cmd.Parameters.Add("@Querry", SqlDbType.NVarChar).Value = strQuerry
        Try
            If blnGetRecId Then
                Return cmd.ExecuteScalar
            Else
                Return cmd.ExecuteNonQuery
            End If
        Catch ex As Exception
            Append2TextFile("SQL error:" & ex.Message & vbCrLf & cmd.CommandText & vbCrLf & strQuerry)
            Return 0
        End Try
    End Function
    Public Function CreateAopQueueInvoice(intLinkId As Integer, strBillInvoice As String, strProd As String, strBU As String, strTrxCode As String _
                                         , decInvAmout As Decimal, decMerchantFee As Decimal _
                                         , strCustIdAop As String, strTrxDate As String, strRefNumber As String, strMemo As String, ByRef conn As SqlClient.SqlConnection _
                                         , Optional strCur As String = "", Optional strQuerryType As String = "ODBC") As Integer
        Dim intQueueRecId As Integer
        Dim strQuerry As String
        Dim strARAccountRefFullName As String = ""
        Dim strInvoiceLineItemRefFullName As String = ""

        If strCur = "" Then
            strCur = "VND"
        End If
        If Not strTrxDate.Contains("{d") Then
            strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
        End If

        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
            strARAccountRefFullName = "CUSTOMER RECEIVABLE (" & strCur & ")"
            strInvoiceLineItemRefFullName = "REVENUE"
        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
            strARAccountRefFullName = "PHAI THU KHACH HANG"
            strInvoiceLineItemRefFullName = "REVENUE:Revenue"
        End If

        Select Case strProd
            Case "NonAir"
                strQuerry = "INSERT INTO InvoiceLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
                    & ",TxnDate,RefNumber,InvoiceLineItemRefFullName, InvoiceLineDesc" _
                    & ",InvoiceLineAmount,memo,FQSaveToCache) VALUES ('" & strCustIdAop & "','" & strBU & "','" & strARAccountRefFullName & "'," _
                    & strTrxDate & ",'" & strRefNumber & "','" & strTrxCode & "','" & strTrxCode _
                    & "'," & decInvAmout & ",'" & strMemo & "',1)"
            Case "Air"
                strQuerry = "INSERT INTO InvoiceLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
                    & ",TxnDate,RefNumber,InvoiceLineItemRefFullName, InvoiceLineDesc" _
                    & ",InvoiceLineAmount,memo,FQSaveToCache) VALUES ('" & strCustIdAop & "','" & strBU & "','" & strARAccountRefFullName & "'," _
                    & strTrxDate & ",'" & strRefNumber & "','" & strInvoiceLineItemRefFullName & "','" & strMemo _
                    & "'," & decInvAmout & ",'" & strMemo & "',1)"

        End Select

        intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, strQuerryType)

        If intQueueRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP Invoice cho TrxCode " & strTrxCode)
            Return 0
        ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn) Then
            MsgBox("Không tạo được bản ghi AOP Invoice cho Tcode " & strTrxCode)
            Return 0
        End If

        If decMerchantFee <> 0 Then
            strQuerry = "INSERT INTO InvoiceLine (CustomerRefListID,ClassRefFullName,ARAccountRefFullName" _
                & ",TxnDate,RefNumber,InvoiceLineItemRefFullName, InvoiceLineDesc" _
                & ",InvoiceLineAmount,memo,FQSaveToCache) VALUES ('" & strCustIdAop & "','" & strBU & "','" & strARAccountRefFullName & "'," _
                & strTrxDate & ",'" & strRefNumber & "','REVENUE','" & strMemo _
                & "'," & 0 - decMerchantFee & ",'" & strMemo & "',1)"
            If CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, False, conn) = 0 Then
                MsgBox("Không tạo được bản ghi AOP Invoice cho TrxCode " & strTrxCode)
                Return 0
            End If
        End If

        strQuerry = "INSERT INTO Invoice (CustomerRefListID,ClassRefFullName, ARAccountRefFullName, TxnDate, RefNumber" _
                & ", Memo) VALUES ('" & strCustIdAop & "','" & strBU & "','" & strARAccountRefFullName & "'," _
                & strTrxDate & ",'" & strRefNumber & "','" & strMemo & "')"

        If CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, False, conn, strQuerryType) = 0 Then
            MsgBox("Không tạo được bản ghi AOP Invoice cho TrxCode " & strTrxCode)
            Return 0
        End If
        Return intQueueRecId
    End Function
    Private Function CreateAopQueueJournalEntry(strBU As String, strTrxCode As String, strProd As String _
                                    , strTrxDate As String, strRefNumber As String, strCurrencyRefFullName As String _
                                    , strJournalCreditLineAccountRefFullName As String, decAmount As Decimal _
                                    , strMemo As String, strJournalCreditLineEntityRefFullName As String _
                                    , strJournalDebitLineAccountRefFullName As String _
                                    , strJournalDebitLineEntityRefFullName As String, ByRef conn As SqlClient.SqlConnection) As Integer
        Dim intQueueRecId As Integer
        Dim intNewRecId As Integer
        Dim strQuerry As String

        If Not strTrxDate.Contains("{d") Then
            strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
        End If

        strQuerry = "INSERT INTO JournalEntryCreditLine (TxnDate,RefNumber,CurrencyRefFullName" _
                        & ",JournalCreditLineAccountRefFullName,JournalCreditLineAmount,JournalCreditLineMemo, JournalCreditLineEntityRefFullName,FQSaveToCache) VALUES (" _
                        & strTrxDate & ",'" & strRefNumber & "','" & strCurrencyRefFullName & "','" & strJournalCreditLineAccountRefFullName _
                        & "'," & decAmount & ",'" & strMemo & "','" & strJournalCreditLineEntityRefFullName & "',1)"


        intQueueRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, "ODBC")

        If intQueueRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP JournalEntryCreditLine cho TrxCode " & strTrxCode)
            Return 0
        ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn) Then
            MsgBox("Không tạo được bản ghi AOP JournalEntryCreditLine cho TrxCode " & strTrxCode)
            Return 0
        End If

        strQuerry = "INSERT INTO JournalEntryDebitLine (TxnDate,RefNumber,CurrencyRefFullName" _
                        & ",JournalDebitLineAccountRefFullName,JournalDebitLineAmount,JournalDebitLineMemo, JournalDebitLineEntityRefFullName,FQSaveToCache) VALUES (" _
                        & strTrxDate & ",'" & strRefNumber & "','" & strCurrencyRefFullName & "','" & strJournalDebitLineAccountRefFullName _
                        & "'," & decAmount & ",'" & strMemo & "','" & strJournalDebitLineEntityRefFullName & "',0)"

        intNewRecId = CreateAopQueueRecord(intQueueRecId, "I", strProd, strBU, strTrxCode, strQuerry, True, conn, "ODBC")
        If intNewRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP JournalEntryDebitLine cho TrxCode " & strTrxCode)
            Return 0
        End If


        Return intQueueRecId
    End Function
    Private Function CreateAopQueueVendorCredit(strVendorAOPListid As String, strCustAopId As String, strBu As String _
                                    , strAccountName As String, strTrxDate As String _
                                   , strRefNumber As String, strItemName As String, decInvAmt As Decimal, decMerchantFee As Decimal _
                                   , strMemo As String, strProd As String, strTrxCode As String, ByRef conn As SqlClient.SqlConnection) As Integer
        Dim strQuerry As String
        Dim intQueueRecId As Integer
        Dim intNewRecId As Integer

        If Not strTrxDate.Contains("{d") Then
            strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
        End If

        strQuerry = "INSERT INTO VendorCreditItemLine (VendorRefListID,APAccountRefFullName" _
                        & ",TxnDate,RefNumber,Memo,ItemLineItemRefFullName,ItemLineDesc" _
                        & ",ItemLineAmount,ItemLineCustomerRefListID,ItemLineClassRefFullName,ItemLineBillableStatus,FQSaveToCache) VALUES ('" _
                        & strVendorAOPListid & "','" & strAccountName & "'," & strTrxDate & ",'" & strRefNumber & "','" & strMemo _
                        & "','" & strItemName & "','" & strMemo _
                        & "'," & decInvAmt & ",'" & strCustAopId & "','" & strBu & "','Billable',1)"

        intQueueRecId = CreateAopQueueRecord(0, "B", strProd, strBu, strTrxCode, strQuerry, True, conn, "ODBC")

        If intQueueRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP VendorCredit cho TrxCode " & strTrxCode)
            Return 0
        ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn) Then
            MsgBox("Không tạo được bản ghi AOP VendorCredit cho TrxCode " & strTrxCode)
            Return 0
        End If

        strQuerry = "INSERT INTO VendorCredit (VendorRefListID,APAccountRefFullName, TxnDate, RefNumber" _
                            & ", Memo) VALUES ('" & strVendorAOPListid & "','" & strAccountName & "'," _
                            & strTrxDate & ",'" & strRefNumber & "','" & strMemo & "')"

        intNewRecId = CreateAopQueueRecord(intQueueRecId, "B", strProd, strBu, strTrxCode, strQuerry, True, conn, "ODBC")

        If intNewRecId = 0 Then
            Return 0
        End If

        Return intQueueRecId

    End Function
    Private Function CreateAopQueueBill(strVendorAOPListid As String, strCustAopId As String, strBu As String _
                                    , strTrxDate As String _
                                   , strRefNumber As String, strItemName As String, decBillAmt As Decimal, decMerchantFee As Decimal _
                                   , strMemo As String, strProd As String, strTrxCode As String, strCur As String, strDueDate As String, ByRef conn As SqlClient.SqlConnection) As Integer

        Dim strQuerry As String
        Dim intQueueRecId As Integer
        Dim intNewRecId As Integer
        Dim strAccountName As String

        If Not strTrxDate.Contains("{d") Then
            strTrxDate = "{d'" & Format(CDate(strTrxDate), "yyyy-MM-dd") & "'}"
        End If

        If Not conn.ConnectionString.ToUpper.Contains("HAN") Then
            strAccountName = "VENDOR PAYABLE (" & strCur & ")"
        ElseIf conn.ConnectionString.ToUpper.Contains("HAN") Then
            strAccountName = "PHAI TRA NGUOI BAN"
        End If

        If strDueDate = "" Then
            strQuerry = "INSERT INTO BillItemLine (ItemLineBillableStatus, ItemLineCustomerRefListID" _
                & ",VendorRefListID,ItemLineClassRefFullName,APAccountRefFullName" _
                & ",TxnDate,RefNumber,ItemLineItemRefFullName, ItemLineDesc" _
                & ",ItemLineAmount,memo,FQSaveToCache) VALUES ('Billable','" & strCustAopId & "','" & strVendorAOPListid _
                & "','" & strBu & "','" & strAccountName & "'," _
                & strTrxDate & ",'" & strRefNumber & "','" & strItemName & "','" & strMemo _
                & "'," & decBillAmt & ",'" & strMemo & "',1)"
        Else
            If Not strDueDate.Contains("{d") Then
                strDueDate = "{d'" & Format(CDate(strDueDate), "yyyy-MM-dd") & "'}"
            End If
            strQuerry = "INSERT INTO BillItemLine (ItemLineBillableStatus, ItemLineCustomerRefListID" _
                & ",VendorRefListID,ItemLineClassRefFullName,APAccountRefFullName" _
                & ",TxnDate,RefNumber,ItemLineItemRefFullName, ItemLineDesc" _
                & ",ItemLineAmount,memo,DueDate,FQSaveToCache) VALUES ('Billable','" & strCustAopId & "','" & strVendorAOPListid _
                & "','" & strBu & "','" & strAccountName & "'," _
                & strTrxDate & ",'" & strRefNumber & "','" & strItemName & "','" & strMemo _
                & "'," & decBillAmt & ",'" & strMemo & "'," & strDueDate & ",1)"
        End If

        intQueueRecId = CreateAopQueueRecord(intQueueRecId, "B", strProd, strBu, strTrxCode, strQuerry, True, conn, "ODBC")

        If intQueueRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP Bill cho Trxcode " & strMemo)
            Return 0
        ElseIf Not ExecuteNonQuerry("Update AopQueue Set LinkId=" & intQueueRecId & " where RecId=" & intQueueRecId, conn) Then
            MsgBox("Không tạo được bản ghi AOP Bill cho TrxCode " & strMemo)
            Return 0
        End If

        strQuerry = "INSERT INTO Bill (VendorRefListID, APAccountRefFullName, TxnDate, RefNumber" _
                & ", Memo) VALUES ('" & strVendorAOPListid & "','" & strAccountName & "'," _
                & strTrxDate & ",'" & strRefNumber & "','" & strMemo & "')"

        intNewRecId = CreateAopQueueRecord(intQueueRecId, "B", strProd, strBu, strTrxCode, strQuerry, True, conn, "ODBC")

        If intNewRecId = 0 Then
            MsgBox("Không tạo được bản ghi AOP Bill cho TrxCode " & strMemo)
            Return 0
        End If

        Return intQueueRecId
    End Function
    '^_^20221017 modi by 7643 -e-
    Public Function GetDueDate4AopBsp(dteDOI As Date) As String
        Dim dteDueDate As Date
        Select Case dteDOI.Day
            Case < 9
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 9)
            Case < 16
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 16)
            Case < 24
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 24)
            Case Else
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 1).AddMonths(1).AddDays(-1)
        End Select
        Return Format(dteDueDate, "yyyy-MM-dd")
    End Function
    Public Function GetDueDate4AopNonBsp(dteDOI As Date) As String
        Dim dteDueDate As Date
        Select Case dteDOI.Day
            Case < 8
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 9)
            Case < 16
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 16)
            Case < 24
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 24)
            Case Else
                dteDueDate = DateSerial(dteDOI.Year, dteDOI.Month, 1).AddMonths(1).AddDays(-1)
        End Select
        Return Format(dteDueDate, "yyyy-MM-dd")
    End Function
    '^_^20221017 mark by 7643 -b-
    'Public Function CreateAopQueueAir(strRcpNo As String) As Boolean
    '    Dim tblRcp As DataTable = GetDataTable("Select * from Rcp where Status='ok' and Rcpno='" & strRcpNo & "'", conn(0))

    '    If tblRcp.Rows.Count = 0 Then
    '        Append2TextFile("Invalid RCPNO " & strRcpNo)
    '        'MsgBox("Invalid RCPNO " & strRcpNo)
    '        Return False
    '    End If
    '    Select Case tblRcp.Rows(0)("Counter")
    '        Case "CWT"
    '            Return CreateAopQueueAirCTS(strRcpNo)
    '        Case "TVS"
    '            Return CreateAopQueueAirTVS(strRcpNo, Now)
    '        Case Else
    '            Return True
    '    End Select

    'End Function
    '^_^20221017 mark by 7643 -e-
    '^_^20221017 modi by 7643 -b-
    Public Function CreateAopQueueAir(strRcpNo As String, ByRef conn As SqlClient.SqlConnection) As Boolean
        Dim tblRcp As DataTable = GetDataTable("Select * from Rcp where Status='ok' and Rcpno='" & strRcpNo & "'", conn)

        If tblRcp.Rows.Count = 0 Then
            Append2TextFile("Invalid RCPNO " & strRcpNo)
            Return False
        End If
        Select Case tblRcp.Rows(0)("Counter")
            Case "CWT"
                Return CreateAopQueueAirCTS(strRcpNo, conn)
            Case "TVS"
                Return CreateAopQueueAirTVS(strRcpNo, Now, conn)
            Case Else
                Return True
        End Select

    End Function
    '^_^20221017 modi by 7643 -e-
    Public Function CreateFromDate(dteInput As Date) As String
        'Return Format(dteInput, "dd MMM yy 00:00")  '^_^20221028 mark by 7643
        Return Format(dteInput, "dd MMM yyyy 00:00")  '^_^20221028 modi by 7643
    End Function
    Public Function CreateToDate(dteInput As Date) As String
        'Return Format(dteInput, "dd MMM yy 23:59")  '^_^20221028 mark by 7643
        Return Format(dteInput, "dd MMM yyyy 23:59")  '^_^20221028 modi by 7643
    End Function
    Public Function GetTourCodeTable(strTourCode) As DataTable
        Dim strQuerry As String = "select t.TourCode, " _
            & " Case When TourType='MICE' then 'MICE_'+t.city else 'IB_SGN' end as TspCust, " _
            & " Case When TourType='MICE' then 'MICE' else 'IB' end as TspClass, SDate " _
            & " from FLX.dbo.TOS_TourCode t join FLX.dbo.tbl_Requests r On t.TourCode=r.TourCode And r.Status='RR'" _
            & " where t.status='OK' and t.city in ('SGN', 'DAD') and t.TourCode='" & strTourCode & "'" _
            & " union " _
            & " select t.TourCode, 'TOURDESK_'+t.city, 'TD', SDate " _
            & " from FLX.dbo.TOS_TourCode t join FLX.dbo.Departure d On t.TourCode=d.Departure And d.Status<>'XX'" _
            & " where t.status='OK' and t.city in ('SGN', 'DAD') and t.TourCode='" & strTourCode & "'"
        Return GetDataTable(strQuerry, conn(1))
    End Function
    'Public Function GetTourCodeTableByRcp(strRcp As String) As DataTable  '^_^20221017 mark by 7643
    Public Function GetTourCodeTableByRcp(strRcp As String, ByRef conn As SqlClient.SqlConnection) As DataTable  '^_^20221017 modi by 7643
        Dim strTourCode As String = GetColumnValuesAsString("FOP", "Document", " where Status='OK' and FOP<>'EXC' and RCPNO='" & strRcp & "'", ",")
        strTourCode = strTourCode.Replace(",", "','")
        Dim strQuerry As String = "Select t.TourCode, " _
            & " Case When TourType='MICE' then 'MICE_'+t.city else 'IB_SGN' end as TspCust, " _
            & " Case When TourType='MICE' then 'MICE' else 'IB' end as TspClass, SDate " _
            & " from FLX.dbo.TOS_TourCode t join FLX.dbo.tbl_Requests r On t.TourCode=r.TourCode And r.Status='RR'" _
            & " where t.status='OK' and t.city in ('SGN', 'DAD','HAN') and t.TourCode in ('" & strTourCode & "')" _
            & " union " _
            & " select t.TourCode, 'TOURDESK_'+t.city, 'TD', SDate " _
            & " from FLX.dbo.TOS_TourCode t join FLX.dbo.Departure d On t.TourCode=d.Departure And d.Status<>'XX'" _
            & " where t.status='OK' and t.city in ('SGN', 'DAD','HAN') and t.TourCode in ('" & strTourCode & "')"
        'Return GetDataTable(strQuerry, conn(1))  '^_^20221017 mark by 7643
        Return GetDataTable(strQuerry, conn)  '^_^20221017 modi by 7643
    End Function
    Public Function GetCustomerAopId(strCustShortName As String) As String
        'Chi dung cho 1 so customer dac biet cua TV
        Select Case strCustShortName
            Case "IB_SGN"
                Return "80000126-1578716238"
            Case "MICE_DAD"
                Return "80000125-1578716204"
            Case "MICE_SGN"
                Return "80000124-1578716162"
            Case "TOURDESK_DAD"
                Return "80000006-1578038251"
            Case "TOURDESK_SGN"
                Return "80000004-1578017737"
            Case Else
                Return ""
        End Select
    End Function
    Public Function GetAopRecordNameTVS(strSrv As String, strGrp As String, strCustShortName As String) As String
        If strSrv = "R" Then
            Return "CreditMemo"
        ElseIf strGrp = "FIT" Then
            Return "Invoice"
        ElseIf strGrp = "FIT" Then
            Return "Invoice"
        ElseIf strCustShortName = "TV_SGN" Then
            Return "Invoice"
        ElseIf strCustShortName = "GDS_SGN" Then
            Return "Invoice"
        Else
            Return "Deposit"
        End If
    End Function
    Public Function GetColumnValuesAsString(strTblName As String, strColumn As String, strCondition As String _
                                       , strSeperator As String) As String
        Dim strResult As String

        strResult = ScalarToString(strTblName, strColumn & "+'" & strSeperator & "'", strCondition & " for xml path('')")
        If strResult <> "" Then
            strResult = Mid(strResult, 1, strResult.Length - strSeperator.Length)
        End If
        Return strResult
    End Function
    Public Function CreateEmail(intCustId, strSubject, strMsg) As Boolean
        Dim strSQL As String = "Insert FT.DBO.Emaillog (CustID, Subj, Msg, Frm, city, Dept) values (" & intCustId _
            & ",'" & strSubject & "','" & strMsg & "','SGN','ALL')"
        ExecuteNonQuerry(strSQL, conn(0))
    End Function
    Public Function Append2TextFile(ByVal strText As String, Optional strLogfile As String = "") As Boolean
        If strLogfile = "" Then
            strLogfile = My.Application.Info.DirectoryPath & "\" _
                                            & Format(Today, "yyMMdd") & pstrPrg & ".txt"
        End If

        Dim objLogFile As New System.IO.StreamWriter(strLogfile, True)
        objLogFile.WriteLine(strText)
        objLogFile.Close()
        objLogFile = Nothing
        Return True
    End Function
    Public Function ExecuteNonQuerry(strQuerry As String, objConn As SqlClient.SqlConnection) As Boolean
        Dim objCmd As SqlClient.SqlCommand = objConn.CreateCommand
        If strQuerry.Contains("AOP_SGN_TVTR") AndAlso objCmd.CommandTimeout < 256 Then
            objCmd.CommandTimeout = 512
        End If
        If strQuerry.Contains("AOP_SGN_TVTR") Then
            Threading.Thread.Sleep(1000)
        End If
        objCmd.CommandText = strQuerry
        If objConn.State = ConnectionState.Closed Then
            objConn.Open()
        End If
        Try
            objCmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Dim strLog As String = vbNewLine & "ERROR|" & Now & vbNewLine & strQuerry & vbNewLine & ex.Message
            If ex.Message.Contains("time out") Then

            End If
            Append2TextFile(strLog)
            Return False
        End Try

    End Function
    Public Function TienBangChu(ByVal sSoTien As String) As String
        Dim DonVi() As String = {"", "nghìn ", "triệu ", "tỷ ", "nghìn ", "triệu "}
        Dim so As String
        Dim chuoi As String = ""
        Dim temp As String
        Dim id As Byte

        If sSoTien = 0 Then
            Return ("Không")
        End If
        sSoTien = sSoTien.Replace(",", "")
        If sSoTien.EndsWith(".00") Then
            sSoTien = Mid(sSoTien, 1, sSoTien.Length - 3)
        End If
        Do While (Not sSoTien.Equals(""))
            If sSoTien.Length <> 0 Then
                so = getNum(sSoTien)
                sSoTien = Strings.Left(sSoTien, sSoTien.Length - so.Length)
                temp = setNum(so)
                so = temp
                If Not so.Equals("") Then
                    temp = temp + DonVi(id)
                    chuoi = temp + chuoi
                End If
                id = id + 1
            End If
        Loop
        temp = UCase(Strings.Left(chuoi, 1))

        Return temp & Strings.Right(chuoi, Len(chuoi) - 1)
    End Function
    Private Function getNum(ByVal sSoTien As String) As String
        Dim so As String

        If sSoTien.Length >= 3 Then
            so = Strings.Right(sSoTien, 3)
        Else
            so = Strings.Right(sSoTien, sSoTien.Length)
        End If
        Return so
    End Function
    Private Function setNum(ByVal sSoTien As String) As String
        Dim chuoi As String = ""
        Dim flag0 As Boolean
        Dim flag1 As Boolean
        Dim temp As String

        temp = sSoTien
        Dim kyso() As String = {"không ", "một ", "hai ", "ba ", "bốn ", "năm ", "sáu ", "bảy ", "tám ", "chín "}
        'Xet hang tram
        If sSoTien.Length = 3 Then
            If Not (Strings.Left(sSoTien, 1) = 0 And Strings.Left(Strings.Right(sSoTien, 2), 1) = 0 And Strings.Right(sSoTien, 1) = 0) Then
                chuoi = kyso(Strings.Left(sSoTien, 1)) + "trăm "
            End If
            sSoTien = Strings.Right(sSoTien, 2)
        End If
        'Xet hang chuc
        If sSoTien.Length = 2 Then
            If Strings.Left(sSoTien, 1) = 0 Then
                If Strings.Right(sSoTien, 1) <> 0 Then
                    chuoi = chuoi + "linh "
                End If
                flag0 = True
            Else
                If Strings.Left(sSoTien, 1) = 1 Then
                    chuoi = chuoi + "mười "
                Else
                    chuoi = chuoi + kyso(Strings.Left(sSoTien, 1)) + "mươi "
                    flag1 = True
                End If
            End If
            sSoTien = Strings.Right(sSoTien, 1)
        End If
        'Xet hang don vi
        If Strings.Right(sSoTien, 1) <> 0 Then
            If Strings.Left(sSoTien, 1) = 5 And Not flag0 Then
                If temp.Length = 1 Then
                    chuoi = chuoi + "năm "
                Else
                    chuoi = chuoi + "lăm "
                End If
            Else
                If Strings.Left(sSoTien, 1) = 1 And Not (Not flag1 Or flag0) And chuoi <> "" Then
                    chuoi = chuoi + "mốt "
                Else
                    chuoi = chuoi + kyso(Strings.Left(sSoTien, 1)) + ""
                End If
            End If
        Else
        End If
        Return chuoi
    End Function
    Public Function DefineDomInt(strRtg As String, strDomCities As String) As String
        Dim arrCities() As String
        Dim i As Integer
        Dim strRtgType As String = "DOM"
        arrCities = Split(strRtg, " ")

        For i = 0 To arrCities.Length - 1 Step 2
            If Not strDomCities.Contains(arrCities(i)) Then
                Return "INT"
            End If
        Next

        Return strRtgType
    End Function
    Public Function GetTaxAmtFromTaxDetails(strTaxCode As String, strTaxDetails As String) As Decimal
        Dim decResult As Decimal = 0

        If strTaxDetails <> "" Then
            Dim arrTaxes() As String = strTaxDetails.Split("|")
            Dim i As Integer
            For i = 0 To arrTaxes.Length - 1
                If Mid(arrTaxes(i), 1, 2) = strTaxCode Then
                    decResult = decResult + Mid(arrTaxes(i), 3)
                End If
            Next
        End If
        Return decResult
    End Function
End Module
