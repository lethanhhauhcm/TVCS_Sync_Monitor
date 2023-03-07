Imports TVCS.Crd_Ctrl
Imports TVCS.MySharedFunctionsWzConn
Module AutoInsertRAS
    Private LocalRCPID As Integer
    Private MyCust As New objCustomer
    Private myAutoTKT As New objAutoTKT
    Private DKAutoEligible As String, DKInsert As String, nVeDuoi As String
    Dim cmd As SqlClient.SqlCommand = conn(0).CreateCommand
    Private Function isBSPStock(ByVal pTKNO As String) As Boolean
        cmd.CommandText = "select RecID from MISC where VAL='" & pTKNO.Substring(4, 4) & "' and cat='BSPSTOCK'"
        Dim i As Integer = cmd.ExecuteScalar
        Return IIf(i > 0, True, False)
    End Function
    Private Function GetDKAuto(pCounter As String) As String
        'Dim frmDate As String = Format(DateAdd(DateInterval.Day, -7, Now.Date), "dd-MMM-yy")  '^_^20221028 mark by 7643
        Dim frmDate As String = Format(DateAdd(DateInterval.Day, -7, Now.Date), "dd-MMM-yyyy")  '^_^20221028 modi by 7643
        Dim FstUpdate As Date = Now.AddMinutes(-4)
        Dim KQ As String
        DKInsert = " and tkno not in (select tkno from tkt where statusAL ='OK' and DATEDIFF(d,DOI,GETDATE())<365)"

        '^_^20221028 mark by 7643 -b-
        'KQ = " autoras=1 and custID>0 and fstupdate <'" & Format(FstUpdate, "dd-MMM-yy HH:mm") & "'" &
        '    " and RLOC<>'' and qty <>0 and srv='S' and doi >'" & frmDate & "'" &
        '    " and counter='" & pCounter & "' and tkno<>'' " & DKInsert &
        '    " and custid in (select custID from cust_detail where status='OK' and cat='Channel' "
        '^_^20221028 mark by 7643 -e-
        '^_^20221028 modi by 7643 -b-
        KQ = " autoras=1 and custID>0 and fstupdate <'" & Format(FstUpdate, "dd-MMM-yyyy HH:mm") & "'" &
            " and RLOC<>'' and qty <>0 and srv='S' and doi >'" & frmDate & "'" &
            " and counter='" & pCounter & "' and tkno<>'' " & DKInsert &
            " and custid in (select custID from cust_detail where status='OK' and cat='Channel' "
        '^_^20221028 modi by 7643 -e-
        If pCounter = "CWT" Then
            KQ = KQ & " and val in ('CS','LC')) and left(FOP,3) in ('INV','CC/','MCE') " & _
                " and PRG='" & pCounter & "' "
        Else
            KQ = KQ & " and val not in ('CS','LC')) and DOI > '31 Dec 2016 23:59'"
            'And left(FOP,3) in ('INV') "
        End If
        Return KQ
    End Function
    Public Sub AutoGetCapturedTKT2RAS_Master()
        pstrVnDomCities = GetColumnValuesAsString("CityCode", "Airport", "where Country='VN'", "_")
        Dim tblTkt1a As DataTable
        Dim k As Integer
        Dim strDomInt As String
        For k = 0 To 1
            tblTkt1a = GetDataTable("select top 64 * from tkt_1a where doctype='ETK'" _
                                    & " And DomInt='' order by recid desc" _
                                    , conn(k))
            For Each objRow As DataRow In tblTkt1a.Rows
                strDomInt = DefineDomInt(objRow("FullRtg"), pstrVnDomCities)
                ExecuteNonQuerry("Update tkt_1a set DomInt='" & strDomInt _
                                 & "' where RecId=" & objRow("RecId"), conn(k))
            Next
        Next

        Try
            Dim TKT1A_RecID As Integer, KQAutoRas As String
            Dim StrCounter As String = "CWT_TVS"
            Dim arrCounter As String() = Split(StrCounter, "_")

            For c As Int16 = 0 To 1
                cmd = conn(c).CreateCommand()
                cmd.CommandText = " select count(*) from RCP where status='QQ' and fstUser='AUT' and FstUpdate>'9-Mar-16'"
                If cmd.ExecuteScalar > 0 Then
                    Continue For
                End If

                DKAutoEligible = GetDKAuto(arrCounter(c))
                For i As Int16 = 1 To 16
                    cmd.CommandText = "select top 1 RecID from tkt_1a where " & DKAutoEligible & " order by RecID desc"
                    TKT1A_RecID = cmd.ExecuteScalar
                    If TKT1A_RecID = 0 Then Exit For
                    myAutoTKT.SetTkid(TKT1A_RecID, conn(c))

                    If myAutoTKT.isNormalRecord And myAutoTKT.AutoBy <> "NILL" Then
                        KQAutoRas = AutoGetCapturedTKT2RAS(conn(c))
                        If InStr("PSP_PPD", myAutoTKT.FOP) > 0 And KQAutoRas <> "" Then

                            Call defineVND_Avail(MyCust.CustID, myAutoTKT.FOP, IIf(myAutoTKT.FOP = "PPD", "01-jan-12", "01-jan-12"),
                                                 MyCust.LstReconcile, "After " & "AUTORAS " & myAutoTKT.TKNO & KQAutoRas, conn(0), "AUT", connStrRAS(0))
                        End If

                        'Create AopQueue
                        If KQAutoRas <> "" Then
                            Threading.Thread.Sleep(5000)
                            'If CreateAopQueueAir(KQAutoRas) Then  '^_^20221017 mark by 7643
                            If CreateAopQueueAir(KQAutoRas, conn(c)) Then  '^_^20221017 modi by 7643
                                'MsgBox("AopQueue created!")
                            End If
                        End If

                    Else
                        UpdateAutoRasStatus(myAutoTKT.RLOC, 4)
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try

    End Sub
    Private Sub UpdateAutoRasStatus(pRLOC As String, pStatus As Int16)
        cmd.CommandText = "update TKT_1a set autoras=" & pStatus & ", lstUpdate=getdate() where rloc='" & pRLOC & "'"
        cmd.ExecuteNonQuery()
    End Sub
    Public Function AutoGetCapturedTKT2RAS(ByRef conn As SqlClient.SqlConnection) As String
        Dim RCPNo As String, INVNo As String, VNDAvail As Decimal

        If myAutoTKT.FOP = "MCE" AndAlso InvalidTourCode(myAutoTKT.DocNo, myAutoTKT.CustID, myAutoTKT.SRV, myAutoTKT.TKNO, True, myAutoTKT.DOI) Then
            UpdateAutoRasStatus(myAutoTKT.RLOC, 6)
            Return ""
        End If

        If myAutoTKT.Curr = "" Or myAutoTKT.ROE = 0 Or myAutoTKT.CustID = 0 Then
            UpdateAutoRasStatus(myAutoTKT.RLOC, 7)
            Return ""
        End If
        MyCust.SetCustId(myAutoTKT.CustID, conn)

        'If InStr("PSP_PPD", myAutoTKT.FOP) > 0 Then  '^_^20221103 mark by 7643
        If InStr("PSP_PPD", myAutoTKT.FOP) > 0 And myAutoTKT.FOP <> "" And myAutoTKT.FOP IsNot Nothing Then  '^_^20221103 modi by 7643
            VNDAvail = defineVND_Avail(MyCust.CustID, myAutoTKT.FOP, IIf(myAutoTKT.FOP = "PPD", "01-jan-12", "01-jan-12"),
                                       MyCust.LstReconcile, "B4 " & "AUTORAS " & myAutoTKT.TKNO, conn, "AUT", connStrRAS(0))
            If VNDAvail < myAutoTKT.RoughAmt Then
                UpdateAutoRasStatus(myAutoTKT.RLOC, 5)
                Return ""
            End If

        End If

        nVeDuoi = ""
        If myAutoTKT.FTKT <> "" AndAlso myAutoTKT.AutoBy = "TKID" Then
            nVeDuoi = XacDinhVenoi(myAutoTKT.TKNO, myAutoTKT.FTKT, 0)
        End If

        Dim strLocation As String = IIf(myAutoTKT.Location = "", "TVH", myAutoTKT.Location)
        Dim strCity As String
        Dim strPOS As String
        Dim RCPInfor(2) As String

        If conn.ConnectionString.ToUpper.Contains("HAN") Then
            strCity = "HAN"
            strPOS = "3"
        Else
            strCity = "SGN"
            strPOS = "0"
        End If

        RCPInfor = GenRCPNo_ID("TS", strPOS, myAutoTKT.TKT_1A_ID, conn)
        LocalRCPID = RCPInfor(0)
        RCPNo = RCPInfor(1)
        If RCPNo = "" Or LocalRCPID = 0 Then Return ""

        '^_^20221028 mark by 7643 -b-
        'cmd.CommandText = " Update dbo.RCP set CustID=" & MyCust.CustID &
        '    ", FstUser='AUT', CustType='" & MyCust.CustType & "', Counter='" & myAutoTKT.Counter & "'" &
        '    ", status='OK', SRV='S', DOS='" & Format(Now, "dd-MMM-yy") & "', CA='" & myAutoTKT.Booker.Replace("--", "") & "'" &
        '    ", Stock='" & myAutoTKT.AL & "', CustshortName='" & MyCust.ShortName _
        '    & "', PrintedCustName=N'" & MyCust.FullName.Replace("--", "") _
        '    & "', PrintedCustAddrr=N'" & MyCust.Addr.Replace("--", "") _
        '    & "', PrintedTaxCode='" & MyCust.taxCode.Replace("--", "") _
        '    & "', Currency='" & myAutoTKT.Curr & "', ROE=" & myAutoTKT.ROE _
        '    & ", City='" & strCity & "', Location='" & strLocation & "', TTLDue=0, ROEID=" &
        '    myAutoTKT.ROEID & ", RMK='" & myAutoTKT.LogTxt & "',Vendor='" & myAutoTKT.Vendor & "',VendorId=" & myAutoTKT.VendorId _
        '    & "  where recid=" & LocalRCPID
        '^_^20221028 mark by 7643 -e-
        '^_^20221028 modi by 7643 -b-
        cmd.CommandText = " Update dbo.RCP set CustID=" & MyCust.CustID &
            ", FstUser='AUT', CustType='" & MyCust.CustType & "', Counter='" & myAutoTKT.Counter & "'" &
            ", status='OK', SRV='S', DOS='" & Format(Now, "dd-MMM-yyyy") & "', CA='" & myAutoTKT.Booker.Replace("--", "") & "'" &
            ", Stock='" & myAutoTKT.AL & "', CustshortName='" & MyCust.ShortName _
            & "', PrintedCustName=N'" & MyCust.FullName.Replace("--", "") _
            & "', PrintedCustAddrr=N'" & MyCust.Addr.Replace("--", "") _
            & "', PrintedTaxCode='" & MyCust.taxCode.Replace("--", "") _
            & "', Currency='" & myAutoTKT.Curr & "', ROE=" & myAutoTKT.ROE _
            & ", City='" & strCity & "', Location='" & strLocation & "', TTLDue=0, ROEID=" &
            myAutoTKT.ROEID & ", RMK='" & myAutoTKT.LogTxt & "',Vendor='" & myAutoTKT.Vendor _
            & "',VendorId=" & myAutoTKT.VendorId & ",RptData1=N'" & myAutoTKT.RptData1 _
            & "',RptData2=N'" & myAutoTKT.RptData2 & "',RptData3=N'" & myAutoTKT.RptData3 _
            & "'  where recid=" & LocalRCPID
        '^_^20221028 modi by 7643 -e-

        cmd.ExecuteNonQuery()

        If CreateRCP_Record(RCPNo, myAutoTKT.TKT_1A_ID, conn) Then
            If isBSPStock(myAutoTKT.TKNO) Then
                'INVNo = GenInvNo_QD153(RCPNo, "AH")
                'Dim InvID As Integer = AutoRAS_TaoBanGhiHoaDon(RCPNo, INVNo)
                'TaoBanGhiTKTNO_INVNO_Standard(LocalRCPID, INVNo, InvID)

                ''tao hoa don dien tu cho VYM , KIWI SRO
                Select Case MyCust.ShortName
                    Case "KIWISRO", "VYM", "AERTICKET"
                        If Not myAutoTKT.TKNO.StartsWith("978") Then
                            If Now < CDate("01 Jun 22") Then
                                'AutoE_Invoice(RCPNo, True) Tam bo phan nay vi chua du vi du
                            End If
                        End If
                End Select
            End If
        Else
            cmd.CommandText = "delete from rcp where recid=" & LocalRCPID &
                "; Update tkt_1A Set autoras=3 where RLOC='" & myAutoTKT.RLOC & "'"
            cmd.ExecuteNonQuery()
        End If
        Return RCPNo
    End Function
    Private Function AutoRAS_TaoBanGhiHoaDon(ByVal parRCPNO As String, ByVal INVNO As String) As Integer
        Dim InvID As Integer, strSQL As String
        InvID = Insert_INV("E", INVNO, INVNO.Substring(0, 2), LocalRCPID)
        cmd.CommandText = "select SRV from rcp where recid=" & LocalRCPID
        Dim SRV As String = cmd.ExecuteScalar
        Dim dtblTVTR As DataTable = GetDataTable("select VAL, VAL1, VAL2 from MISC where cat='TVCompany' and description='TVT'", conn(0))

        strSQL = "update INV set "
        strSQL = strSQL & "SRV='" & SRV & "', "
        strSQL = strSQL & "CustID=" & MyCust.CustID & ", "
        strSQL = strSQL & "CustShortName='" & MyCust.ShortName & "', "
        If MyCust.CustType = "CS" Or MyCust.CustType = "LC" Then
            strSQL = strSQL & "CustFullName=N'" & dtblTVTR.Rows(0)("VAL") & "', "
            strSQL = strSQL & "CustAddress=N'" & dtblTVTR.Rows(0)("VAL1") & "', "
            strSQL = strSQL & "CustTaxCode='" & dtblTVTR.Rows(0)("VAL2") & "', "
        Else
            strSQL = strSQL & "CustFullName=N'" & MyCust.FullName & "', "
            strSQL = strSQL & "CustAddress=N'" & MyCust.Addr & "', "
            strSQL = strSQL & "CustTaxCode='" & MyCust.taxCode & "', "
        End If
        strSQL = strSQL & "Amount=" & myAutoTKT.TTLFTC * myAutoTKT.ROE & ", "
        strSQL = strSQL & "FstUser='AUT', FOP='" & myAutoTKT.FOPDetail & "' Where RecID=" & InvID
        cmd.CommandText = strSQL
        cmd.ExecuteNonQuery()
        Return InvID
    End Function

    Private Function CreateRCP_Record(ByVal pRCPNO As String, ByVal pTKT1AID As Integer _
                                      , ByRef conn As SqlClient.SqlConnection) As Boolean
        Dim t As SqlClient.SqlTransaction = conn.BeginTransaction
        cmd.Transaction = t
        Try
            TaoBanGhiTKT(pRCPNO, t)
            cmd.CommandText = "select sum(fare+tax+charge+chargeTV-commval) from tkt where rcpid=" & LocalRCPID
            Dim ChKTTL As Decimal = cmd.ExecuteScalar ' chk la da tao duoc TTK ok
            If ChKTTL = 0 Then
                t.Rollback()
                Return False
            End If
            cmd.CommandText = "update RCP set TTLDue=@TTLDue, Charge=@Charge, Charge_d=@Charge_d where recID=@RecID"
            cmd.Parameters.Clear()
            cmd.Parameters.Add("@TTLDue", SqlDbType.Decimal).Value = myAutoTKT.TTLFTC + myAutoTKT.SFinFareCurr + myAutoTKT.MerchantFee
            cmd.Parameters.Add("@Charge", SqlDbType.Decimal).Value = myAutoTKT.MerchantFee
            cmd.Parameters.Add("@Charge_d", SqlDbType.VarChar).Value = IIf(myAutoTKT.MerchantFee = 0, "", "CRD:" & myAutoTKT.MerchantFee.ToString)
            cmd.Parameters.Add("@RecID", SqlDbType.Int).Value = LocalRCPID
            cmd.ExecuteNonQuery()

            TaoBanGhiFOP(pRCPNO, myAutoTKT.FOPDetail, t)

            cmd.CommandText = "update tkt_1A set autoras=2, TinhTrang='RE' where "
            If myAutoTKT.AutoBy = "TKID" Then
                cmd.CommandText = cmd.CommandText & " recid=" & myAutoTKT.TKT_1A_ID
                If nVeDuoi <> "" Then
                    cmd.CommandText = cmd.CommandText & " or tkno in " & nVeDuoi
                End If
            Else
                cmd.CommandText = cmd.CommandText & " RLOC='" & myAutoTKT.RLOC & "'"
            End If
            cmd.ExecuteNonQuery() ' de test/edu thi ko update
            t.Commit()
            Return True
        Catch ex As Exception
            t.Rollback()
            Return False
        End Try
    End Function
    Private Function XacDinhVenoi(ByVal pTKNO As String, ByVal pFTKT As String, intConxId As Integer) As String
        Dim KQ As String = "", VeI As String, veI_1 As String, tmpConj As String
        If pFTKT.Length < 2 Then Return ""
        veI_1 = pTKNO
        For i As Int16 = 2 To 4
            VeI = veI_1.Substring(0, 10) & Format(CLng(veI_1.Substring(10, 5)) + 1, "00000")
            KQ = KQ & "," & VeI
            cmd.CommandText = "select FTKT from " & arrTvcsDb(intConxId) & ".dbo.tkt_1A where status<>'XX' and TKNO='" & VeI & "'"
            tmpConj = cmd.ExecuteScalar
            If tmpConj.Trim.Length = 3 Then Exit For
            veI_1 = VeI
        Next
        If KQ.Length > 2 Then
            KQ = KQ.Substring(1)
            KQ = KQ.Replace(",", "','")
            KQ = "('" & KQ & "')"
        End If
        Return KQ
    End Function
    Private Sub TaoBanGhiTKT(ByVal pRCPNO As String, ByVal pt As SqlClient.SqlTransaction)
        cmd.Transaction = pt
        Dim decVatAmt2AL As Decimal

        If myAutoTKT.AutoBy = "RLOC" Then
            cmd.CommandText = "insert tkt (RCPID, RCPNO, Currency, AL" _
                & ", TKNO, FTKT, SRV, Qty, DOI, DOF, BkgClass" _
                & ", PaxName, paxType,TourCode, Fare, ShownFare, Tax,Tax2AL, Charge, ChargeTV, CommVAL" _
                & ", NetToAL, Itinerary, FareBasis, Charge_D, RMK " _
                & ", StockCtrl, FstUser, LstUser, DocType, Tax_D,Booker,Rloc,DomInt,ReturnDate,Email,VatInfoId,TktIssuedBy)" _
                & " select  " & LocalRCPID & ",'" & pRCPNO & "',currency,'" & myAutoTKT.AL _
                & "', TKNO, FTKT,srv, qty, doi, dof, bkgclass, paxname, paxType," _
                & " tourcode, qty * fare, shownfare, qty*tax, qty*tax, qty*Charge,qty*(svcFee+MU)" _
                & myAutoTKT.Formula_SF_ROE & ", commVAL, qty*nettoAL, fullRTG, farebasis, " _
                & " ChargeDetail,'" & myAutoTKT.Booker &
                "', left(stockctrl,13),'" & myAutoTKT.FstUser _
                & "', 'AUT', 'ETK', taxdetail,Booker,Rloc,DomInt,ReturnDate,Email,VatInfoId,TktIssuedBy" _
                & " from tkt_1A where  RLOC='" & myAutoTKT.RLOC & "'" & DKInsert
            cmd.ExecuteNonQuery()
        Else
            decVatAmt2AL = GetTaxAmtFromTaxDetails("UE", myAutoTKT.TaxDetail)
            cmd.CommandText = "insert tkt (RCPID, RCPNO, Currency, AL" _
                & ", TKNO, FTKT, SRV, Qty, DOI, DOF, BkgClass" _
                & ", PaxName, paxType,TourCode, Fare, ShownFare" _
                & ", Tax,Tax2AL,VatAmt2AL, Charge, ChargeTV, CommVAL, NetToAL" _
                & ", Itinerary, FareBasis, Charge_D, RMK, " _
                & " StockCtrl, FstUser, LstUser, DocType, Tax_D,Booker,Rloc,DomInt,ReturnDate,Email,VatInfoId,TktIssuedBy)" _
                & " select  " & LocalRCPID & ",'" & pRCPNO & "',currency,'" & myAutoTKT.AL _
                & "', TKNO, FTKT,srv, qty, doi, dof, bkgclass, paxname, paxType," _
                & " tourcode, qty * fare, shownfare, qty*tax, qty*tax," & decVatAmt2AL _
                & ",qty*Charge, qty*(svcFee+MU)" _
                & myAutoTKT.Formula_SF_ROE & ", commVAL, qty*nettoAL, fullRTG, farebasis, " _
                & " ChargeDetail,'" & myAutoTKT.Booker & "', left(stockctrl,13),'" & myAutoTKT.FstUser _
                & "','AUT', 'ETK', taxdetail,Booker,Rloc,DomInt,ReturnDate,Email,VatInfoId,TktIssuedBy" _
                & " from tkt_1A where recID=" _
                & myAutoTKT.TKT_1A_ID & DKInsert
            cmd.ExecuteNonQuery()
            If nVeDuoi <> "" Then
                cmd.CommandText = "insert tkt (RCPID, RCPNO, Currency, AL, TKNO, FTKT, SRV, Qty, DOI, DOF, BkgClass, PaxName, paxType," &
                    " Itinerary, FareBasis, FstUser, LstUser, DocType,Tax_D,Booker,Rloc,DomInt,ReturnDate)" _
                    & " Select " & LocalRCPID & ",'" & pRCPNO & "',currency,'" _
                    & myAutoTKT.AL & "', TKNO, FTKT, srv,0 , doi, dof, bkgclass" _
                    & ", paxname, paxType, fullRTG, farebasis,'" _
                    & myAutoTKT.FstUser & "','AUT', 'ETK', taxdetail,Booker,Rloc,DomInt,ReturnDate " _
                    & " from tkt_1A where  TKNO in " & nVeDuoi & DKInsert
                cmd.ExecuteNonQuery()
            End If
        End If
    End Sub
    Private Sub TaoBanGhiFOP(ByVal ppRCPNO As String, ByVal pFOPDetail As String, ByVal pt As SqlClient.SqlTransaction)
        Dim vStatus As String = "OK"

        If pFOPDetail.Split("|")(0).Length < 8 Then Exit Sub  '^_^20221103 add by 7643

        If pFOPDetail.Split("|")(0).Substring(0, 3) = "CRD" And
            (myAutoTKT.Counter = "CWT" Or myAutoTKT.CustID = 57641) Then
            'ngoai le cho Asia desk
            vStatus = "QQ"
        End If
        cmd.Transaction = pt
        cmd.CommandText = "insert fop (fop, currency, roe, amount, RCPID, RCPNO, Document, customerID, FstUser, Status, LstUser) values ('" &
            pFOPDetail.Split("|")(0).Substring(0, 3) & "','" & pFOPDetail.Split("|")(0).Substring(4, 3) & "'," &
            IIf(pFOPDetail.Split("|")(0).Substring(4, 3) = "VND", 1, myAutoTKT.ROE) & "," &
            pFOPDetail.Split("|")(0).Substring(8) & "," & LocalRCPID & ",'" & ppRCPNO & "','" & myAutoTKT.DocNo & "'," &
            myAutoTKT.CustID & ",'" & myAutoTKT.FstUser & "','" & vStatus & "','AUT')"
        For i As Int16 = 1 To UBound(pFOPDetail.Split("|"))
            cmd.CommandText = cmd.CommandText & "; insert fop (fop, currency, roe, amount, RCPID, RCPNO, CustomerID, FstUser, Status, LstUser) " &
                "values ('" & pFOPDetail.Split("|")(i).Substring(0, 3) & "','" & pFOPDetail.Split("|")(i).Substring(4, 3) & "'," &
                IIf(pFOPDetail.Split("|")(i).Substring(4, 3) = "VND", 1, myAutoTKT.ROE) & "," &
                pFOPDetail.Split("|")(i).Substring(8) & "," & LocalRCPID & ",'" & ppRCPNO & "'," & myAutoTKT.CustID & ",'" &
                myAutoTKT.FstUser & "','" & vStatus & "','AUT')"
        Next
        cmd.ExecuteNonQuery()
    End Sub
    Public Function AutoE_Invoice(strRcpNo As String, blnIssued2TV As Boolean) As Boolean
        Dim tblCust As System.Data.DataTable
        Dim decTax As Decimal
        Dim decCharge As Decimal
        Dim decInvTotal As Decimal
        Dim intRefundMultiplier As Integer = 1
        Dim strProduct As String
        Dim strSrv As String
        Dim strKindOfService As String

        Dim strLoadTkt As String
        Dim tblTkts As DataTable

        Dim intCustId As Integer
        Dim strCustShortName As String
        Dim strCustFullName As String
        Dim strCustAddress As String
        Dim strTaxCode As String
        Dim strEmail As String
        Dim lstProducts As New List(Of clsProduct)
        Dim intInvId As Integer
        Dim objOriInv As DataRow
        Dim intAdjustType As Integer = 0
        Dim objE_Inv As New clsE_Invoice
        Dim objConnect As New clsE_InvConnect(True, "TVTR")
        Dim intNewRecId As Integer
        Dim strOriFkey As String = ""
        Dim dteOldInvDOI As Date

        pblnTT78 = True

        Dim tblE_InvSettings As DataTable = GetDataTable("Select * From lib.dbo.E_InvSettings78 " _
                                            & " Where TVC ='TVTR' and Status='OK' and Biz='PAX' and AL='TS'", conn(0))
        Dim strMauSo As String
        Dim strKyHieu As String

        strMauSo = tblE_InvSettings.Rows(0)("MauSo")
        strKyHieu = CreateKyHieu(tblE_InvSettings.Rows(0)("KyHieu"))

        tblCust = GetDataTable("select c.*, r.Srv from lib.dbo.Customer c" _
                            & " left join Rcp r on c.RecId=r.CustId" _
                            & " where r.Status='OK'" _
                            & " And r.RcpNo='" & strRcpNo & "'", conn(0))
        If tblCust.Rows.Count = 0 Then
            Return False
        End If

        intCustId = tblCust.Rows(0)("RecId")
        strCustShortName = tblCust.Rows(0)("CustShortName")
        strCustFullName = tblCust.Rows(0)("CustFullName")
        strCustAddress = tblCust.Rows(0)("CustAddress")
        strTaxCode = tblCust.Rows(0)("CustTaxCode")
        strEmail = ""

        strProduct = "AIR"

        strSrv = tblCust.Rows(0)("SRV")
        strKindOfService = KindOfService.Hóa_đơn_GTGT

        strLoadTkt = "Select t.RecId,t.RcpId,t.Tkno,t.itinerary as Rtg" _
            & ", ((t.Fare+t.Tax)*qty+t.Charge+t.ChargeTV)*ROE as Total,t.StockCtrl" _
            & " from tkt t left join rcp r on r.recid=t.rcpid" _
            & " where Qty<>0 And t.Status<>'xx' and t.Rcpno='" & strRcpNo & "'"
        tblTkts = GetDataTable(strLoadTkt, conn(0))

        If strSrv = "R" Then
            intRefundMultiplier = -1
        End If

        If tblCust.Rows(0)("SRV") = "R" Or tblTkts.Rows(0)("StockCtrl") <> "" Then

            Dim strOldTkno As String
            If strSrv = "R" Then
                intAdjustType = 3
                strOldTkno = tblTkts.Rows(0)("Tkno")
            Else
                Dim tblOldTkts As DataTable
                intAdjustType = 2

                tblOldTkts = GetDataTable("select top 1 RcpId,RecId,Tkno,StockCtrl from Tkt where SRV='S' and replace(Tkno,' ','')='" _
                                        & Replace(tblTkts.Rows(0)("StockCtrl"), " ", "") _
                                        & "' order by RecId desc", conn(0))
                If tblOldTkts.Rows.Count = 0 Then
                    CreateEmail(-98, "unable to find Original Ticket for " & tblTkts.Rows(0)("Tkno") _
                        , "Please ask accounting dept to issue invoice for this tkt!")
                    Return False
                Else
                    strOldTkno = tblOldTkts.Rows(0)("Tkno")
                End If
            End If
            objOriInv = GetOriginalInv(True, strOldTkno)
            If objOriInv IsNot Nothing Then
                dteOldInvDOI = objOriInv("DOI")
            Else
                objOriInv = GetOriginalInv(False, strOldTkno)
                If objOriInv IsNot Nothing Then
                    dteOldInvDOI = objOriInv("DOI")
                Else
                    CreateEmail(-98, "unable to find Original Invoice for " & tblTkts.Rows(0)("Tkno") _
                            , "Please ask accounting dept to issue invoice for this tkt!")
                    Return False
                End If
            End If

        End If

        For Each objRow As DataRow In tblTkts.Rows
            Dim objProduct As New clsProduct
            With objRow
                objProduct.ProdQuantity = 1
                objProduct.ProdName = objRow("Tkno")
                If tblCust.Rows(0)("SRV") = "R" Then
                    objProduct.ProdName = "Hoàn vé " & objProduct.ProdName
                End If
                objProduct.Extra1 = objRow("Rtg")
                objProduct.ProdPrice = objRow("Total")
                objProduct.Amount = objRow("Total")
                objProduct.VatAmount = 0
                objProduct.VatRate = 0
                objProduct.TotalPrice = objRow("Total")
                objProduct.IsSum = 0
                objProduct.ProdUnit = "Vé"
                objProduct.ProdQuantity = 1
                lstProducts.Add(objProduct)
                decInvTotal = decInvTotal + objRow("Total")
            End With
        Next

        If objOriInv IsNot Nothing Then
            strOriFkey = objOriInv("InvId")
        End If
        intInvId = InsertE_Inv(intAdjustType, tblTkts, "TVTR", "PAX", "", strSrv, intCustId, strCustShortName, strCustFullName, strCustAddress _
                               , strTaxCode, 0, 0, strMauSo, strKyHieu, "", intNewRecId, strOriFkey)

        Try
            Dim blnSuccess As Boolean = True
            If intAdjustType = 0 Then
                If Not objE_Inv.ImportAndPublishInv(objConnect.WsUrl, objConnect.UserName, objConnect.UserPass, objConnect.AccountName, objConnect.AccountPass _
                                            , intCustId, strCustFullName, strCustAddress, "", strTaxCode, "CK", intInvId, strKindOfService _
                                            , 0, lstProducts, strMauSo, strKyHieu, decTax, decCharge,, strEmail) Then
                    'If Not objE_Inv.ImportAndPublishInv(objConnect.WsUrl, objConnect.UserName, objConnect.UserPass, objConnect.AccountName, objConnect.AccountPass _
                    '                            , intCustId, strCustFullName, strCustAddress, "", strTaxCode, "CK", intInvId, strKindOfService _
                    '                            , 0, lstProducts, strMauSo, strKyHieu, decTax, decCharge,, strEmail) Then
                    InsertEmail(-98, "Unable to create E Invoice for " & strRcpNo, "Please create manually", "SGN", "TVS")
                    blnSuccess = False
                End If
                If blnSuccess AndAlso objE_Inv.ReponseCode.StartsWith("OK") Then
                    Dim arrBreaks As String() = objE_Inv.ReponseCode.Split("-")
                    Dim arrMauSoKyHieu As String() = arrBreaks(0).Split(";")
                    Dim arrKeyNbr As String() = arrBreaks(1).Split("_")
                    Dim intInvNo As Integer = 0
                    If arrKeyNbr.Length = 2 Then
                        intInvNo = arrKeyNbr(1)
                    End If

                    Dim strQuerry As String = "Update lib.dbo.E_Inv78 set MauSo='" & Mid(arrMauSoKyHieu(0), 4) _
                        & "',KyHieu='" & arrMauSoKyHieu(1) & "',InvoiceNo=" & intInvNo _
                        & ",DOI=getdate(),Draft='FALSE' where Recid=" & intNewRecId

                    If ExecuteNonQuerry(strQuerry, conn(1)) Then
                        InsertEmail(-98, "E Invoice for " & strRcpNo, "Please CHECK to verify it is correct", "SGN", "TVS")
                        Return True

                    Else
                        MsgBox("Unable to update E Invoice into RAS Database for InvId:" & intInvId _
                               & vbNewLine & objE_Inv.ResponseDesc _
                               & vbNewLine & ". Please report NMK!")
                        Return False
                    End If
                Else
                    Return False
                End If

                InsertEmail(-89, "E Invoice for " & strRcpNo, "Please CHECK to verify it is correct", "SGN", "TVS")
            Else
                If Not AdjustInvoice("TVTR", intNewRecId, intInvId, intAdjustType, intCustId _
                    , strCustFullName, strCustAddress _
                    , strTaxCode, strEmail, "CK", strMauSo, strKyHieu, objOriInv("InvId") _
                    , objOriInv("InvoiceNo"), objOriInv("MauSo"), objOriInv("KyHieu"), dteOldInvDOI) Then

                    InsertEmail(-98, "Unable to create E Invoice for " & strRcpNo _
                                , "Please create manually", "SGN", "TVS")
                    blnSuccess = False
                    Return False
                Else
                    InsertEmail(-98, "E Invoice for " & strRcpNo, "Please CHECK to verify it is correct", "SGN", "TVS")
                    Return True
                End If
            End If

        Catch ex As Exception

        End Try

    End Function
    Private Function InsertE_Inv(intAdjustType As Integer, tblTkts As DataTable, strTvc As String, strBiz As String, strAL As String, strSrv As String, strCustId As String _
                                 , strCustShortName As String, strCustFullName As String, strAddress As String _
                                 , strTaxCode As String, decTax As Decimal, decCharge As Decimal _
                                 , strMauSo As String, strKyHieu As String, strEmail As String _
                                 , ByRef intNewInvRecId As Integer, Optional strOriFkey As String = "") As Integer
        Dim lstQuerry As New List(Of String)
        Dim strQuerry As String
        Dim strFields As String
        Dim strValues As String
        Dim strInvId As String
        Dim intNewInvId As Integer
        Dim blnNoOriInv As Boolean
        Dim strTkno As String

        If IsNumeric(strOriFkey) AndAlso CInt(strOriFkey) < 40000 Then
            blnNoOriInv = True
        End If
        lstQuerry.Clear()
        strInvId = "(select isnull(Max(InvId),0)+1 from lib.dbo.E_Inv78)"

        strFields = "insert into lib.dbo.E_Inv78 (TVC,Biz,AL,Srv,CustId,CustShortName" _
                    & ", InvID,InvoiceNo, CustFullName, Address, TaxCode" _
                    & ", Period, Status, FstUser, City,Tax,Charge,MauSo,KyHieu,Buyer" _
                    & ",Email,Booker,BU,DomInt,CodeTour,NbrOfPax,AdjustType,OriFkey,NoOriInv"

        strValues = ") values ('" & strTvc & "','" & strBiz & "','" & strAL _
            & "','" & strSrv & "'," & strCustId & ",'" & strCustShortName _
            & "'," & strInvId & ",0,N'" & strCustFullName _
            & "',N'" & strAddress & "','" & strTaxCode & "',N'','--','AUT','SGN'," & decTax & "," & decCharge _
            & ",'" & strMauSo & "','" & strKyHieu & "',N'','" & strEmail & "',N'" _
            & "','','',N'',1," & intAdjustType & ",'" & strOriFkey & "','" & blnNoOriInv & "'"

        strQuerry = strFields & strValues & ")"
        lstQuerry.Add(strQuerry)
        If Not UpdateListOfQuerries(lstQuerry, conn(1), True, intNewInvRecId) Then
            Return 0
        Else
            lstQuerry.Clear()
        End If

        intNewInvId = ScalarToInt("E_Inv78", "InvId", "Recid=" & intNewInvRecId)

        lstQuerry.Add("UPDATE lib.dbo.E_Inv78 set Status='OK' where RecId=" & intNewInvRecId)

        For Each objRow As DataRow In tblTkts.Rows
            If intAdjustType = 3 Then
                strTkno = "Hoàn vé " & objRow("Tkno")
            Else
                strTkno = objRow("Tkno")
            End If
            lstQuerry.Add("insert into lib.dbo.E_InvDetails78 (InvID, Tkno, Description,Unit,Qty,Price, Amount, VatPct" _
                                    & ", VAT,Total,IsSum, Status, FstUser,City) values (" & intNewInvId _
                                    & ",N'" & strTkno & "',N'" & Replace(objRow("Rtg"), " ", "") _
                                    & "','Vé',1," & objRow("Total") & "," & objRow("Total") _
                                    & "," & 0 & "," & 0 _
                                    & "," & objRow("Total") & ",0,'OK','AUT','SGN')")

            lstQuerry.Add("Insert into lib.dbo.E_InvLinks78 (Prod, TKTID, InvId, Status, FstUser,City,RcpId)" _
                              & " values ('AIR'," & objRow("RecId") _
                              & "," & intNewInvId & ",'OK','AUT','SGN'," & objRow("RcpId") & ")")

        Next


        If UpdateListOfQuerries(lstQuerry, conn(1)) Then

        Else
            MsgBox("Unable to update E_InvDetails!")
        End If

        Return intNewInvId
    End Function
    Public Function UpdateListOfQuerries(lstQuerries As List(Of String), objConn As SqlClient.SqlConnection _
                                         , Optional blnGetLastInsertedRecId As Boolean = False _
                                         , Optional ByRef intLastInsertedRecId As Integer = 0) As Boolean

        Dim i As Integer
        Dim strQuerry As String = String.Empty
        If objConn.State = ConnectionState.Closed Then
            objConn.Open()
        End If
        Dim trcSql As SqlClient.SqlTransaction = objConn.BeginTransaction()
        Dim objCmd As SqlClient.SqlCommand = objConn.CreateCommand
        objCmd.Transaction = trcSql

        Try
            For i = 0 To lstQuerries.Count - 1
                strQuerry = lstQuerries(i)
                objCmd.CommandText = strQuerry
                If Not String.IsNullOrEmpty(strQuerry) Then
                    objCmd.ExecuteNonQuery()
                    If blnGetLastInsertedRecId AndAlso UCase(strQuerry).StartsWith("INSERT") Then
                        objCmd.CommandText = "select SCOPE_IDENTITY()"
                        intLastInsertedRecId = objCmd.ExecuteScalar
                    End If
                End If
            Next
            trcSql.Commit()
            Return True
        Catch ex As Exception

            trcSql.Rollback()
            Append2TextFile(vbNewLine & "ERROR|AUT|" & Now & vbNewLine & strQuerry & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Public Function ScalarToInt(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As Integer
        Dim KQ As Integer
        cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
        KQ = cmd.ExecuteScalar
        Return KQ
    End Function
    Public Function ScalarToDec(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As Decimal
        Dim KQ As Decimal
        cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
        KQ = cmd.ExecuteScalar
        Return KQ
    End Function

    Public Function ScalarToString(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As String
        Dim KQ As String
        cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
        KQ = cmd.ExecuteScalar
        Return KQ
    End Function

    'Public Function ScalarToDate(ByVal pTbl As String, ByVal pField As String, ByVal pDK_Order As String) As Date
    '    Dim KQ As Date
    '    cmd.CommandText = "SELECT " & pField & " from " & pTbl & " where " & Finetune_pDK(pDK_Order)
    '    KQ = cmd.ExecuteScalar
    '    Return KQ
    'End Function
    Private Function Finetune_pDK(ByVal pDK As String) As String
        If pDK.Trim.Substring(0, 4).ToUpper = "WHER" Then
            Return pDK.Trim.Substring(5)
        End If
        Return pDK
    End Function
    Public Function CreateKyHieu(strKyHieu As String) As String
        Return "C" & Format(Now, "yy") & "T" & strKyHieu
    End Function
End Module
