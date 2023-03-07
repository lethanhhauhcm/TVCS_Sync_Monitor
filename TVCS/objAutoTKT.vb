Public Class objAutoTKT
    Private _TKT_1A_ID As Integer, _ROEID As Integer, _CustID As Integer
    Private _AL As String, _DocCode As String, _RLOC As String, _Counter As String, _DocNo As String, _FOP As String
    Private _Curr As String, _TKNO As String, _Booker As String, _FTKT As String, _AutoBy As String
    Private _Charge_D As String, _Tax_D As String, _SRV As String
    Private _ROE As Decimal, _CRDAmt As Decimal, _SFinFareCurr As Decimal
    Private _Charge As Decimal, _RoughAmt As Decimal, _TTLFTC As Decimal, _MerchantFee As Decimal, _FOPDetail As String
    Private _DOI As Date, _LogTxt As String, _isNormalRecord As Boolean
    Private _Formula_SF_ROE As String, _FstUser As String, _Location As String
    Private OrgFOP As String, OrgSF As Decimal, SF_Curr As String, CRDCurr As String, TKTsInRLOC As Int16, ALinRLOC As Int16
    Private myCust As New objCustomer
    Private _BspStock As Boolean
    Private mstrVendor As String
    Private mintVendorId As Integer
    Private mstrBookerOnly As String
    Private mdteReturnDate As Date
    Private mstrDomInt As Integer
    Private mstrRptData1 As String
    Private mstrRptData2 As String
    Private mstrRptData3 As String

    ReadOnly Property BookerOnly As String
        Get
            Return mstrBookerOnly
        End Get
    End Property
    ReadOnly Property isNormalRecord As Boolean
        Get
            Return _isNormalRecord
        End Get
    End Property
    ReadOnly Property DOI As Date
        Get
            Return _DOI
        End Get
    End Property
    ReadOnly Property TTLFTC As Decimal
        Get
            Return _TTLFTC
        End Get
    End Property
    ReadOnly Property MerchantFee As Decimal
        Get
            Return _MerchantFee
        End Get
    End Property

    ReadOnly Property RoughAmt As Decimal
        Get
            Return _RoughAmt
        End Get
    End Property
    ReadOnly Property CRDAmt As Decimal
        Get
            Return _CRDAmt
        End Get
    End Property
    ReadOnly Property Charge As Decimal
        Get
            Return _Charge
        End Get
    End Property
    ReadOnly Property SFinFareCurr As Decimal
        Get
            Return _SFinFareCurr
        End Get
    End Property
    ReadOnly Property ROE As Decimal
        Get
            Return _ROE
        End Get
    End Property
    ReadOnly Property CustID As Integer
        Get
            Return _CustID
        End Get
    End Property

    ReadOnly Property ROEID As Integer
        Get
            Return _ROEID
        End Get
    End Property
    ReadOnly Property Booker As String
        Get
            Return _Booker
        End Get
    End Property
    ReadOnly Property Formula_SF_ROE As String
        Get
            Return _Formula_SF_ROE
        End Get
    End Property
    ReadOnly Property LogTxt As String
        Get
            Return _LogTxt
        End Get
    End Property
    ReadOnly Property Location As String
        Get
            Return _Location
        End Get
    End Property
    ReadOnly Property FstUser As String
        Get
            Return _FstUser
        End Get
    End Property

    ReadOnly Property AutoBy As String
        Get
            Return _AutoBy
        End Get
    End Property

    ReadOnly Property Counter As String
        Get
            Return _Counter
        End Get
    End Property
    ReadOnly Property FTKT As String
        Get
            Return _FTKT
        End Get
    End Property
    ReadOnly Property FOPDetail As String
        Get
            Return _FOPDetail
        End Get
    End Property

    ReadOnly Property SRV As String
        Get
            Return _SRV
        End Get
    End Property


    ReadOnly Property TKNO As String
        Get
            Return _TKNO
        End Get
    End Property
    ReadOnly Property AL As String
        Get
            Return _AL
        End Get
    End Property
    ReadOnly Property DocNo As String
        Get
            Return _DocNo
        End Get
    End Property

    ReadOnly Property DocCode As String
        Get
            Return _DocCode
        End Get
    End Property
    ReadOnly Property RLOC As String
        Get
            Return _RLOC
        End Get
    End Property
    ReadOnly Property FOP As String
        Get
            Return _FOP
        End Get
    End Property
    ReadOnly Property Charge_D As String
        Get
            Return _Charge_D
        End Get
    End Property

    ReadOnly Property Curr As String
        Get
            Return _Curr
        End Get
    End Property
    ReadOnly Property BspStock As Boolean
        Get
            Return _BspStock
        End Get
    End Property
    ReadOnly Property Vendor As String
        Get
            Return mstrVendor
        End Get
    End Property
    ReadOnly Property VendorId As Integer
        Get
            Return mintVendorId
        End Get
    End Property
    Public Function SetTkid(iTKT1A_ID As Integer, ByRef conn As SqlClient.SqlConnection) As Boolean
        _Charge_D = ""
        _SFinFareCurr = 0
        _TTLFTC = 0
        _FOPDetail = ""
        _isNormalRecord = False
        _MerchantFee = 0
        _TKT_1A_ID = iTKT1A_ID

        Dim strSelectBspStockId As String = "(select COUNT (RecId) from lib.dbo.MISC where cat='bspstock' and val= substring(tkno,5,4)) as BspStockId"
        Dim dTable As DataTable = GetDataTable("Select *, " & strSelectBspStockId _
                                               & " from TKT_1a where RecID=" & _TKT_1A_ID, conn)

        Dim TTLDue As Decimal, NikeTS24 As String, TS24Amt As Decimal = 0
        Dim cmd As SqlClient.SqlCommand = conn.CreateCommand
        _CustID = dTable.Rows(0)("CustID")
        myCust.SetCustId(_CustID, conn)
        OrgFOP = dTable.Rows(0)("FOP")
        _DocNo = ""
        _Curr = dTable.Rows(0)("Currency")
        _SRV = dTable.Rows(0)("SRV")
        _DOI = dTable.Rows(0)("DOI")
        If Not IsDBNull(dTable.Rows(0)("ReturnDate")) Then
            mdteReturnDate = dTable.Rows(0)("ReturnDate")
        End If

        _ROE = dTable.Rows(0)("ROE")
        _Charge = dTable.Rows(0)("charge")
        _RLOC = dTable.Rows(0)("RLOC")
        _FstUser = dTable.Rows(0)("FstUser")
        OrgSF = dTable.Rows(0)("SvcFee") + dTable.Rows(0)("MU")

        _BspStock = IIf(dTable.Rows(0)("BspStockId") > 0, True, False)

        NikeTS24 = dTable.Rows(0)("Remark")
        If _CustID = 8085 And NikeTS24.Contains("TS24/") Then TS24Amt = TS24Amt + CDec(NikeTS24.Split("/")(2))

        cmd.CommandText = "select count(*) from tkt_1a where RLOC='" & _RLOC & "' and Qty<>0 and custID=" & _CustID
        TKTsInRLOC = cmd.ExecuteScalar
        cmd.CommandText = "select location from tblUser where SICode='" & _FstUser & "' and status<>'XX'"
        _Location = cmd.ExecuteScalar

        _RoughAmt = (dTable.Rows(0)("Fare") + dTable.Rows(0)("Tax") + Charge) * _ROE ' de check credit
        _TKNO = dTable.Rows(0)("TKNO").ToString

        _FTKT = dTable.Rows(0)("FTKT").ToString
        _DocCode = TKNO.Substring(0, 3)
        If _DocCode = "978" Then
            _AL = "VJ"
        Else
            cmd.CommandText = "select AL from Airline where docCode='" & DocCode & "'"
            _AL = cmd.ExecuteScalar
        End If
        SF_Curr = dTable.Rows(0)("SvcFeeCur")
        _Booker = dTable.Rows(0)("Booker")
        mstrBookerOnly = dTable.Rows(0)("Booker")
        _Counter = dTable.Rows(0)("Counter")
        _Location = dTable.Rows(0)("Location")
        _Charge_D = dTable.Rows(0)("ChargeDetail").ToString
        _Tax_D = dTable.Rows(0)("TaxDetail").ToString
        mstrVendor = dTable.Rows(0)("Vendor")
        mintVendorId = dTable.Rows(0)("VendorId")

        If _Booker = "ZPERSONAL" Then
            _Booker = "BIZF|BKR" & _Booker
        Else
            _Booker = "BIZT|BKR" & _Booker
        End If

        _LogTxt = _TKT_1A_ID.ToString & OrgFOP & SF_Curr & OrgSF.ToString
        If TKTsInRLOC = 1 Then
            _AutoBy = "TKID"
            _TTLFTC = dTable.Rows(0)("Fare") + dTable.Rows(0)("Tax") + dTable.Rows(0)("charge")
        ElseIf OrgFOP = "OT" Then
            _AutoBy = "NILL"
        Else
            If InStr("INV_MCE", OrgFOP.Substring(0, 3)) > 0 _
                Or _CustID = 57641 And OrgFOP.StartsWith("CC/") Then
                'ngoai le cho Asia desk
                _AutoBy = "RLOC"
                dTable = GetDataTable("select TKNO, fare, tax, charge, commval, SvcFee, MU, Remark,RptData1,RptData2,RptData3 from tkt_1A where qty <>0 and RLOC='" & _RLOC & "'", conn)  '^_^20230307 add RptData1,RptData2,RptData3 by 7643
                OrgSF = 0
                TS24Amt = 0
                For i As Int16 = 0 To dTable.Rows.Count - 1
                    If dTable.Rows(i)("TKNO").ToString.Substring(0, 3) <> dTable.Rows(0)("TKNO").ToString.Substring(0, 3) Then _AutoBy = "NILL"
                    _TTLFTC = _TTLFTC + dTable.Rows(i)("Fare") + dTable.Rows(i)("Tax") + dTable.Rows(i)("charge")
                    OrgSF = OrgSF + dTable.Rows(i)("SvcFee") + dTable.Rows(i)("MU")
                    NikeTS24 = dTable.Rows(i)("Remark")
                    If _CustID = 8085 And NikeTS24.Contains("TS24/") Then
                        TS24Amt = TS24Amt + CDec(NikeTS24.Split("/")(2))
                    End If

                Next
            Else
                _AutoBy = "NILL"
            End If
        End If

        If _AutoBy = "NILL" Then Return True

        If _BspStock Then
            '^_^20221028 mark by 7643 -b-
            'cmd.CommandText = "select top 1 RecID from ForEx" _
            '                & " where Status='OK' and Currency='USD'" _
            '                & " and ApplyRoeTo='GDS' and effectDate='" & Format(_DOI, "dd MMM yy") _
            '                & "' order by recID desc"
            '^_^20221028 mark by 7643 -e-
            '^_^20221028 modi by 7643 -b-
            cmd.CommandText = "select top 1 RecID from ForEx" _
                            & " where Status='OK' and Currency='USD'" _
                            & " and ApplyRoeTo='GDS' and effectDate='" & Format(_DOI, "dd MMM yyyy") _
                            & "' order by recID desc"
            '^_^20221028 modi by 7643 -e-
            _ROEID = cmd.ExecuteScalar
        Else
            _ROEID = ForEX_12(Now.Date, "USD", "RECID", "TS")
            'Else
            '    cmd.CommandText = "select top 1 RecID from ras12.dbo.forEx where BSR=" & _ROE &
            '        " And Currency='" & _Curr & "' order by recID desc, status"
            '    _ROEID = cmd.ExecuteScalar
        End If
        If _ROEID = 0 Then
            MsgBox(Now & vbNewLine & "Đề nghị quầy TVS SGN chạy FasTicket để lấy tỷ giá cho ngày " & _DOI)
            Application.Exit()
            End
        End If
        _SFinFareCurr = OrgSF
        If _Curr = SF_Curr Then
            _Formula_SF_ROE = "*1"
        Else
            If _Curr = "VND" Then
                cmd.CommandText = "select BSR from ForEx where RecID=" & _ROEID
                Dim tmpROE As Decimal = cmd.ExecuteScalar
                _SFinFareCurr = _SFinFareCurr * tmpROE
                _Formula_SF_ROE = "*" & tmpROE.ToString
            Else
                _SFinFareCurr = _SFinFareCurr / _ROE
                _Formula_SF_ROE = "/" & _ROE.ToString
            End If
        End If
        TTLDue = _TTLFTC + _SFinFareCurr

        If OrgFOP.Contains("CC/") Then
            _DocNo = OrgFOP.Split("/")(1)
            CRDCurr = OrgFOP.Split("/")(2).Substring(0, 3)
            _CRDAmt = OrgFOP.Split("/")(2).Substring(3)
            _FOP = "CRD"
            If CRDCurr = _Curr Then
                _MerchantFee = _CRDAmt - (_TTLFTC + _SFinFareCurr)
            ElseIf CRDCurr = "VND" Then
                _MerchantFee = _CRDAmt - (_TTLFTC + _SFinFareCurr) * _ROE
                _MerchantFee = _MerchantFee / _ROE
            ElseIf _Curr = "VND" Then
                _MerchantFee = 0 ' nen theo doi so lieu
            End If
            _FOPDetail = _FOP & "_" & CRDCurr & "_" & _CRDAmt.ToString

            If _CustID = 8085 AndAlso TS24Amt <> 0 AndAlso _Booker.StartsWith("BIZT|BKR") Then 'NIKE
                _FOPDetail = _FOPDetail & "|3RD_" & SF_Curr & "_" & TS24Amt.ToString.Trim
            End If

            _MerchantFee = Math.Round(_MerchantFee, 2)
            TTLDue = TTLDue + _MerchantFee

        ElseIf OrgFOP = "CASH" Then
            _FOP = "CSH"
            _FOPDetail = "CSH_VND_" & (TTLDue * _ROE).ToString

        ElseIf Mid(OrgFOP, 1, 3) = "INV" Then
            If myCust.DelayType = "DEB" Then Return False
            _FOP = myCust.DelayType
            If OrgFOP.Contains("/") Then
                _DocNo = OrgFOP.Split("/")(1)
            End If

            If _CustID = 5830 Then ' MAST
                _FOPDetail = _FOP & "_" & SF_Curr & "_" & OrgSF.ToString
                _FOPDetail = _FOPDetail & "|" & _FOP & "_VND_" & (_TTLFTC * _ROE).ToString.Trim

            ElseIf _CustID = 55152 Or _CustID = 55154 Then   ' PG INDOCHINA, PG VIETNAM
                _FOPDetail = "PSP_" & SF_Curr & "_" & OrgSF.ToString
                _FOPDetail = _FOPDetail & "|PSP_VND_" & (_TTLFTC * _ROE).ToString.Trim

            ElseIf _FOP = "PSP" Then
                _FOPDetail = "PSP_" & _Curr & "_" & TTLDue.ToString

                If _CustID = 8085 AndAlso TS24Amt <> 0 Then ' NIKE
                    _FOPDetail = _FOPDetail & "|3RD_" & SF_Curr & "_" & TS24Amt.ToString.Trim
                    _FOPDetail = _FOPDetail & "|PSP_" & SF_Curr & "_-" & TS24Amt.ToString.Trim
                End If
            Else
                _FOPDetail = "PPD_VND_" & (TTLDue * _ROE).ToString
            End If
        ElseIf OrgFOP.Contains("MCE/") Then
            _DocNo = OrgFOP.Split("/")(1)
            _FOP = "MCE"
            _FOPDetail = _FOP & "_" & _Curr & "_" & TTLDue.ToString
        End If
        mstrRptData1 = dTable.Rows(0)("RptData1")
        mstrRptData2 = dTable.Rows(0)("RptData2")
        mstrRptData3 = dTable.Rows(0)("RptData3")
        _isNormalRecord = True
        Return True
    End Function
    ReadOnly Property TKT_1A_ID As Integer
        Get
            Return _TKT_1A_ID
        End Get

    End Property
    Public Property ReturnDate As Date
        Get
            Return mdteReturnDate
        End Get
        Set(value As Date)
            mdteReturnDate = value
        End Set
    End Property
    Public Property DomInt As String
        Get
            Return mstrDomInt
        End Get
        Set(value As String)
            mstrDomInt = value
        End Set
    End Property
    Public Property TaxDetail As String
        Get
            Return _Tax_D
        End Get
        Set(value As String)
            _Tax_D = value
        End Set
    End Property
    ReadOnly Property RptData1 As String
        Get
            Return mstrRptData1
        End Get
    End Property
    ReadOnly Property RptData2 As String
        Get
            Return mstrRptData2
        End Get
    End Property
    ReadOnly Property RptData3 As String
        Get
            Return mstrRptData3
        End Get
    End Property
End Class
