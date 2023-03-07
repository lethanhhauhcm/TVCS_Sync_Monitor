Public Class objCustomer
    Private _CustID As Integer, _ShortName As String, _FullName As String, _Addr As String, _TaxCode As String
    Private _CustType As String, _DelayType As String, _AdhType As String, _LstReconcile As Date
    Private _CrditOnAL As String, _CustOf_AL As String, _MinBLC As Decimal, _Coef As Decimal
    Private _CurrBLC As Decimal
    Dim _CnStr As String

    Property CnStr() As String
        Get
            Return _CnStr
        End Get
        Set(ByVal sCnStr As String)
            _CnStr = sCnStr
        End Set
    End Property
    ReadOnly Property Coef() As Decimal
        Get
            Return _Coef
        End Get
    End Property

    ReadOnly Property CurrBLC() As Decimal
        Get
            Return _CurrBLC
        End Get
    End Property

    ReadOnly Property MinBLC() As Decimal
        Get
            Return _MinBLC
        End Get
    End Property


    ReadOnly Property LstReconcile() As Date
        Get
            Return _LstReconcile
        End Get
    End Property
    ReadOnly Property CrditOnAL() As String
        Get
            Return _CrditOnAL
        End Get
    End Property
    ReadOnly Property CustOf_AL() As String
        Get
            Return _CustOf_AL
        End Get
    End Property
    ReadOnly Property AdhType() As String
        Get
            Return _AdhType
        End Get
    End Property
    Public Function SetCustId(ByVal iCustID As Integer, ByRef conn As SqlClient.SqlConnection) As Boolean
        _CustID = iCustID
        _CustOf_AL = ""
        _Coef = 0
        _MinBLC = 0
        _CrditOnAL = ""
        _AdhType = ""
        If _CustID = 0 Then Return False
        Dim dTable As DataTable
        Dim cmd As SqlClient.SqlCommand = conn.CreateCommand

        dTable = GetDataTable("select CustShortName, CustFullName,CustAddress, CustTaxCode, Email" _
                              & " from CustomerList " & "where status<>'XX' and recID=" & _CustID, conn)

        _ShortName = dTable.Rows(0)("CustShortName")
        _FullName = dTable.Rows(0)("CustFullName")
        _FullName = _FullName.Replace("'", "")
        _Addr = dTable.Rows(0)("CustAddress")
        _Addr = _Addr.Replace("'", "")
        _TaxCode = dTable.Rows(0)("CustTaxCode")


        If _ShortName <> "" Then
            dTable = GetDataTable("select CAT, VAL from Cust_Detail where custid=" & _CustID & " and status <>'XX'", conn)
            For i As Int16 = 0 To dTable.Rows.Count - 1
                If dTable.Rows(i)("CAT") = "AL" Then _CustOf_AL = _CustOf_AL & "_" & dTable.Rows(i)("VAL")
            Next
            If _CustOf_AL.Length > 1 Then _CustOf_AL = _CustOf_AL.Substring(1)
            _CustOf_AL = _CustOf_AL.Replace("_", "','")
            _CustOf_AL = "('" & _CustOf_AL & "')"

            dTable = GetDataTable("select CRCoef, PPCoef, Adh, FoxCoef, MinBLC, AL, ADH " &
                                  "from CC_Setting where status='OK' and  custid=" & _CustID, conn)
            _DelayType = "DEB"
            _CurrBLC = 0
            If dTable.Rows.Count = 1 Then
                If dTable.Rows(0)("CRCoef") > 0 Then
                    _DelayType = "PSP"
                ElseIf dTable.Rows(0)("PPCoef") > 0 Then
                    _DelayType = "PPD"
                End If
                _Coef = IIf(dTable.Rows(0)("CRCoef") > 0, dTable.Rows(0)("CRCoef") > 0, dTable.Rows(0)("PPCoef") > 0)
                _MinBLC = dTable.Rows(0)("MinBLC")
                _CrditOnAL = dTable.Rows(0)("AL")
                _AdhType = dTable.Rows(0)("ADH")

            End If
            cmd.CommandText = "select VAL from cust_detail where custID=" & _CustID & " and status='OK' and cat='Channel'"
            _CustType = cmd.ExecuteScalar

            If InStr("PPD_PSP", _DelayType) > 0 Then
                cmd.CommandText = "select " & IIf(_DelayType = "PPD", "top 1 VND_PPD_Avail", "top 1 VND_PSP_Avail") &
                    " from cc_BLC where custID=" & _CustID & " order by recID desc"
                _CurrBLC = cmd.ExecuteScalar
            End If
            If InStr("PSP_PPD", _DelayType) > 0 Then
                cmd.CommandText = "select top 1 AsOf from chotcongno where custid=" & _CustID & " and status <>'XX' order by asof desc"
                _LstReconcile = cmd.ExecuteScalar
            End If
        End If
        Return True
    End Function
    ReadOnly Property CustID() As Integer
        Get
            Return _CustID
        End Get

    End Property
    ReadOnly Property ShortName() As String
        Get
            Return _ShortName
        End Get
    End Property

    ReadOnly Property FullName() As String
        Get
            Return _FullName
        End Get
    End Property

    ReadOnly Property taxCode() As String
        Get
            Return _TaxCode
        End Get
    End Property
    ReadOnly Property Addr() As String
        Get
            Return _Addr
        End Get
    End Property

    Property CustType() As String
        Get
            Return _CustType
        End Get
        Set(ByVal sCustType As String)
            _CustType = sCustType
        End Set
    End Property

    ReadOnly Property DelayType() As String
        Get
            Return _DelayType
        End Get
    End Property
End Class
