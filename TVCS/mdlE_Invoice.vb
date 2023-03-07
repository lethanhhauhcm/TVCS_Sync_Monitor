Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Xml
Imports System.Xml.Linq
Imports TVCS.MySharedFunctions
Module mdlE_Invoice
    Public pstrSerialCertAPG As String = "540101040D9E2D54F27A436778B666B2"
    Public pblnTestInv As Boolean
    Public pblnTT78 As Boolean
    Public Enum InvAdjustType
        IncreaseAmt = 2
        DecreaseAmt = 3
        ChangeInfo = 4
        Replace = 9
    End Enum
    Structure Tvc4E_Inv
        Const ACR = "ACR"
        Const GDS = "GDS"
        Const GDS_HAN = "GDS HAN"
        Const TVPM = "TVPM"
        Const TVPM_SGN = "TVPM SGN"
        Const TVTR = "TVTR"
        Const TVTR_DAD = "TVTR DAD"
        Const TVTR_HAN = "TVTR HAN"
        Const VLM = "VLM"
        Const VLM_HAN = "VLM HAN"

    End Structure
    Structure FunctionName
        Const AdjustActionAssignedNo = "AdjustActionAssignedNo"
        Const AdjustInvoiceAction = "AdjustInvoiceAction"
        Const AdjustInvoiceNote = "AdjustInvoiceNote"
        Const AdjustInvoiceNoPublish = "AdjustInvoiceNoPublish"
        Const AdjustReplaceInvWithToken = "AdjustReplaceInvWithToken"
        Const AdjustWithoutInv = "AdjustWithoutInv"
        Const cancelInv = "cancelInv"
        Const cancelInvNoPay = "cancelInvNoPay"
        Const CancelInvoiceWithToken = "CancelInvoiceWithToken"
        Const deleteInvoiceByFkey = "deleteInvoiceByFkey"
        Const deleteInvoiceByID = "deleteInvoiceByID"
        Const downloadInv = "downloadInv"
        Const downloadInvFkey = "downloadInvFkey"
        Const downloadInvFkeyNoPay = "downloadInvFkeyNoPay"
        Const downloadInvNoPay = "downloadInvNoPay"
        Const downloadInvPDF = "downloadInvPDF"
        Const downloadInvPDFNoPay = "downloadInvPDFNoPay"
        Const downloadNewInvPDFFkey = "downloadNewInvPDFFkey"
        Const getHashInv = "getHashInv"
        Const getHashInvWithToken = "getHashInvWithToken"
        Const GetCertInfo = "GetCertInfo"
        Const GetInvByFkey = "GetInvByFkey"
        Const getInvView = "getInvView"
        Const getInvViewNoPay = "getInvViewNoPay"
        Const getInvViewFkeyNoPay = "getInvViewFkeyNoPay"
        Const getNewInvViewFkey = "getNewInvViewFkey"
        Const HandleInvoiceErrors = "HandleInvoiceErrors"
        Const ImportAndPublishInv = "ImportAndPublishInv"
        Const ImportInv = "ImportInv"
        Const ImportInvByPattern = "ImportInvByPattern"
        Const listInvFromNoToNo = "listInvFromNoToNo"
        Const PublishInvFkey = "PublishInvFkey"
        Const publishInvWithToken = "publishInvWithToken"
        Const ReceivedInvoiceErrors = "ReceivedInvoiceErrors"
        Const ReplaceInvoiceAction = "ReplaceInvoiceAction"
        Const ReplaceInvoiceNoPublish = "ReplaceInvoiceNoPublish"
        Const ReplaceWithoutInv = "ReplaceWithoutInv"
        Const SendAgainEmailServ = "SendAgainEmailServ"
        Const SendInvNoticeErrors = "SendInvNoticeErrors"
        Const UpdateCus = "UpdateCus"

    End Structure
    Structure FOP
        Const Thanh_toán_chuyển_khoản = "CK"
        Const Thanh_toán_tiền_mặt = "TM"
        Const Thanh_toán_thẻ_tín_dụng = "TTD"
        Const Hình_thức_HDDT = "HDDT"
        Const Hình_thức_thanh_toán_tiền_mặt_hoặc_chuyển_khoản = "TM/CK"
        Const Thanh_toán_bù_trừ = "BT"
    End Structure
    Structure InvoiceStatus
        Const Hóa_đơn_vừa_khởi_tạo = 0
        Const Hóa_đơn_có_đủ_chữ_ký = 1
        Const Hóa_đơn_đã_khai_báo_thuế = 2
        Const Hóa_đơn_sai_sót_bị_thay_thế = 3
        Const Hóa_đơn_sai_sót_bị_điều_chỉnh = 4
        Const Hóa_đơn_xóa_bỏ = 5
    End Structure
    Structure InvoiceType
        Const Hóa_đơn_thông_thường = 0
        Const Hóa_đơn_thay_thế = 1
        Const Hóa_đơn_điều_chỉnh_tăng = 2
        Const Hóa_đơn_điều_chỉnh_giảm = 3
        Const Hóa_đơn_điều_chỉnh_thông_tin = 4
    End Structure
    Structure KindOfService
        Const Hóa_đơn_GTGT = 0
        Const Hoàn_trả_vé = 3

    End Structure
    Structure VatRate
        Const Không_chịu_thuế = "-1"
        Const Không_kê_khai_và_nộp_thuế = "-2"
        Const Không_% = 0
        Const Năm_% = 5
        Const Mười_% = 10
    End Structure

    Public Function UpLoadCustomer2VNPT(lstWsConnects As List(Of clsE_InvConnect) _
                                        , lstCustData As List(Of XElement), intCustId As Integer) As Boolean
        Dim objE_Inv As New clsE_Invoice
        Dim blnErrorHappened As Boolean = False
        For Each objWsConnect As clsE_InvConnect In lstWsConnects
            If Not objE_Inv.MassUpdateCus2(objWsConnect.WsUrl, objWsConnect.UserName, objWsConnect.UserPass _
                , lstCustData) Then
                MsgBox("unable to update Customer with CustId " & intCustId _
                       & " for " & objWsConnect.Tvc & vbNewLine & objE_Inv.ResponseDesc)
                'Append2TextFile("E_InVError:" & objWsConnect.WsUrl & vbNewLine & "Request:" & vbNewLine & objE_Inv.LastRequest & vbNewLine & "Response:" & vbNewLine & objE_Inv.LastResponse)
                blnErrorHappened = True
            ElseIf objE_Inv.ReponseCode < 0 Then
                MsgBox("unable to update Customer with CustId " & intCustId _
                       & " for " & objWsConnect.Tvc & vbNewLine & objE_Inv.ResponseDesc)
                ''Append2TextFile("E_InVError:" & objWsConnect.WsUrl & vbNewLine & objE_Inv.LastRequest & vbNewLine & objE_Inv.LastResponse)
                blnErrorHappened = True
            End If
        Next
        If blnErrorHappened Then

            Return False
        ElseIf Not ExecuteNonQuerry("Update LIB.DBO.Customer set UploadDate=Getdate() where RecId=" _
                                 & intCustId, conn(0)) Then
            MsgBox("unable to update UploadDate for Customer " & intCustId)
            Return False
        End If

        Return True
    End Function
    Public Function GetE_InvConnects(strBiz As String, strCity As String, Optional strTvc As String = "") As List(Of clsE_InvConnect)
        Dim tblTvc As DataTable
        Dim lstWsConnects As New List(Of clsE_InvConnect)
        Dim strQuerry As String = "select distinct biz,tvc,City from lib.dbo.E_InvSettings" _
                              & " where status='OK' and Biz in ('" _
                              & Replace(strBiz, ",", "','") _
                              & "') and City='" & strCity & "'"

        If strTvc <> "" Then
            strQuerry = strQuerry & " and TVC='" & strTvc & "'"
        End If
        strQuerry = strQuerry & " order by biz,City"
        tblTvc = GetDataTable(strQuerry, conn_Web)

        For Each objRow As DataRow In tblTvc.Rows
            Dim objConnect As New clsE_InvConnect(pblnTT78, objRow("TVC"))
            lstWsConnects.Add(objConnect)
        Next

        Return lstWsConnects
    End Function

    Public Function MassUpLoadCustomer2VNPT()
        Dim tblCustomer As System.Data.DataTable
        Dim strQuerry As String
        Dim objE_Inv As New clsE_Invoice
        Dim intCustType As Integer = 0
        Dim lstWsConnectsPaxSGN As New List(Of clsE_InvConnect)
        Dim lstWsConnectsPaxHAN As New List(Of clsE_InvConnect)
        Dim lstWsConnectsCgoSGN As New List(Of clsE_InvConnect)
        Dim lstWsConnectsCgoHAN As New List(Of clsE_InvConnect)
        Dim lstWsConnectsPaxCgoSGN As New List(Of clsE_InvConnect)
        Dim lstWsConnectsPaxCgoHAN As New List(Of clsE_InvConnect)
        Dim lstWsConnectsPaxSGN_APG As New List(Of clsE_InvConnect)

        Dim blnErrorHappened As Boolean
        Dim lstCustData As New List(Of XElement)
        Dim lstQuerries As New List(Of String)

        'intCustType    = 1 nghia la khach hang se ky hoa don dien tu

        'lstWsConnectsPaxSGN = GetE_InvConnects("PAX", "SGN")
        'lstWsConnectsPaxHAN = GetE_InvConnects("PAX", "HAN")
        'lstWsConnectsCgoSGN = GetE_InvConnects("CGO", "SGN")
        'lstWsConnectsCgoHAN = GetE_InvConnects("CGO", "HAN")
        'lstWsConnectsPaxCgoSGN = GetE_InvConnects("PAX,CGO", "SGN")
        'lstWsConnectsPaxCgoHAN = GetE_InvConnects("PAX,CGO", "HAN")
        lstWsConnectsPaxSGN_APG = GetE_InvConnects("PAX", "SGN", "APG")

        strQuerry = "Select * from Lib.dbo.Customer" _
            & " where RECID>0  and CustFullName<>'' and STATUS='OK' and (APP Like '%RAS%' or APP Like '%COS%')" _
            & " order by RecId "
        strQuerry = "Select * from Lib.dbo.Customer" _
            & " where RECID>0  and CustFullName<>'' and STATUS='OK' and (APP Like '%RAS%' or APP Like '%COS%') and RecId=20 " _
            & " order by RecId "

        '    & " and (UploadDate is null or LstUpdate>UploadDate)" _
        '    & " order by RecId "
        tblCustomer = GetDataTable(strQuerry, conn_Web)
        For Each objRow As DataRow In tblCustomer.Rows
            Dim strEmail As String = Replace(objRow("InvoiceEmail"), ";", ",")
            'strEmail = "khanh.nguyenminh@transviet.com"
            blnErrorHappened = False
            If objRow("CustSignatureRq") Then
                intCustType = 1
            Else
                intCustType = 0
            End If

            lstCustData.Add(CreateCustData(objRow("RecId"), objRow("CustFullName"), objRow("CustShortName") _
                    , objRow("CustTaxCode"), objRow("CustAddress"), strEmail, intCustType _
                    , objRow("Phone"), objRow("City")))

            UpLoadCustomer2VNPT(lstWsConnectsPaxSGN_APG, lstCustData, objRow("RecId"))

            'Select Case objRow("City")
            '    Case "SGN"
            '        If objRow("APP").ToString.Contains("RAS") AndAlso objRow("APP").ToString.Contains("COS") Then
            '            UpLoadCustomer2VNPT(lstWsConnectsPaxCgoSGN, lstCustData, objRow("RecId"))
            '        ElseIf objRow("APP").ToString.Contains("RAS") Then
            '            UpLoadCustomer2VNPT(lstWsConnectsPaxSGN, lstCustData, objRow("RecId"))
            '        ElseIf objRow("APP").ToString.Contains("COS") Then
            '            UpLoadCustomer2VNPT(lstWsConnectsCgoSGN, lstCustData, objRow("RecId"))
            '        End If
            '    Case "HAN"
            '        If objRow("APP").ToString.Contains("RAS") AndAlso objRow("APP").ToString.Contains("COS") Then
            '            UpLoadCustomer2VNPT(lstWsConnectsPaxCgoHAN, lstCustData, objRow("RecId"))
            '        ElseIf objRow("APP").ToString.Contains("RAS") Then
            '            UpLoadCustomer2VNPT(lstWsConnectsPaxHAN, lstCustData, objRow("RecId"))
            '        ElseIf objRow("APP").ToString.Contains("COS") Then
            '            UpLoadCustomer2VNPT(lstWsConnectsCgoHAN, lstCustData, objRow("RecId"))
            '        End If
            'End Select
            lstCustData.Clear()
        Next
        Return True
    End Function

    'Public Function CreateKyHieu(strKyHieu As String) As String
    '    Return strKyHieu & "/20E"
    '    'Return strKyHieu & "/" & Format(dteDOI, "yy") & "E"
    'End Function
    Public Function CreateMauSo(strMauSo As String) As String
        If pblnTT78 Then
            Return strMauSo
        Else
            Return "01GTKT0/" & strMauSo
        End If

    End Function
    Public Function CreateKyHieu(strKyHieu As String) As String
        If pblnTT78 Then
            Return "C" & Format(Now, "yy") & "T" & strKyHieu
        Else
            Return strKyHieu
        End If

    End Function
    Public Function CreateInvProduct(objProduct As clsProduct, Optional blnTT78 As Boolean = False) As XElement
        Dim objXElement As XElement
        Dim decProductPrice As Decimal = objProduct.ProdPrice

        If Not blnTT78 Then
            decProductPrice = decProductPrice
        End If

        objXElement =
                <Product>
                    <Remark><%= objProduct.Seq %></Remark>
                    <Code><%= objProduct.ProductCode %></Code>
                    <ProdName><%= objProduct.ProdName %></ProdName>
                    <Total><%= objProduct.TotalPrice %></Total>
                    <DiscountAmount><%= objProduct.DiscountAmount %></DiscountAmount>
                    <VATRate><%= objProduct.VatRate %></VATRate>
                    <VATAmount><%= objProduct.VatAmount %></VATAmount>
                    <Amount><%= objProduct.TotalPrice + objProduct.VatAmount %></Amount>
                    <Extra1 <%= objProduct.Extra1 %>></Extra1>
                </Product>
        '<ProdUnit></ProdUnit>

        If objProduct.ProdUnit <> "" Then
            Dim objUnit As XElement =
                <ProdUnit><%= objProduct.ProdUnit %></ProdUnit>
            objXElement.Add(objUnit)
        End If
        If objProduct.ProdPrice <> 0 Then
            Dim objPrice As XElement =
            <ProdPrice><%= decProductPrice %></ProdPrice>
            objXElement.Add(objPrice)
        End If
        If objProduct.ProdQuantity <> 0 Then
            Dim objQuantity As XElement =
                <ProdQuantity><%= objProduct.ProdQuantity %></ProdQuantity>
            objXElement.Add(objQuantity)
        End If

        If pblnTT78 Then
            Dim objQuantity As XElement =
                <IsSum><%= objProduct.IsSum %></IsSum>
            objXElement.Add(objQuantity)
        End If

        Return objXElement
        'Catch ex As Exception
        '    MsgBox("unable to create product for VAT Inv")
        '    Return Nothing
        'End Try

    End Function
    Public Function CreateInvProductTt78(objProduct As clsProduct) As XElement
        Dim objXElement As XElement
        Dim decProductPrice As Decimal = objProduct.ProdPrice
        Dim intTinhChat As Integer
        Select Case objProduct.IsSum
            Case 4
                intTinhChat = 4
            Case Else
                intTinhChat = objProduct.IsSum + 1
        End Select
        objXElement =
                <HHDVu>
                    <TChat><%= intTinhChat %></TChat>
                    <STT><%= objProduct.Seq %></STT>
                    <MHHDVu><%= objProduct.ProductCode %></MHHDVu>
                    <THHDVu><%= objProduct.ProdName %></THHDVu>
                    <TLCKhau><%= objProduct.DiscountRate %></TLCKhau>
                    <STCKhau><%= objProduct.DiscountAmount %></STCKhau>
                    <TSuat><%= ConvertVatPct2ThueSuat(objProduct.VatRate) %></TSuat>
                    <ThTien><%= objProduct.TotalPrice %></ThTien>
                    <TThue><%= objProduct.VatAmount %></TThue>
                </HHDVu>
        '<TSuat>KKKNT</TSuat>
        '<TSThue><%= objProduct.TotalPrice + objProduct.VatAmount %></TSThue>
        objXElement.Add(<TTKhac>
                            <TTin>
                                <TTruong>Extra1</TTruong>
                                <KDLieu>string</KDLieu>
                                <DLieu><%= objProduct.Extra1 %></DLieu>
                            </TTin>
                            <TTin>
                                <TTruong>Amount</TTruong>
                                <KDLieu>numeric</KDLieu>
                                <DLieu><%= objProduct.TotalPrice + objProduct.VatAmount %></DLieu>
                            </TTin>
                            <TTin>
                                <TTruong>VATAmount</TTruong>
                                <KDLieu>numeric</KDLieu>
                                <DLieu><%= objProduct.VatAmount %></DLieu>
                            </TTin>
                        </TTKhac>)

        If objProduct.ProdUnit <> "" Then
            Dim objUnit As XElement =
                <DVTinh><%= objProduct.ProdUnit %></DVTinh>
            objXElement.Add(objUnit)
        End If
        If objProduct.ProdPrice <> 0 Then
            Dim objPrice As XElement =
            <DGia><%= decProductPrice %></DGia>
            objXElement.Add(objPrice)
        End If
        If objProduct.ProdQuantity <> 0 Then
            Dim objQuantity As XElement =
                <SLuong><%= objProduct.ProdQuantity %></SLuong>
            objXElement.Add(objQuantity)
        End If

        Return objXElement
        'Catch ex As Exception
        '    MsgBox("unable to create product for VAT Inv")
        '    Return Nothing
        'End Try

    End Function
    Public Function CreateInvData(strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strSerialCert As String = "", Optional intVatDiscount As Integer = 0) As XElement

        Dim objXElement As XElement
        Dim objInv As XElement
        Dim objInvoice As XElement
        Dim objVatDiscount As XElement

        Dim objXProd As XElement

        Dim decVat0 As Decimal = 0
        Dim decVat5 As Decimal = 0
        Dim decVat8 As Decimal = 0
        Dim decVat10 As Decimal = 0
        Dim decVatableNull As Decimal = 0
        Dim decVatable0 As Decimal = 0
        Dim decVatable5 As Decimal = 0
        Dim decVatable8 As Decimal = 0
        Dim decVatable10 As Decimal = 0

        Dim decTotalDiscount As Decimal = 0
        Dim decTotalVat As Decimal = 0
        Dim decTotalVatable As Decimal = 0
        Dim decTotalInv As Decimal = 0
        Dim intRefundMultiplier As Integer = 1

        If strKindOfService = 3 Then
            intRefundMultiplier = -1
        End If

        If strPattern = "" Then
            If pblnTT78 Then
                strPattern = "1/001"
            Else
                strPattern = "01GTKT0/001"
            End If

        End If
        If strSerial = "" Then
            If pblnTT78 Then
                strSerial = "C" & Format(Now, "yy") & "TAA"
            Else
                strSerial = "AA/" & Format(Now, "yy") & "E"
            End If

        End If

        If strSerialCert = "" Then
            objXElement =
            <Invoices>
            </Invoices>
        Else
            objXElement =
            <Invoices>
                <SerialCert><%= strSerialCert %></SerialCert>
            </Invoices>
        End If

        objInv =
                <Inv>
                    <key><%= strInvoiceKey %></key>
                </Inv>

        objXProd =
                <Products>
                </Products>
        For Each objProduct As clsProduct In lstProducts
            objXProd.Add(CreateInvProduct(objProduct))
            Select Case objProduct.VatRate
                Case 10
                    decVatable10 = decVatable10 + objProduct.TotalPrice
                    decVat10 = decVat10 + objProduct.VatAmount
                Case 8
                    decVatable8 = decVatable8 + objProduct.TotalPrice
                    decVat8 = decVat8 + objProduct.VatAmount
                Case 5
                    decVatable5 = decVatable5 + objProduct.TotalPrice
                    decVat5 = decVat5 + objProduct.VatAmount
                Case 0
                    decVatable0 = decVatable0 + objProduct.TotalPrice
                    decVat0 = decVat0 + objProduct.VatAmount
                Case -1
                    decVatableNull = decVatableNull + objProduct.TotalPrice

            End Select
            decTotalDiscount = decTotalDiscount + objProduct.DiscountAmount
            decTotalVat = decTotalVat + objProduct.VatAmount
            decTotalVatable = decTotalVatable + objProduct.TotalPrice
            decTotalInv = decTotalInv + objProduct.Amount - (objProduct.DiscountAmount * intRefundMultiplier)
        Next

        If strPattern.EndsWith("1") Then
            decTotalInv = decTotalInv + decCharge
        Else
            decTotalInv = decTotalInv + (decTax * intRefundMultiplier) + decCharge
        End If

        objInvoice =
                <Invoice>
                    <Buyer><%= strBuyer %></Buyer>
                    <CusCode><%= strCustId %></CusCode>
                    <CusName><%= strCustFullName %></CusName>
                    <CusAddress><%= strAddress %></CusAddress>
                    <CusPhone><%= strPhone %></CusPhone>
                    <EmailDeliver><%= strEmail %></EmailDeliver>
                    <CusTaxCode><%= strTaxCode %></CusTaxCode>
                    <PaymentMethod><%= strFOP %></PaymentMethod>
                    <KindOfService><%= strKindOfService %></KindOfService>
                    <Total><%= decTotalVatable %></Total>
                    <DiscountAmount><%= decTotalDiscount %></DiscountAmount>
                    <VATAmount><%= decTotalVat %></VATAmount>
                    <VatAmount0><%= decVat0 %></VatAmount0>
                    <VatAmount5><%= decVat5 %></VatAmount5>
                    <VatAmount10><%= decVat10 %></VatAmount10>
                    <GrossValue><%= decVatableNull * intRefundMultiplier %></GrossValue>
                    <GrossValue0><%= decVatable0 %></GrossValue0>
                    <GrossValue5><%= decVatable5 %></GrossValue5>
                    <GrossValue10><%= decVatable10 %></GrossValue10>
                    <Extra1><%= decTax %></Extra1>
                    <Extra2><%= decCharge %></Extra2>
                    <Amount><%= decTotalInv %></Amount>
                    <AmountInWords><%= TienBangChu(Math.Floor(decTotalInv)) & " đồng" %></AmountInWords>
                </Invoice>
        If pblnTT78 Then
            Dim objVatAmt8 As XElement =
                <VatAmount8><%= decVat8 %></VatAmount8>
            Dim objGrossAmt8 As XElement =
                <GrossValue8><%= decVatable8 %></GrossValue8>
            objInvoice.Add(objVatAmt8)
            objInvoice.Add(objGrossAmt8)
        End If

        'objInvoice =
        '<Invoice>
        '    <Buyer><%= strBuyer %></Buyer>
        '    <CusCode><%= strCustId %></CusCode>
        '    <CusName><%= strCustFullName %></CusName>
        '    <CusAddress><%= strAddress %></CusAddress>
        '    <CusPhone><%= strPhone %></CusPhone>
        '    <EmailDeliver><%= strEmail %></EmailDeliver>
        '    <CusTaxCode><%= strTaxCode %></CusTaxCode>
        '    <PaymentMethod><%= strFOP %></PaymentMethod>
        '    <KindOfService><%= strKindOfService %></KindOfService>
        '    <Total><%= Math.Abs(decTotalVatable) %></Total>
        '    <DiscountAmount><%= decTotalDiscount %></DiscountAmount>
        '    <VATAmount><%= Math.Abs(decTotalVat) %></VATAmount>
        '    <VatAmount0><%= Math.Abs(decVat0) %></VatAmount0>
        '    <VatAmount5><%= Math.Abs(decVat5) %></VatAmount5>
        '    <VatAmount8><%= Math.Abs(decVat8) %></VatAmount8>
        '    <VatAmount10><%= Math.Abs(decVat10) %></VatAmount10>
        '    <GrossValue><%= decVatableNull * intRefundMultiplier %></GrossValue>
        '    <GrossValue0><%= Math.Abs(decVatable0) %></GrossValue0>
        '    <GrossValue5><%= Math.Abs(decVatable5) %></GrossValue5>
        '    <GrossValue8><%= Math.Abs(decVatable8) %></GrossValue8>
        '    <GrossValue10><%= Math.Abs(decVatable10) %></GrossValue10>
        '    <Extra1><%= decTax %></Extra1>
        '    <Extra2><%= decCharge %></Extra2>
        '    <Amount><%= Math.Abs(decTotalInv) %></Amount>
        '    <AmountInWords><%= TienBangChu(Math.Abs(decTotalInv)) & " đồng" %></AmountInWords>
        '</Invoice>


        'add thong tin hoan chinh vao day
        objInvoice.Add(objXProd)
        objInv.Add(objInvoice)
        objXElement.Add(objInv)

        Return objXElement
    End Function
    Public Function CreateInvDataAdjust(strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strSerialCert As String = "", Optional intVatDiscount As Integer = 0) As XElement

        'Dim objXElement As XElement
        'Dim objInv As XElement
        Dim objInvoice As XElement
        'Dim objVatDiscount As XElement

        Dim objXProd As XElement

        Dim decVat0 As Decimal = 0
        Dim decVat5 As Decimal = 0
        Dim decVat8 As Decimal = 0
        Dim decVat10 As Decimal = 0
        Dim decVatableNull As Decimal = 0
        Dim decVatable0 As Decimal = 0
        Dim decVatable5 As Decimal = 0
        Dim decVatable8 As Decimal = 0
        Dim decVatable10 As Decimal = 0

        Dim decTotalDiscount As Decimal = 0
        Dim decTotalVat As Decimal = 0
        Dim decTotalVatable As Decimal = 0
        Dim decTotalInv As Decimal = 0


        'If strPattern = "" Then
        '    If pblnTT78 Then
        '        strPattern = "1/001"
        '    Else
        '        strPattern = "01GTKT0/001"
        '    End If

        'End If
        'If strSerial = "" Then
        '    If pblnTT78 Then
        '        strSerial = "C" & Format(Now, "yy") & "TAA"
        '    Else
        '        strSerial = "AA/" & Format(Now, "yy") & "E"
        '    End If

        'End If

        'If strSerialCert = "" Then
        '    objXElement =
        '    <Invoices>
        '    </Invoices>
        'Else
        '    objXElement =
        '    <Invoices>
        '        <SerialCert><%= strSerialCert %></SerialCert>
        '    </Invoices>
        'End If

        'objInvoice =
        '        <AdjustInv>
        '            <key><%= strInvoiceKey %></key>
        '        </AdjustInv>

        objXProd =
                <Products>
                </Products>
        For Each objProduct As clsProduct In lstProducts
            objXProd.Add(CreateInvProduct(objProduct, True))
            Select Case objProduct.VatRate
                Case 10
                    decVatable10 = decVatable10 + objProduct.TotalPrice
                    decVat10 = decVat10 + objProduct.VatAmount
                Case 8
                    decVatable8 = decVatable8 + objProduct.TotalPrice
                    decVat8 = decVat8 + objProduct.VatAmount
                Case 5
                    decVatable5 = decVatable5 + objProduct.TotalPrice
                    decVat5 = decVat5 + objProduct.VatAmount
                Case 0
                    decVatable0 = decVatable0 + objProduct.TotalPrice
                    decVat0 = decVat0 + objProduct.VatAmount
                Case -1
                    decVatableNull = decVatableNull + objProduct.TotalPrice

            End Select
            decTotalDiscount = decTotalDiscount + objProduct.DiscountAmount
            decTotalVat = decTotalVat + objProduct.VatAmount
            decTotalVatable = decTotalVatable + objProduct.TotalPrice
            decTotalInv = decTotalInv + objProduct.Amount - objProduct.DiscountAmount
        Next

        If strPattern.EndsWith("1") Then
            decTotalInv = decTotalInv + decCharge
        Else
            decTotalInv = decTotalInv + decTax + decCharge
        End If

        objInvoice =
                <AdjustInv>
                    <key><%= strInvoiceKey %></key>
                    <Buyer><%= strBuyer %></Buyer>
                    <CusCode><%= strCustId %></CusCode>
                    <CusName><%= strCustFullName %></CusName>
                    <CusAddress><%= strAddress %></CusAddress>
                    <CusPhone><%= strPhone %></CusPhone>
                    <EmailDeliver><%= strEmail %></EmailDeliver>
                    <CusTaxCode><%= strTaxCode %></CusTaxCode>
                    <PaymentMethod><%= strFOP %></PaymentMethod>
                    <KindOfService><%= strKindOfService %></KindOfService>
                    <Total><%= decTotalVatable %></Total>
                    <DiscountAmount><%= decTotalDiscount %></DiscountAmount>
                    <VATAmount><%= decTotalVat %></VATAmount>
                    <VatAmount0><%= decVat0 %></VatAmount0>
                    <VatAmount5><%= decVat5 %></VatAmount5>
                    <VatAmount8><%= decVat8 %></VatAmount8>
                    <VatAmount10><%= decVat10 %></VatAmount10>
                    <GrossValue><%= decVatableNull %></GrossValue>
                    <GrossValue0><%= decVatable0 %></GrossValue0>
                    <GrossValue5><%= decVatable5 %></GrossValue5>
                    <GrossValue8><%= decVatable8 %></GrossValue8>
                    <GrossValue10><%= decVatable10 %></GrossValue10>
                    <Amount><%= decTotalInv %></Amount>
                    <AmountInWords><%= TienBangChu(Math.Abs(decTotalInv)) & " đồng" %></AmountInWords>
                    <Type><%= strAdjustType %></Type>
                </AdjustInv>


        'add thong tin hoan chinh vao day
        objInvoice.Add(objXProd)
        Return objInvoice
        'objInv.Add(objInvoice)
        'objXElement.Add(objInv)

        'Return objXElement
    End Function
    Public Function CreateInvDataAdjustNoInv(strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strSerialCert As String = "", Optional intVatDiscount As Integer = 0) As XElement

        'Dim objXElement As XElement
        'Dim objInv As XElement
        Dim objInvoice As XElement
        'Dim objVatDiscount As XElement

        Dim decVat0 As Decimal = 0
        Dim decVat5 As Decimal = 0
        Dim decVat7 As Decimal = 0
        Dim decVat8 As Decimal = 0
        Dim decVat10 As Decimal = 0
        Dim decVatableNull As Decimal = 0
        Dim decVatable0 As Decimal = 0
        Dim decVatable5 As Decimal = 0
        Dim decVatable8 As Decimal = 0
        Dim decVatable10 As Decimal = 0
        Dim decVatable7 As Decimal = 0
        Dim decTotalDiscount As Decimal = 0
        Dim decTotalVat As Decimal = 0
        Dim decTotalVatable As Decimal = 0
        Dim decTotalInv As Decimal = 0
        Dim blnVatPctNullFound As Boolean
        Dim blnVatPct0Found As Boolean
        Dim blnVatPct5Found As Boolean
        Dim blnVatPct7Found As Boolean
        Dim blnVatPct8Found As Boolean
        Dim blnVatPct10Found As Boolean

        Dim objDSHHDVu As XElement = <DSHHDVu></DSHHDVu>
        Dim objNDHD As XElement = <NDHDon>
                                      <NBan>
                                          <Ten>CÔNG TY TNHH DU LỊCH TRẦN VIỆT</Ten>
                                          <MST>0301069809</MST>
                                          <DChi>170-172 Nam Kỳ Khởi Nghĩa, Phường Võ Thị Sáu, Quận 3, TP Hồ Chí Minh, Việt Nam</DChi>
                                          <SDThoai>028 3933 0777</SDThoai>
                                      </NBan>
                                      <NMua>
                                          <Ten><%= strCustFullName %></Ten>
                                          <MST><%= strTaxCode %></MST>
                                          <DChi><%= strAddress %></DChi>
                                          <MKHang><%= strCustId %></MKHang>
                                          <SDThoai><%= strPhone %></SDThoai>
                                          <DCTDTu><%= strEmail %></DCTDTu>
                                          <HVTNMHang><%= strBuyer %></HVTNMHang>
                                      </NMua>
                                  </NDHDon>
        For Each objProduct As clsProduct In lstProducts
            objDSHHDVu.Add(CreateInvProductTt78(objProduct))
            Select Case objProduct.VatRate
                Case 10
                    decVatable10 = decVatable10 + objProduct.TotalPrice
                    decVat10 = decVat10 + objProduct.VatAmount
                    blnVatPct10Found = True
                Case 8
                    decVatable8 = decVatable8 + objProduct.TotalPrice
                    decVat8 = decVat8 + objProduct.VatAmount
                    blnVatPct8Found = True
                Case 5
                    decVatable5 = decVatable5 + objProduct.TotalPrice
                    decVat5 = decVat5 + objProduct.VatAmount
                    blnVatPct5Found = True
                Case 0
                    decVatable0 = decVatable0 + objProduct.TotalPrice
                    decVat0 = decVat0 + objProduct.VatAmount
                    blnVatPct0Found = True
                Case -1
                    decVatableNull = decVatableNull + objProduct.TotalPrice
                    blnVatPctNullFound = True
                Case 7
                    decVatable7 = decVatable7 + objProduct.TotalPrice
                    decVat7 = decVat7 + objProduct.VatAmount
                    blnVatPct7Found = True
            End Select
            decTotalDiscount = decTotalDiscount + objProduct.DiscountAmount
            decTotalVat = decTotalVat + objProduct.VatAmount
            decTotalVatable = decTotalVatable + objProduct.TotalPrice
            decTotalInv = decTotalInv + objProduct.Amount - objProduct.DiscountAmount
        Next
        objNDHD.Add(objDSHHDVu)


        If strPattern.EndsWith("1") Then
            decTotalInv = decTotalInv + decCharge
        Else
            decTotalInv = decTotalInv + decTax + decCharge
        End If
        If strAdjustType = "9" Then
            objInvoice =
                <ThayTheHD>
                    <key><%= strInvoiceKey %></key>
                    <InvoiceNo/>
                    <TTChung>
                        <MHSo/>
                        <SBKe/>
                        <NBKe/>
                        <DVTTe>VND</DVTTe>
                        <TGia>1</TGia>
                        <HTTToan><%= strFOP %></HTTToan>
                    </TTChung>
                </ThayTheHD>
        Else
            objInvoice =
                <DieuChinhHD>
                    <key><%= strInvoiceKey %></key>
                    <Type><%= strAdjustType %></Type>
                    <InvoiceNo/>
                    <TTChung>
                        <MHSo/>
                        <SBKe/>
                        <NBKe/>
                        <DVTTe>VND</DVTTe>
                        <TGia>1</TGia>
                        <HTTToan><%= strFOP %></HTTToan>
                    </TTChung>
                </DieuChinhHD>
        End If
        '<Type><%= strAdjustType %></Type>

        Dim objTToan As XElement = <TToan>
                                       <TgTCThue><%= decTotalVatable %></TgTCThue>
                                       <TgTThue><%= decTotalVat %></TgTThue>
                                       <TTCKTMai><%= decTotalDiscount %></TTCKTMai>
                                       <TgTTTBSo><%= decTotalInv %></TgTTTBSo>
                                       <TgTTTBChu><%= TienBangChu(Math.Abs(decTotalInv)) & " đồng" %></TgTTTBChu>
                                   </TToan>
        Dim objSumThueSuat As XElement = <THTTLTSuat></THTTLTSuat>

        If decVatableNull <> 0 Then
            objSumThueSuat.Add(CreateSumThueSuat(-1, 0, decVatableNull))
        End If
        If blnVatPct0Found Then
            objSumThueSuat.Add(CreateSumThueSuat(0, 0, decVatable0))

        End If
        If blnVatPct5Found Then
            objSumThueSuat.Add(CreateSumThueSuat(5, decVat5, decVatable5))
        End If
        If blnVatPct8Found Then
            objSumThueSuat.Add(CreateSumThueSuat(8, decVat8, decVatable8))
        End If
        If blnVatPct7Found Then
            objSumThueSuat.Add(CreateSumThueSuat(7, decVat7, decVatable7))
        End If
        If blnVatPct10Found Then
            objSumThueSuat.Add(CreateSumThueSuat(10, decVat0, decVatable10))
        End If
        objTToan.AddFirst(objSumThueSuat)

        objNDHD.Add(objTToan)

        'add thong tin hoan chinh vao day
        objInvoice.Add(objNDHD)

        Return objInvoice

    End Function
    Public Function CreateInvDataReplace(strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strSerialCert As String = "", Optional intVatDiscount As Integer = 0) As XElement

        Dim objInvoice As XElement
        Dim objXProd As XElement

        Dim decVat0 As Decimal = 0
        Dim decVat5 As Decimal = 0
        Dim decVat8 As Decimal = 0
        Dim decVat10 As Decimal = 0
        Dim decVatableNull As Decimal = 0
        Dim decVatable0 As Decimal = 0
        Dim decVatable5 As Decimal = 0
        Dim decVatable8 As Decimal = 0
        Dim decVatable10 As Decimal = 0

        Dim decTotalDiscount As Decimal = 0
        Dim decTotalVat As Decimal = 0
        Dim decTotalVatable As Decimal = 0
        Dim decTotalInv As Decimal = 0



        'If strSerialCert = "" Then
        '    objXElement =
        '    <Invoices>
        '    </Invoices>
        'Else
        '    objXElement =
        '    <Invoices>
        '        <SerialCert><%= strSerialCert %></SerialCert>
        '    </Invoices>
        'End If

        'objInvoice =
        '        <AdjustInv>
        '            <key><%= strInvoiceKey %></key>
        '        </AdjustInv>

        objXProd =
                <Products>
                </Products>
        For Each objProduct As clsProduct In lstProducts
            objXProd.Add(CreateInvProduct(objProduct, True))
            Select Case objProduct.VatRate
                Case 10
                    decVatable10 = decVatable10 + objProduct.TotalPrice
                    decVat10 = decVat10 + objProduct.VatAmount
                Case 8
                    decVatable8 = decVatable8 + objProduct.TotalPrice
                    decVat8 = decVat8 + objProduct.VatAmount
                Case 5
                    decVatable5 = decVatable5 + objProduct.TotalPrice
                    decVat5 = decVat5 + objProduct.VatAmount
                Case 0
                    decVatable0 = decVatable0 + objProduct.TotalPrice
                    decVat0 = decVat0 + objProduct.VatAmount
                Case -1
                    decVatableNull = decVatableNull + objProduct.TotalPrice

            End Select
            decTotalDiscount = decTotalDiscount + objProduct.DiscountAmount
            decTotalVat = decTotalVat + objProduct.VatAmount
            decTotalVatable = decTotalVatable + objProduct.TotalPrice
            decTotalInv = decTotalInv + objProduct.Amount - objProduct.DiscountAmount
        Next

        If strPattern.EndsWith("1") Then
            decTotalInv = decTotalInv + decCharge
        Else
            decTotalInv = decTotalInv + decTax + decCharge
        End If

        objInvoice =
                <ReplaceInv>
                    <key><%= strInvoiceKey %></key>
                    <Buyer><%= strBuyer %></Buyer>
                    <CusCode><%= strCustId %></CusCode>
                    <CusName><%= strCustFullName %></CusName>
                    <CusAddress><%= strAddress %></CusAddress>
                    <CusPhone><%= strPhone %></CusPhone>
                    <EmailDeliver><%= strEmail %></EmailDeliver>
                    <CusTaxCode><%= strTaxCode %></CusTaxCode>
                    <PaymentMethod><%= strFOP %></PaymentMethod>
                    <KindOfService><%= strKindOfService %></KindOfService>
                    <Total><%= Math.Abs(decTotalVatable) %></Total>
                    <DiscountAmount><%= decTotalDiscount %></DiscountAmount>
                    <VATAmount><%= Math.Abs(decTotalVat) %></VATAmount>
                    <VatAmount0><%= Math.Abs(decVat0) %></VatAmount0>
                    <VatAmount5><%= Math.Abs(decVat5) %></VatAmount5>
                    <VatAmount8><%= Math.Abs(decVat8) %></VatAmount8>
                    <VatAmount10><%= Math.Abs(decVat10) %></VatAmount10>
                    <GrossValue><%= decVatableNull %></GrossValue>
                    <GrossValue0><%= Math.Abs(decVatable0) %></GrossValue0>
                    <GrossValue5><%= Math.Abs(decVatable5) %></GrossValue5>
                    <GrossValue8><%= Math.Abs(decVatable8) %></GrossValue8>
                    <GrossValue10><%= Math.Abs(decVatable10) %></GrossValue10>
                    <Extra1><%= decTax %></Extra1>
                    <Extra2><%= decCharge %></Extra2>
                    <Amount><%= Math.Abs(decTotalInv) %></Amount>
                    <AmountInWords><%= TienBangChu(Math.Abs(decTotalInv)) & " đồng" %></AmountInWords>
                </ReplaceInv>


        'add thong tin hoan chinh vao day
        objInvoice.Add(objXProd)
        Return objInvoice

    End Function
    Private Function CreateSumThueSuat(intVatPct As Integer, decVatAmt As Decimal, decTotalAmt As Decimal) As XElement
        Dim objThueSuat As XElement = <LTSuat>
                                          <TThue><%= decVatAmt %></TThue>
                                          <ThTien><%= decTotalAmt %></ThTien>
                                      </LTSuat>
        Select Case intVatPct
            Case -2
                objThueSuat.AddFirst(<TSuat>KKKNT</TSuat>)
            Case -1
                objThueSuat.AddFirst(<TSuat>KCT</TSuat>)
            Case 7
                objThueSuat.AddFirst(<TSuat>KHAC:7%</TSuat>)
            Case Else
                objThueSuat.AddFirst(<TSuat><%= intVatPct %>%</TSuat>)
                'objThueSuat.AddFirst(<TSuat><%= intVatPct %></TSuat>)
        End Select
        Return objThueSuat
    End Function
    Private Function ConvertVatPct2ThueSuat(intVatPct As Integer) As String
        Select Case intVatPct
            Case -2
                Return "KKKNT"
            Case -1
                Return "KCT"
            Case 7
                Return "KHAC:7%"
            Case Else
                Return intVatPct & "%"
        End Select
    End Function
    Public Function CreateCustData(intCustId As Integer, strCustFullName As String, strCustShortName As String _
                        , strTaxCode As String, strAddress As String, strEmail As String _
                        , intCustType As Integer, Optional strPhone As String = "" _
                        , Optional strContact As String = "" _
                        , Optional strBankAccountName As String = "", Optional strBankName As String = "" _
                        , Optional strAccountNumber As String = "") As XElement

        Dim objCustData As XElement

        objCustData =
                <Customer>
                    <Name><%= strCustFullName %></Name>
                    <Code><%= intCustId %></Code>
                    <TaxCode><%= strTaxCode %></TaxCode>
                    <Address><%= strAddress %></Address>
                    <BankAccountName><%= strBankAccountName %></BankAccountName>
                    <BankName><%= strBankName %></BankName>
                    <BankNumber><%= strBankName %></BankNumber>
                    <Email><%= strEmail %></Email>
                    <Fax></Fax>
                    <Phone><%= strPhone %></Phone>
                    <ContactPerson><%= strContact %></ContactPerson>
                    <RepresentPerson><%= strCustShortName %></RepresentPerson>
                    <CusType><%= intCustType %></CusType>
                </Customer>
        Return objCustData
    End Function

    Public Function GetWsResponse(objXmlDoc As XmlDocument, strTagName As String) As String
        Dim strResponse As String = ""
        Dim colNodeList As XmlNodeList = objXmlDoc.GetElementsByTagName(strTagName)

        If colNodeList.Count > 0 Then
            strResponse = colNodeList(0).InnerText
        End If
        'For Each objC1 As XmlNode In objXmlDoc.ChildNodes
        '    MsgBox(objC1.Name)
        'Next
        Return strResponse
    End Function
    Public Function TranslateServiceName2Vietnamese(strServiceName As String) As String
        Dim strDesc As String = String.Empty
        Select Case strServiceName
            Case "Accommodations"
                strDesc = "Tiền phòng khách sạn"
            Case "Transfer"
                strDesc = "Tiền xe"
            Case "Meal"
                strDesc = "Tiền ăn"
            Case "Visa"
                strDesc = "Phí Visa"
            Case "Miscellaneous"
                strDesc = "Dịch vụ khác"
            Case "Bank Fee"
                strDesc = "Phí chuyển khoản"
            Case "Merchant Fee"
                strDesc = "Phí cà thẻ"
            Case "TransViet SVC Fee"
                strDesc = "Phí dịch vụ"
            Case "Conf.Room"
                strDesc = "Tiền phòng họp"
            Case Else
                MsgBox("Dich vu Non Air moi. Can yeu cau Khanhnm bổ sung:" & strServiceName)
        End Select
        Return strDesc
    End Function
    Public Function TranslateUnitName2Vietnamese(strServiceName As String, strUnitName As String) As String
        Dim strDesc As String = String.Empty
        Select Case strServiceName
            Case "Accommodations"
                'If strUnitName = "R/N" Then
                strDesc = "Đêm"
                'Else
                '    strDesc = "Phòng"
                'End If

            Case "Transfer"
                strDesc = "Xe"
            Case "Meal"
                strDesc = "Tiền ăn"
            Case "Visa"
                strDesc = "Visa"
            Case "Miscellaneous"
                'strDesc = "Dịch vụ"
            Case "Bank Fee"
                strDesc = "Lần"
            Case "Merchant Fee"
                strDesc = "Lần"
            Case "TransViet SVC Fee", "Conf.Room"
                strDesc = "Lần"
            Case Else
                MsgBox("Dich vu Non Air moi. Can yeu cau Khanhnm bổ sung:" & strServiceName)
        End Select
        Return strDesc
    End Function
    Public Function DefineDomIntByVatPct(intVatPct As Integer) As String
        Select Case intVatPct
            Case 5, 10, 7, 3.5, 8
                Return "DOM"
            Case Else
                Return "INT"
        End Select
    End Function
    Public Function Base64ToString(strBase64Coded) As String
        Return System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(strBase64Coded))
    End Function
    Public Function Base64ToFile(strBase64Coded, strPath) As Boolean
        IO.File.WriteAllBytes(strPath, Convert.FromBase64String(strBase64Coded))
        Return True
    End Function
    Public Function CreateInvToken(strPattern As String, strSerial As String, strInvoiceNo As String) As String
        Return Replace(strPattern & ";" & strSerial & ";" & strInvoiceNo, " ", "")
    End Function
    Public Function SignHashVNPT(ByVal strHash As String, ByVal strSerialCert As String) As String
        Dim my As X509Store = New X509Store(StoreName.My, StoreLocation.CurrentUser)
        my.Open(OpenFlags.[ReadOnly])
        Dim csp As RSACryptoServiceProvider = Nothing

        For Each cert As X509Certificate2 In my.Certificates
            'MsgBox(cert.Subject)
            If cert.SerialNumber = strSerialCert Then
                csp = CType(cert.PrivateKey, RSACryptoServiceProvider)
                Exit For
            End If
        Next

        If csp Is Nothing Then
            Throw New Exception("No valid Certificate was found")
        End If

        Dim sha1 As SHA1Managed = New SHA1Managed()
        Dim hash As Byte() = Convert.FromBase64String(strHash)
        Return Convert.ToBase64String(csp.SignHash(hash, CryptoConfig.MapNameToOID("SHA1")))
    End Function
    Public Function ReformatXml(strText As String) As String
        strText = Replace(strText, "&lt;", "<")
        strText = Replace(strText, "&gt;", ">")
        Return strText
    End Function
    Public Function GenerateInvoiceNoticeLine(intSeq As Integer, strMCQTCap As String, strMauSo As String, strKyHieu As String _
                                              , intInvoiceNo As Integer, dteInvDOI As Date, intTCTBao As Integer _
                                              , strReason As String, strFkey As String, intLoaiHoaDon As Integer) As XElement
        Dim objInv As XElement
        objInv = <HDon>
                     <STT><%= intSeq %></STT>
                     <MCQTCap><%= strMCQTCap %></MCQTCap>
                     <KHMSHDon><%= strMauSo %></KHMSHDon>
                     <KHHDon><%= strKyHieu %></KHHDon>
                     <SHDon><%= intInvoiceNo %></SHDon>
                     <Ngay><%= Format(dteInvDOI, "yyyy-MM-dd") %></Ngay>
                     <LADHDDT><%= intLoaiHoaDon %></LADHDDT>
                     <TCTBao><%= intTCTBao %></TCTBao>
                     <LDo><%= strReason %></LDo>
                     <Fkey><%= strFkey %></Fkey>
                 </HDon>
        Return objInv
        '<Ngay><%= Format(dteInvDOI, "yyyy-MM-dd") %></Ngay>
        '<Ngay><%= Format(dteInvDOI, "dd/MM/yyyy") %></Ngay>
    End Function
    '<LADHDDT></LADHDDT>


    Public Function GetInvHtmlByInvNo(strTvc As String, blnTt78 As Boolean, blnNoOriInv As Boolean _
                            , strMauSo As String, strKyHieu As String, intInvNo As Integer) As String
        Dim objE_Inv As New clsE_Invoice
        Dim objConnect As New clsE_InvConnect((pblnTT78 AndAlso Not blnNoOriInv), strTvc)
        Dim strInvoiceToken As String = CreateInvToken(strMauSo, strKyHieu, intInvNo)
        If objE_Inv.getInvViewNoPay(objConnect.PortalServiceUrl, objConnect.UserName, objConnect.UserPass _
                                       , strInvoiceToken) Then
            Return objE_Inv.ResponseDesc
        Else
            MsgBox("Unable to get E Invoice html!" & vbNewLine & objE_Inv.ResponseDesc)
            Return ""
        End If
    End Function
    Public Function GetInvHtmlByFkey(strTvc As String, blnTt78 As Boolean, blnNoOriInv As Boolean _
                            , strFkey As String) As String
        Dim objE_Inv As New clsE_Invoice
        Dim objConnect As New clsE_InvConnect((pblnTT78 AndAlso Not blnNoOriInv), strTvc)

        If objE_Inv.getNewInvViewFkey(objConnect.PortalServiceUrl, objConnect.UserName, objConnect.UserPass _
                                       , strFkey) Then
            Return objE_Inv.ResponseDesc
        Else
            MsgBox("Unable to get E Invoice html!" & vbNewLine & objE_Inv.ResponseDesc)
            Return ""
        End If
    End Function
    Public Function GetOriginalInvBoth(strTkno As String) As DataRow
        Dim objOriInv As DataRow
        objOriInv = GetOriginalInv(True, strTkno)
        If objOriInv Is Nothing Then
            objOriInv = GetOriginalInv(False, strTkno)
        End If
        Return objOriInv
    End Function
    Public Function GetOriginalInv(blnTt78 As Boolean, strTkno As String) As DataRow
        Dim tblOldTkts As DataTable
        Dim tblOldInv As DataTable
        Dim objOriInv As DataRow
        Dim strInvTable As String = "lib.dbo.E_inv78"
        Dim strInvDetailTable As String = "lib.dbo.E_invDetails78"
        Dim strOriFkey As String = String.Empty

        strTkno = Replace(strTkno, " ", "")

        If blnTt78 Then
            strInvTable = "lib.dbo.E_inv78"
            strInvDetailTable = "lib.dbo.E_invDetails78"
        Else
            strInvTable = "lib.dbo.E_inv"
            strInvDetailTable = "lib.dbo.E_invDetails"
        End If
        tblOldTkts = GetDataTable("select top 1 RcpId,RecId,Tkno from Tkt where SRV='S' and replace(Tkno,' ','')='" _
                                                     & strTkno & "' order by RecId desc", conn(0))
        tblOldInv = GetDataTable("select * from " & strInvDetailTable & " d" _
                                 & " left join " & strInvTable & " i on d.InvId=i.InvId" _
                                 & " where i.InvoiceNo<>0 and i.SRV='S' And (replace(Tkno,' ','')='" _
                                 & strTkno & "' or replace(Description,' ','') like '%" _
                                 & strTkno & "%') order by i.InvId desc", conn(0))

        Select Case tblOldInv.Rows.Count
            Case 0
                Return Nothing
            Case 1
                objOriInv = tblOldInv.Rows(0)
                strOriFkey = tblOldInv.Rows(0)("InvId")
                If blnTt78 Then

                Else

                End If



            Case Else

                Return Nothing
        End Select

        If blnTt78 AndAlso strOriFkey <> "" Then
            tblOldInv = GetDataTable("select  i.* from lib.dbo.E_invDetails78 d" _
                                 & " left join lib.dbo.E_Inv78 i on d.InvId=i.InvId" _
                                 & " where i.Srv='S' and i.InvId=" & strOriFkey, conn(0))
            Select Case tblOldInv.Rows.Count
                Case 0
                    Return Nothing
                Case 1
                    objOriInv = tblOldInv.Rows(0)
                    strOriFkey = tblOldInv.Rows(0)("InvId")
                Case Else
                    Return Nothing
            End Select
        End If
        Return objOriInv
    End Function
    Public Function AdjustInvoice(strTvc As String, IntInvRecId As Integer, intInvId As Integer, intAdjustType As Integer _
                                  , intCustId As Integer, strCustFullName As String _
                                  , strCustAddress As String, strTaxCode As String, strEmail As String _
                                  , strFop As String, strMauSo As String, strKyHieu As String _
                                  , strOriFkey As String, intOriInvNo As Integer, strOldPattern As String, strOldSerial As String _
                                  , dteOldDOI As Date, Optional blnViewOnly As Boolean = False) As Boolean
        Dim objE_InvConnect As New clsE_InvConnect(pblnTT78, strTvc)
        Dim objE_Invoice As New clsE_Invoice
        Dim lstProduct As New List(Of clsProduct)
        Dim strKindOfService As String = ""
        Dim strSerialCert As String = String.Empty
        Dim decTax As Decimal = 0
        Dim tblInvDetails As DataTable = GetDataTable("select * from lib.dbo.E_InvDetails78 where InvId=" _
                                                      & intInvId & " order by RecId", conn(0))
        For Each objRow As DataRow In tblInvDetails.Rows
            Dim objProd As New clsProduct
            With objRow
                objProd.IsSum = objRow("IsSum")
                objProd.ProdName = objRow("Tkno")
                objProd.Extra1 = objRow("Description")
                objProd.ProdUnit = objRow("Unit")

                If IsNumeric(objRow("Qty")) AndAlso objRow("Qty") <> 0 Then
                    objProd.ProdQuantity = objRow("Qty")
                End If
                objProd.TotalPrice = objRow("Amount")
                objProd.ProdPrice = objRow("Price")
                objProd.VatRate = objRow("VatPct")
                objProd.VatAmount = objRow("Vat")
                objProd.Amount = objRow("Total")

            End With
            lstProduct.Add(objProd)
        Next
        strKindOfService = KindOfService.Hóa_đơn_GTGT
        If blnViewOnly Then
            Dim blnViewOK As Boolean = False
            If objE_Invoice.AdjustInvoiceNoPublish(objE_InvConnect.BusinessServiceUrl, objE_InvConnect.UserName, objE_InvConnect.UserPass _
                                        , objE_InvConnect.AccountName, objE_InvConnect.AccountPass _
                                        , intCustId, strCustFullName _
                                        , strCustAddress, "", strTaxCode, strFop _
                                        , strOriFkey, strKindOfService, InvoiceType.Hóa_đơn_thông_thường, lstProduct _
                                        , intInvId, intAdjustType, strMauSo, strKyHieu, decTax _
                                        , 0, "", strEmail, strOldPattern) Then
                blnViewOK = True
            Else
                'MsgBox("Unable to view Invoice" & vbNewLine & objE_Invoice.ResponseDesc)
                Return False
            End If
            If blnViewOK Then
                Dim frmShow As New frmShowHtml(objE_Invoice.ResponseDesc)
                frmShow.ShowDialog()
            End If
            Return True
        End If

        If IsNumeric(strOriFkey) AndAlso CInt(strOriFkey) < 40000 Then
            If Not objE_Invoice.AdjustWithoutInv(objE_InvConnect.BusinessServiceUrl, objE_InvConnect.UserName, objE_InvConnect.UserPass _
                                        , objE_InvConnect.AccountName, objE_InvConnect.AccountPass _
                                        , intCustId, strCustFullName _
                                        , strCustAddress, "", strTaxCode, "CK" _
                                        , intInvId, strKindOfService, InvoiceType.Hóa_đơn_thông_thường, lstProduct _
                                        , strOriFkey, intAdjustType, strOldPattern, strOldSerial, intOriInvNo, dteOldDOI _
                                        , strMauSo, strKyHieu,,, "", strEmail) Then
                Return False
            End If
        Else
            If Not objE_Invoice.AdjustInvoiceAction(objE_InvConnect.BusinessServiceUrl, objE_InvConnect.UserName, objE_InvConnect.UserPass _
                                        , objE_InvConnect.AccountName, objE_InvConnect.AccountPass _
                                        , intCustId, strCustFullName _
                                        , strCustAddress, "", strTaxCode, "CK" _
                                        , intInvId, strKindOfService, InvoiceType.Hóa_đơn_thông_thường, lstProduct _
                                        , strOriFkey, intAdjustType, strMauSo, strKyHieu, decTax _
                                        , 0, "", strEmail) Then
                Return False
            End If
        End If


        If objE_Invoice.ReponseCode.StartsWith("OK") Then
            'Dim arrBreaks As String() = objE_Invoice.ReponseCode.Split("-")
            Dim arrMauSoKyHieu As String() = objE_Invoice.ReponseCode.Split(";")
            Dim arrKeyNbr As String() = arrMauSoKyHieu(2).Split("_")
            Dim intInvNo As Integer = 0
            If arrKeyNbr.Length = 2 Then
                intInvNo = arrKeyNbr(1)
            End If
            Dim lstQuerries As New List(Of String)
            lstQuerries.Add("Update lib.dbo.E_Inv78 set MauSo='" & Mid(arrMauSoKyHieu(0), 4) _
                    & "',KyHieu='" & arrMauSoKyHieu(1) & "',InvoiceNo=" & intInvNo _
                    & ",DOI=getdate() where Recid=" & IntInvRecId)

            If UpdateListOfQuerries(lstQuerries, conn_Web) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
        Return True
    End Function
    Public Function ApproveDraftInv(blnITT78 As Boolean, strTvc As String, strMauSo As String _
                                    , strKyHieu As String, strFkey As String) As Boolean
        Dim strInvoiceTable As String = "E_Inv"
        Dim objE_Inv As New clsE_Invoice
        Dim objConnect As New clsE_InvConnect(pblnTT78, strTvc)

        If blnITT78 Then strInvoiceTable = "E_Inv78"

        If strTvc = "APG" Then
            If Not objE_Inv.getHashInv(objConnect.WsUrl, objConnect.UserName, objConnect.UserPass _
                                           , objConnect.AccountName, objConnect.AccountPass _
                                           , pstrSerialCertAPG, strFkey, strMauSo) Then
                MsgBox("Unable to create E Invoice!" & vbNewLine & objE_Inv.ResponseDesc)
                Return False
            Else
                MsgBox(objE_Inv.InvToken.HashValue)
                Return False
            End If
        End If

        If Not objE_Inv.PublishInvFkey(objConnect.WsUrl, objConnect.AccountName, objConnect.AccountPass _
                                       , strFkey, objConnect.UserName, objConnect.UserPass _
                                       , strMauSo, strKyHieu) Then
            MsgBox("Unable to create E Invoice!" & vbNewLine & objE_Inv.ResponseDesc)
            Return False
        Else
            Dim arrResults As String() = objE_Inv.ResponseDesc.Split(",")
            For Each strInvoiceId_No As String In arrResults
                Dim arrBreaks As String() = strInvoiceId_No.Split("_")
                If ExecuteNonQuerry("Update " & strInvoiceTable & " set Draft='False', InvoiceNo=" & arrBreaks(1) _
                                        & " where Status='OK' and Draft='True' and InvId=" _
                                        & strFkey, conn(0)) Then
                    MsgBox("Created Invoice number:" & arrBreaks(1))
                    Return True
                Else
                    Dim strError As String = "Unable to Update E Invoice Number into Database!" & vbNewLine & objE_Inv.ResponseDesc
                    MsgBox(strError)
                    Append2TextFile(strError)
                End If
            Next
        End If
        Return True
    End Function
    Public Function DeleteDraftInv(blnTt78 As Boolean, strTvc As String, strFkey As String) As Boolean
        Dim objE_Inv As New clsE_Invoice
        Dim objConnect As New clsE_InvConnect(blnTt78, strTvc)
        Dim strInvoiceTable As String = "lib.dbo.E_Inv"
        Dim strInvDetailsTable As String = "lib.dbo.E_InvDetails"
        Dim strInvLinksTable As String = "lib.dbo.E_InvLinks"
        Dim lstQuerries As New List(Of String)
        If blnTt78 Then
            strInvoiceTable = "lib.dbo.E_Inv78"
            strInvDetailsTable = "lib.dbo.E_InvDetails78"
            strInvLinksTable = "lib.dbo.E_InvLinks78"
        End If
        'If .Cells("TVC").Value = "APG" Then
        '    If Not objE_Inv.getHashInv(objConnect.WsUrl, objConnect.UserName, objConnect.UserPass _
        '                               , objConnect.AccountName, objConnect.AccountPass _
        '                               , pstrSerialCertAPG, .Cells("InvId").Value, .Cells("MauSo").Value) Then
        '        MsgBox("Unable to create E Invoice!" & vbNewLine & objE_Inv.ResponseDesc)
        '    Else
        '        MsgBox(objE_Inv.InvToken.HashValue)
        '    End If
        'End If

        If Not objE_Inv.deleteInvoiceByFkey(objConnect.WsUrl, objConnect.UserName, objConnect.UserPass _
                                                , objConnect.AccountName, objConnect.AccountPass, strFkey) Then
            MsgBox("Unable to delete Draft E Invoice!" & vbNewLine & objE_Inv.ResponseDesc)
            Return False
        Else
            lstQuerries.Add("DELETE " & strInvoiceTable & " where Draft='true' and InvId=" & strFkey)
            lstQuerries.Add("DELETE " & strInvDetailsTable & " where InvId=" & strFkey)
            lstQuerries.Add("DELETE " & strInvLinksTable & " where InvId=" & strFkey)
            If UpdateListOfQuerries(lstQuerries, conn(0)) Then
                Return True
            Else
                Dim strError As String = "Unable to delete Draft E Invoice record in Database!"
                MsgBox(strError)
                Append2TextFile(strError)
                Return False
            End If
        End If
        Return True
    End Function

End Module
