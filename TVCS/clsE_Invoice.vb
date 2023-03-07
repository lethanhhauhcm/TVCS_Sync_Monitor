Imports System.Net
Imports System.IO
Imports System.Xml.Serialization

Imports System.Xml.Linq
Imports System.Xml


Public Class clsE_Invoice
    'Private mstrPublishServiceEndPoint As String '= "https://tranviethcmadmindemo.vnpt-invoice.com.vn/PublishService.asmx"
    'Private mstrPortalServiceEndPoint As String ' = "https://tranviethcmadmindemo.vnpt-invoice.com.vn/PortalService.asmx"
    'Private mstrBusinessServiceEndPoint As String ' = "https://tranviethcmadmindemo.vnpt-invoice.com.vn/BusinessService.asmx"
    Private mstrUserName As String '= "tranviethcmservice"
    Private mstrPass As String '= "123456aA@"
    Private mstrReponseCode As String
    Private mstrFuntionName As String
    Private mstrResponseDesc As String
    Private mstrLastResponse As String
    Private mstrLastRequest As String
    Private mobjXmlDoc As New Xml.XmlDocument
    Private mobjHttp As New MSXML2.XMLHTTP60
    Private mobjInvToken As New clsInvToken
    Private mobjCertInfo As New clsCertInfo
    Public Function ResetProperties(strFunctionName As String) As Boolean
        mstrLastRequest = ""
        mstrReponseCode = ""
        mstrResponseDesc = ""
        mstrFuntionName = strFunctionName

        Return True
    End Function
    Public Function ConvertResponseCode2Desc(strFunctionName As String, strResponseCode As String) As String
        Dim strDesc As String = ""
        Select Case strFunctionName
            Case FunctionName.AdjustActionAssignedNo
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Hóa đơn cần điều chỉnh không tồn tại"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Không phát hành được hóa đơn"
                    Case "ERR:6"
                        strDesc = "Dải hóa đơn cũ đã hết"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user."
                    Case "ERR:8"
                        strDesc = "Hóa đơn cần điều chỉnh đã bị thay thế. Không thể điều chỉnh được nữa."
                    Case "ERR:13"
                        strDesc = "Lỗi trùng fkey"
                    Case "ERR:14"
                        strDesc = "Lỗi trong quá trình thực hiện cấp số hóa đơn"
                    Case "ERR:15"
                        strDesc = "Lỗi khi thực hiện Deserialize chuỗi hóa đơn đầu vào"
                    Case "ERR:19"
                        strDesc = "Pattern truyền vào không giống với hóa đơn cần điều chỉnh"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và serial không phù hợp"
                    Case "ERR:29"
                        strDesc = "Lỗi chứng thư hết hạn"
                    Case "ERR:30"
                        strDesc = "Danh sách hóa đơn tồn tại ngày hóa đơn nhỏ hơn ngày hóa đơn đã phát hành"
                    Case "ERR:31"
                        strDesc = "Số hóa đơn truyền vào không hợp lệ"
                    Case Else
                        If strResponseCode.StartsWith("ERR:9") Then
                            strDesc = "Trạng thái hóa đơn không được điều chỉnh" & vbNewLine & strResponseCode
                        Else
                            strDesc = strResponseCode
                        End If



                End Select
            Case FunctionName.AdjustInvoiceAction
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Hóa đơn cần điều chỉnh không tồn tại"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Không phát hành được hóa đơn"
                    Case "ERR:6"
                        strDesc = "Không còn đủ số lượng hóa đơn để phát hành"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user."
                    Case "ERR:8"
                        strDesc = "Hóa đơn cần điều chỉnh đã bị thay thế. Không thể điều chỉnh được nữa."
                    Case "ERR:9"
                        strDesc = "Trạng thái hóa đơn không được điều chỉnh"
                    Case "ERR:13"
                        strDesc = "Lỗi trùng fkey"
                    Case "ERR:14"
                        strDesc = "Lỗi trong quá trình thực hiện cấp số hóa đơn"
                    Case "ERR:15"
                        strDesc = "Lỗi khi thực hiện Deserialize chuỗi hóa đơn đầu vào"
                    Case "ERR:19"
                        strDesc = "Pattern truyền vào không giống với hóa đơn cần điều chỉnh"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và serial không phù hợp"
                    Case "ERR:21"
                        strDesc = "Trùng Fkey truyền vào"
                    Case "ERR:29"
                        strDesc = "Lỗi chứng thư hết hạn"
                    Case "ERR:30"
                        strDesc = "Danh sách hóa đơn tồn tại ngày hóa đơn nhỏ hơn ngày hóa đơn đã phát hành"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.AdjustInvoiceNoPublish
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Hóa đơn cần điều chỉnh không tồn tại"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Có lỗi trong quá trình tạo mới hóa đơn điều chỉnh"
                    Case "ERR:6"
                        strDesc = "Dải hóa đơn cũ đã hết"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user."
                    Case "ERR:8"
                        strDesc = "Hóa đơn cần điều chỉnh đã bị thay thế. Không thể điều chỉnh được nữa."
                    Case "ERR:9"
                        strDesc = "Trạng thái hóa đơn không được điều chỉnh"
                    Case "ERR:15"
                        strDesc = "Lỗi khi thực hiện Deserialize chuỗi hóa đơn đầu vào"
                    Case "ERR:19"
                        strDesc = "Pattern truyền vào không giống với hóa đơn cần điều chỉnh"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và serial không phù hợp"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.AdjustInvoiceNote
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:4"
                        strDesc = "Không lấy được công ty"
                    Case "ERR:5"
                        strDesc = "Lỗi không xác định. Exception, kiểm tra log"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:20"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user."
                    Case "ERR:25"
                        strDesc = "Không tạo được invoice service"
                    Case "ERR:29"
                        strDesc = "Lỗi chứng thư hết hạn"
                    Case "ERR:55"
                        strDesc = "Pattern truyền vào không giống với pattern hóa đơn cần điều chỉnh"
                    Case "ERR:61"
                        strDesc = "Fkey hóa đơn bị lỗi có giá trị rỗng hoặc null"
                    Case "ERR:62"
                        strDesc = "Không lấy được nội dung điều chỉnh hóa đơn trong xml"
                    Case Else
                        If strResponseCode.StartsWith("ERR:56") Then
                            strDesc = "Trạng thái hóa đơn không hợp lệ. " & strResponseCode
                        Else
                            strDesc = strResponseCode
                        End If


                End Select
            Case FunctionName.AdjustReplaceInvWithToken
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:21"
                        strDesc = "Không tìm thấy công ty hoặc tài khoản không tồn tại"
                    Case "ERR:22"
                        strDesc = "Công ty chưa đăng ký chứng thư"
                    Case "ERR:28"
                        strDesc = "Chưa có thông tin chứng thư trong hệ thống"
                    Case "ERR:24"
                        strDesc = "Chứng thư truyền lên không đúng với chứng thư đăng ký trong hệ thống"
                    Case "ERR:27"
                        strDesc = "Chứng thư chưa đến thời điểm sử dụng"
                    Case "ERR:26"
                        strDesc = "Chứng thư hết hạn"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:20"
                        strDesc = "Không tìm thấy dải hóa đơn"
                    Case "ERR:6"
                        strDesc = "Không còn đủ số lượng hóa đơn để phát hành"
                    Case "ERR:10"
                        strDesc = "Lô có số hóa đơn vượt quá max cho phép"
                    Case "ERR:5"
                        strDesc = "Có lỗi xảy ra (Lỗi không xác định)"
                    Case "ERR:30"
                        strDesc = "Tạo mới hóa đơn có lỗi"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.AdjustWithoutInv
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Hóa đơn cần điều chỉnh không tồn tại"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Không phát hành được hóa đơn hoặc lỗi hệ thống"
                    Case "ERR:6"
                        strDesc = "Dải hóa đơn cũ đã hết"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user."
                    Case "ERR:12"
                        strDesc = "Ngày hóa đơn cũ không hợp lệ"
                    Case "ERR:14"
                        strDesc = "Lỗi trong quá trình thực hiện cấp số hóa đơn"
                    Case "ERR:15"
                        strDesc = "Lỗi khi thực hiện Deserialize chuỗi hóa đơn đầu vào"
                    Case "ERR:19"
                        strDesc = "Pattern truyền vào không giống với hóa đơn cần điều chỉnh"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và serial không phù hợp"
                    Case "ERR:29"
                        strDesc = "Lỗi chứng thư hết hạn"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.cancelInv
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Không tìm thấy hóa đơn"
                    Case "ERR:6"
                        strDesc = "Lỗi không xác định"
                    Case "ERR:7"
                        strDesc = "Không tìm thấy công ty hoặc tài khoản không tồn tại"
                    Case "ERR:8"
                        strDesc = "Hóa đơn đã bị điều chỉnh / hủy / hóa đơn mới tạo không thể hủy được"
                    Case "ERR:9"
                        strDesc = "Hóa đơn đã thanh toán, không cho phép hủy"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và serial không phù hợp Kiểu "
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.cancelInvNoPay
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:4"
                        strDesc = "Không tìm thấy công ty"
                    Case "ERR:7"
                        strDesc = "Tài khoản Account không tồn tại"
                    Case "ERR:18"
                        strDesc = "Không tồn tại hóa đơn này"
                    Case "ERR:20"
                        strDesc = "Không tìm thấy TBPH(mẫu hóa đơn)"
                    Case "ERR:29"
                        strDesc = "Hóa đơn là hđ mới tạo, hoặc bị điều chỉnh, hoặc đã đc hủy, nên không được phép hủy"
                    Case "ERR:30"
                        strDesc = "Có lỗi xảy ra khi hủy hóa đơn"
                    Case Else
                        strDesc = strResponseCode
                End Select

            Case FunctionName.deleteInvoiceByFkey
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:5"
                        strDesc = "Lỗi không xác định"
                    Case "ERR:7"
                        strDesc = "Không tìm thấy công ty"
                    Case "ERR:10"
                        strDesc = "Số hóa đơn truyền vào vượt quá số lượng cho phép"
                    Case "ERR:20"
                        strDesc = "Pattern và serial không phù hợp, hoặc không tồn tại hóa đơn đã đăng kí có sử dụng Pattern và serial truyền vào"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.deleteInvoiceByID
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:5"
                        strDesc = "Có lỗi xảy ra"
                    Case "ERR:20"
                        strDesc = "Pattern và serial không phù hợp, hoặc không tồn tại hóa đơn đã đăng kí có sử dụng Pattern và serial truyền vào"
                    Case Else
                        strDesc = strResponseCode
                End Select

            Case FunctionName.downloadInv, FunctionName.downloadInvFkey, FunctionName.downloadInvNoPay _
                , FunctionName.downloadInvPDF, FunctionName.downloadInvPDFNoPay, FunctionName.getInvView
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Chuỗi Invoice Token không chính xác"
                    Case "ERR:4"
                        strDesc = "Không tìm thấy Pattern"
                    Case "ERR:6"
                        strDesc = "Không tìm thấy hóa đơn (Chưa được phân quyền Serial hoặc hóa đơn không tồn tại)"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy thông tin công ty tương ứng cho user."
                    Case "ERR:11"
                        strDesc = "Chuỗi Invoice token đúng định dạng nhưng không tồn tại, hoặc là của hóa đơn đã bị hủy, bị thay thế, hoặc hóa đơn chưa thanh toán"
                    Case "ERR:12"
                        strDesc = "Hoá đơn có mã chưa được thuế chấp nhận"
                    Case "ERR:13"
                        strDesc = "Hoá đơn không mã chưa được thuế chấp nhận"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.getInvViewNoPay, FunctionName.getInvViewFkeyNoPay _
                , FunctionName.getNewInvViewFkey, FunctionName.GetInvByFkey
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Chuỗi Invoice Token không chính xác/Fkey rỗng"
                    Case "ERR:4"
                        strDesc = "Công ty chưa được đăng kí mẫu hóa đơn nào"
                    Case "ERR:6"
                        strDesc = "Không tìm thấy hóa đơn"
                    Case "ERR:7"
                        strDesc = "Không tìm thấy thông tin công ty"
                    Case "ERR:11"
                        strDesc = "Hóa đơn chưa thanh toán nên không xem được"
                    Case "ERR:12"
                        strDesc = "Hoá đơn có mã chưa được thuế chấp nhận"
                    Case "ERR:13"
                        strDesc = "Hoá đơn không mã chưa được thuế chấp nhận"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.GetCertInfo
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:21"
                        strDesc = "Không tìm thấy công ty hoặc tài khoản không tồn tại"
                    Case "ERR:22"
                        strDesc = "Công ty chưa đăng ký chứng thư số"
                    Case "ERR:28"
                        strDesc = "Chưa có thông tin chứng thư trong hệ thống"
                    Case "ERR:5"
                        strDesc = "Có lỗi xảy ra (Lỗi không xác định)"
                    Case Else
                        strDesc = strResponseCode
                End Select

            Case FunctionName.getHashInv, FunctionName.getHashInvWithToken, FunctionName.publishInvWithToken, FunctionName.CancelInvoiceWithToken
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:21"
                        strDesc = "Không tìm thấy công ty hoặc tài khoản không tồn tại"
                    Case "ERR:22"
                        strDesc = "Công ty chưa đăng ký chứng thư số"
                    Case "ERR:28"
                        strDesc = "Chưa có thông tin chứng thư trong hệ thống"
                    Case "ERR:24"
                        strDesc = "Chứng thư truyền lên không đúng với chứng thư đăng ký trong hệ thống"
                    Case "ERR:27"
                        strDesc = "Chứng thư chưa đến thời điểm sử dụng"
                    Case "ERR:26"
                        strDesc = "Chứng thư số hết hạn"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và Serial không phù hợp"
                    Case "ERR:6"
                        strDesc = "Không còn đủ số lượng hóa đơn để phát hành"
                    Case "ERR:10"
                        strDesc = "Lô có số hóa đơn vượt quá số lượng tối đa cho phép"
                    Case "ERR:5"
                        strDesc = "Lỗi không xác định"
                    Case "ERR:30"
                        strDesc = "Tạo mới hóa đơn có lỗi"
                    Case Else
                        strDesc = strResponseCode
                End Select

            Case FunctionName.HandleInvoiceErrors
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Mã thông điệp không họp lệ, không tìm thấy bản ghi transaction"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:4"
                        strDesc = "Chưa có kêt quả thuế trả về, trạng thái chi tiết chưa được cập nhật"
                    Case "ERR:5"
                        strDesc = "Có lỗi xảy ra"
                    Case "ERR:6"
                        strDesc = "Có lỗi xảy ra trong quá trình update trạng thái hóa đơn sai sót"
                    Case "ERR:7"
                        strDesc = "Không tìm thấy chi tiết hóa đơn sai sót"
                    Case "ERR:8"
                        strDesc = "Không tìm thấy danh sách hóa đơn hủy để gửi mẫu 04, không có hóa đơn thuế từ chối"
                    Case "ERR:21"
                        strDesc = "Không tìm thấy công ty hoặc tài khoản không tồn tại"
                    Case "ERR:22"
                        strDesc = "Công ty chưa đăng ký chứng thư số"
                    Case "ERR:24"
                        strDesc = "Chứng thư truyền lên không đúng với chứng thư đăng ký trong hệ thống"
                    Case "ERR:26"
                        strDesc = "Chứng thư số hết hạn"
                    Case "ERR:27"
                        strDesc = "Chứng thư chưa đến thời điểm sử dụng"
                    Case "ERR:28"
                        strDesc = "Chưa có thông tin chứng thư trong hệ thống"
                    Case "ERR:51"
                        strDesc = "Verify chứng thư lỗi"
                    Case Else
                        strDesc = strResponseCode
                End Select

            Case FunctionName.listInvFromNoToNo
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:7"
                        strDesc = "Không tìm thấy thông tin công ty"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.ReplaceInvoiceAction, FunctionName.ReplaceInvoiceNoPublish
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Không tồn tại hóa đơn cần thay thế"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Có lỗi trong quá trình thay thế hóa đơn"
                    Case "ERR:6"
                        strDesc = "Dải hóa đơn cũ đã hết"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user."
                    Case "ERR:8"
                        strDesc = "Hóa đơn đã được thay thế rồi. Không thể thay thế nữa"
                    Case "ERR:9"
                        strDesc = "Trạng thái hóa đơn không được thay thế"
                    Case "ERR:13"
                        strDesc = "Lỗi trùng fkey"
                    Case "ERR:14"
                        strDesc = "Lỗi trong quá trình thực hiện cấp số hóa đơn"
                    Case "ERR:15"
                        strDesc = "Lỗi khi thực hiện Deserialize chuỗi hóa đơn đầu vào"
                    Case "ERR:19"
                        strDesc = "Pattern truyền vào không giống với hóa đơn cần điều chỉnh"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và serial không phù hợp"
                    Case "ERR:21"
                        strDesc = "Trùng Fkey truyền vào"
                    Case "ERR:29"
                        strDesc = "Lỗi chứng thư hết hạn"
                    Case "ERR:30"
                        strDesc = "Danh sách hóa đơn tồn tại ngày hóa đơn nhỏ hơn ngày hóa đơn đã phát hành"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.ReplaceWithoutInv
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Không tồn tại hóa đơn cần thay thế"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Có lỗi trong quá trình thay thế hóa đơn"
                    Case "ERR:6"
                        strDesc = "Dải hóa đơn cũ đã hết"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user."
                    Case "ERR:8"
                        strDesc = "Hóa đơn đã được thay thế rồi. Không thể thay thế nữa"
                    Case "ERR:9"
                        strDesc = "Trạng thái hóa đơn không được thay thế"
                    Case "ERR:12"
                        strDesc = "Ngày hóa đơn cũ không hợp lệ"
                    Case "ERR:13"
                        strDesc = "Lỗi trùng fkey"
                    Case "ERR:14"
                        strDesc = "Lỗi trong quá trình thực hiện cấp số hóa đơn"
                    Case "ERR:15"
                        strDesc = "Lỗi khi thực hiện Deserialize chuỗi hóa đơn đầu vào"
                    Case "ERR:19"
                        strDesc = "Pattern truyền vào không giống với hóa đơn cần điều chỉnh"
                    Case "ERR:20"
                        strDesc = "Dải hóa đơn hết, User/Account không có quyền với Serial/Pattern và serial không phù hợp"
                    Case "ERR:21"
                        strDesc = "Trùng Fkey truyền vào"
                    Case "ERR:29"
                        strDesc = "Lỗi chứng thư hết hạn"
                    Case "ERR:30"
                        strDesc = "Danh sách hóa đơn tồn tại ngày hóa đơn nhỏ hơn ngày hóa đơn đã phát hành"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.SendInvNoticeErrors
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Có lỗi xảy ra"
                    Case "ERR:10"
                        strDesc = "Lô có số hóa đơn vượt quá số lượng tối đa cho phép"
                    Case "ERR:21"
                        strDesc = "Không tìm thấy công ty hoặc tài khoản không tồn tại"
                    Case "ERR:22"
                        strDesc = "Công ty chưa đăng ký chứng thư số"
                    Case "ERR:28"
                        strDesc = "Chưa có thông tin chứng thư trong hệ thống"
                    Case "ERR:24"
                        strDesc = "Chứng thư truyền lên không đúng với chứng thư đăng ký trong hệ thống"
                    Case "ERR:26"
                        strDesc = "Chứng thư số hết hạn"
                    Case "ERR:27"
                        strDesc = "Chứng thư chưa đến thời điểm sử dụng"
                    Case Else
                        strDesc = strResponseCode
                End Select
            Case FunctionName.UpdateCus
                Select Case strResponseCode
                    Case -1
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case -2
                        strDesc = "Không import được khách hàng vào db"
                    Case -3
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case -5
                        strDesc = "Có khách hàng đã tồn tại"
                    Case > 0
                        strDesc = strResponseCode & "khách hàng đã import và update"
                    Case Else
                        strDesc = "Lỗi không xác định"
                End Select
            Case FunctionName.ImportAndPublishInv, FunctionName.ImportInvByPattern, FunctionName.ImportInv
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:3"
                        strDesc = "Dữ liệu xml đầu vào không đúng quy định"
                    Case "ERR:5"
                        strDesc = "Không phát hành được hóa đơn"
                    Case "ERR:6"
                        strDesc = "Không đủ số hóa đơn cho lô phát hành"
                    Case "ERR:7"
                        strDesc = "User name không phù hợp, không tìm thấy company tương ứng cho user"
                    Case "ERR:10"
                        strDesc = "Lô có số hóa đơn vượt quá max cho phép "
                    Case "ERR:13"
                        strDesc = "Trùng Invoice ID"
                    Case "ERR:20"
                        strDesc = "Pattern và serial không phù hợp, hoặc không tồn tại hóa đơn đã đăng kí có sử dụng Pattern và serial truyền vào"
                    Case "ERR:28"
                        strDesc = "Chưa có thông tin chứng thư trong hệ thống"
                    Case Else
                        strDesc = strResponseCode
                End Select

            Case FunctionName.PublishInvFkey
                Dim arrErrors As String() = Split(strResponseCode, "||")
                Dim arrErrorBreak As String()
                For Each strError As String In arrErrors
                    arrErrorBreak = Split(strError, "#")
                    If arrErrorBreak(1) <> "" Then
                        Select Case arrErrorBreak(0)
                            Case "ERR:1"
                                strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                            Case "ERR:5"
                                strDesc = "Không phát hành được hóa đơn"
                            Case "ERR:6"
                                strDesc = "Danh sách Invoice ID không tồn tại"
                            Case "ERR:10"
                                strDesc = "Vượt quá 200 Invoice ID"
                            Case "ERR:15"
                                strDesc = "Danh sách Invoice ID đã phát hành"
                            Case "ERR:20"
                                strDesc = "Pattern và serial không phù hợp, hoặc không tồn tại hóa đơn đã đăng kí có sử dụng Pattern và serial truyền vào"
                            Case "ERR:28"
                                strDesc = "Chưa có thông tin chứng thư trong hệ thống"
                            Case Else
                                strDesc = strError
                        End Select
                        'chi lay loi dau tien
                        Exit For
                    End If

                Next

            Case FunctionName.ReceivedInvoiceErrors
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:2"
                        strDesc = "Mã thông điệp không họp lệ, không tìm thấy bản ghi transaction"
                    Case "ERR:4"
                        strDesc = "Chưa có kêt quả thuế trả về, trạng thái chi tiết chưa được cập nhật"
                    Case "ERR:5"
                        strDesc = "Có lỗi xảy ra"
                    Case "ERR:7"
                        strDesc = "Không tìm thấy chi tiết hóa đơn sai sót"
                    Case "ERR:8"
                        strDesc = "Lỗi nhận kết quả từ cơ quan thuế"
                    Case Else
                        strDesc = strResponseCode
                End Select

            Case FunctionName.SendAgainEmailServ
                Select Case strResponseCode
                    Case "ERR:1"
                        strDesc = "Tài khoản đăng nhập sai hoặc không có quyền"
                    Case "ERR:3"
                        strDesc = "Thiếu Mẫu số hóa đơn"
                    Case "ERR:4"
                        strDesc = "Không tìm thấy hóa đơn để tạo và gửi lại email"
                    Case "ERR:5", "ERR:7"
                        strDesc = "Có lỗi xảy ra, vui lòng thực hiện lại"
                    Case "ERR:6"
                        strDesc = "Không tìm thấy email để gửi lại"
                    Case "ERR:8"
                        strDesc = "Số hóa đơn trống"
                    Case "ERR:21"
                        strDesc = "Không tồn tại Khách hàng trên hệ thống"
                    Case Else
                        strDesc = strResponseCode
                End Select
        End Select

        mstrResponseDesc = strDesc
        Return strDesc
    End Function
    Public Function cancelInv(strWsBusinessService As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strFkey As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.cancelInv)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <cancelInv xmlns="http://tempuri.org/">
                                 <Account><%= strStaffAccount %></Account>
                                 <ACpass><%= strStaffPass %></ACpass>
                                 <fkey><%= strFkey %></fkey>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </cancelInv>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strWsBusinessService, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/cancelInv")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsBusinessService & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If
        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "cancelInvResponse")

        ConvertResponseCode2Desc(FunctionName.cancelInv, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    Public Function cancelInvNoPay(strWsBusinessService As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strFkey As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.cancelInvNoPay)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <cancelInvNoPay xmlns="http://tempuri.org/">
                                 <Account><%= strStaffAccount %></Account>
                                 <ACpass><%= strStaffPass %></ACpass>
                                 <fkey><%= strFkey %></fkey>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </cancelInvNoPay>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strWsBusinessService, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/cancelInvNoPay")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "cancelInvNoPayResponse")

        ConvertResponseCode2Desc(FunctionName.cancelInvNoPay, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function downloadInv(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strFkey As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.downloadInv)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <downloadInv xmlns="http://tempuri.org/">
                                 <invToken><%= strFkey %></invToken>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </downloadInv>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/downloadInv")

        'Send the SOAP request
        strSoapBody = objRequest.ToString

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "downloadInvResponse")

        ConvertResponseCode2Desc(FunctionName.downloadInv, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function downloadInvFkey(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strFkey As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.downloadInvFkey)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <downloadInvFkey xmlns="http://tempuri.org/">
                                 <fkey><%= strFkey %></fkey>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </downloadInvFkey>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/downloadInvFkey")

        'Send the SOAP request
        strSoapBody = objRequest.ToString

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "downloadInvFkeyResponse")

        ConvertResponseCode2Desc(FunctionName.downloadInvFkey, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    Public Function downloadInvFkeyNoPay(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strFkey As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.downloadInvFkeyNoPay)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <downloadInvFkeyNoPay xmlns="http://tempuri.org/">
                                 <fkey><%= strFkey %></fkey>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </downloadInvFkeyNoPay>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/downloadInvFkeyNoPay")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody
        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" _
                                & vbNewLine & ReformatXml(mstrLastRequest))
        End If
        objHttp.send(strSoapBody)

        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If
        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "downloadInvFkeyNoPayResponse")

        ConvertResponseCode2Desc(FunctionName.downloadInvFkeyNoPay, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" _
                                & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    Public Function downloadInvNoPay(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String
        Dim objRequest As XElement

        ResetProperties(FunctionName.downloadInvNoPay)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <downloadInvNoPay xmlns="http://tempuri.org/">
                                 <invToken><%= strInvoiceToken %></invToken>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </downloadInvNoPay>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/downloadInvNoPay")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody
        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "downloadInvNoPayResponse")

        ConvertResponseCode2Desc(FunctionName.downloadInvNoPay, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function GetInvbyFkey(strBusinessServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.GetInvByFkey)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <GetInvbyFkey xmlns="http://tempuri.org/">
                                 <fkey><%= strInvoiceToken %></fkey>
                                 <username><%= strWsUser %></username>
                                 <pass><%= strWsPass %></pass>
                             </GetInvbyFkey>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strBusinessServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/GetInvbyFkey")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody
        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strBusinessServiceUrl & vbNewLine _
                            & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If
        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "GetInvbyFkeyResponse")

        ConvertResponseCode2Desc(FunctionName.GetInvByFkey, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strBusinessServiceUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    Public Function getInvView(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.getInvView)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <getInvView xmlns="http://tempuri.org/">
                                 <invToken><%= strInvoiceToken %></invToken>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </getInvView>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/getInvView")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody
        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If
        mobjXmlDoc.LoadXml(mstrLastResponse)
        mstrReponseCode = GetWsResponse(mobjXmlDoc, "getInvViewResult")

        ConvertResponseCode2Desc(FunctionName.getInvView, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function getInvViewNoPay(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.getInvViewNoPay)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <getInvViewNoPay xmlns="http://tempuri.org/">
                                 <invToken><%= strInvoiceToken %></invToken>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </getInvViewNoPay>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/getInvViewNoPay")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "getInvViewNoPayResponse")

        ConvertResponseCode2Desc(FunctionName.getInvViewNoPay, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function getInvViewFkeyNoPay(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.getInvViewFkeyNoPay)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <getInvViewFkeyNoPay xmlns="http://tempuri.org/">
                                 <fkey><%= strInvoiceToken %></fkey>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </getInvViewFkeyNoPay>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/getInvViewFkeyNoPay")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "getInvViewFkeyNoPayResponse")

        ConvertResponseCode2Desc(FunctionName.getInvViewFkeyNoPay, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    Public Function getNewInvViewFkey(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.getNewInvViewFkey)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <getNewInvViewFkey xmlns="http://tempuri.org/">
                                 <fkey><%= strInvoiceToken %></fkey>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </getNewInvViewFkey>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/getNewInvViewFkey")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody
        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine _
                            & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If
        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "getNewInvViewFkeyResponse")

        ConvertResponseCode2Desc(FunctionName.getNewInvViewFkey, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    Public Function downloadInvPDF(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.downloadInvPDF)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <downloadInvPDF xmlns="http://tempuri.org/">
                                 <token><%= strInvoiceToken %></token>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </downloadInvPDF>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/downloadInvPDF")

        'Send the SOAP request
        strSoapBody = objRequest.ToString

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "downloadInvPDFResponse")

        ConvertResponseCode2Desc(FunctionName.downloadInvPDF, Base64ToString(mstrReponseCode))
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function downloadInvPDFNoPay(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strInvoiceToken As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.downloadInvPDFNoPay)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <downloadInvPDFNoPay xmlns="http://tempuri.org/">
                                 <token><%= strInvoiceToken %></token>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </downloadInvPDFNoPay>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/downloadInvPDFNoPay")

        'Send the SOAP request
        strSoapBody = objRequest.ToString

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "downloadInvPDFNoPayResponse")

        ConvertResponseCode2Desc(FunctionName.downloadInvPDFNoPay, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            MsgBox(Base64ToString(mstrReponseCode))
            Return True
        Else

            Return False
        End If

    End Function
    Public Function downloadNewInvPDFFkey(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strFkey As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.downloadNewInvPDFFkey)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <downloadNewInvPDFFkey xmlns="http://tempuri.org/">
                                 <fkey><%= strFkey %></fkey>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </downloadNewInvPDFFkey>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/downloadNewInvPDFFkey")

        'Send the SOAP request
        strSoapBody = objRequest.ToString

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "downloadNewInvPDFFkeyResponse")

        ConvertResponseCode2Desc(FunctionName.downloadNewInvPDFFkey, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            'MsgBox(Base64ToString(mstrReponseCode))
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" _
                                & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    'Public Function PostDataPublishService(strFunctionName As String) As Boolean

    '    ResetProperties(strFunctionName)

    '    Return True
    'End Function
    'Public Function UpdateCus(intCustId As Integer, strCustFullName As String, strCustShortName As String _
    '                            , strTaxCode As String, strAddress As String, strEmail As String _
    '                            , intCustType As Integer, Optional strPhone As String = "" _
    '                            , Optional strContact As String = "" _
    '                            , Optional strBankAccountName As String = "", Optional strBankName As String = "" _
    '                            , Optional strAccountNumber As String = "") As Boolean
    '    Dim objRequest As New UpdateCusRequest
    '    Dim objContent As New UpdateCusRequestBody
    '    Dim objXElement As XElement
    '    Dim objResponse As UpdateCusResponse
    '    Dim objSvc As New PublishServiceSoapClient

    '    ResetProperties(FunctionName.UpdateCus)

    '    objContent.username = mstrUserName
    '    objContent.pass = mstrPass
    '    objXElement =
    '        <Customers>
    '            <Customer>
    '                <Name><%= strCustFullName %></Name>
    '                <Code><%= intCustId %></Code>
    '                <TaxCode><%= strTaxCode %></TaxCode>
    '                <Address><%= strAddress %></Address>
    '                <BankAccountName><%= strBankAccountName %></BankAccountName>
    '                <BankName><%= strBankName %></BankName>
    '                <BankNumber><%= strBankName %></BankNumber>
    '                <Email><%= strEmail %></Email>
    '                <Fax></Fax>
    '                <Phone><%= strPhone %></Phone>
    '                <ContactPerson><%= strContact %></ContactPerson>
    '                <RepresentPerson><%= strCustShortName %></RepresentPerson>
    '                <CusType><%= intCustType %></CusType>
    '            </Customer>
    '        </Customers>

    '    objContent.XMLCusData = objXElement.ToString
    '    objRequest.Body = objContent

    '    Try
    '        objResponse = objSvc.E_InvoicePublish_PublishServiceSoap_UpdateCus(objRequest)
    '        Return True
    '    Catch ex As Exception
    '        Dim sw1 = New StreamWriter(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RQ.xml")
    '        Dim xs1 As New XmlSerializer(objRequest.GetType)
    '        xs1.Serialize(sw1, objRequest)
    '        sw1.Close()

    '        If objResponse IsNot Nothing Then
    '            Dim sw2 = New StreamWriter(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RP.xml")
    '            Dim xs2 As New XmlSerializer(objResponse.GetType)
    '            xs2.Serialize(sw2, objResponse)
    '            sw2.Close()
    '        End If
    '        Return False
    '    End Try

    'End Function
    Public Function SendAgainEmailServ(strPublishServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strPattern As String, strSerial As String _
                        , strInvoiceId As String, strEmails As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        Dim strSoapBody As String

        Dim objRequest As XElement
        Dim objData As XElement = <Invs>
                                      <Inv>
                                          <Fkey><%= strInvoiceId %></Fkey>
                                          <EmailDeliver><%= strEmails %></EmailDeliver>
                                      </Inv>
                                  </Invs>

        ResetProperties(FunctionName.SendAgainEmailServ)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <SendAgainEmailServ xmlns="http://tempuri.org/">
                                 <Account><%= strStaffAccount %></Account>
                                 <ACpass><%= strStaffPass %></ACpass>
                                 <username><%= strWsUser %></username>
                                 <password><%= strWsPass %></password>
                                 <xmlDataInvoiceEmail><%= objData.ToString %></xmlDataInvoiceEmail>
                                 <hdPattern><%= strPattern %></hdPattern>
                                 <Serial><%= strSerial %></Serial>
                             </SendAgainEmailServ>
                         </soap:Body>
                     </soap:Envelope>


        objHttp.open("POST", strPublishServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/SendAgainEmailServ")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "SendAgainEmailServResponse")

        ConvertResponseCode2Desc(FunctionName.SendAgainEmailServ, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function UpdateCus(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , intCustId As Integer, strCustFullName As String, strCustShortName As String _
                        , strTaxCode As String, strAddress As String, strEmail As String _
                        , intCustType As Integer, Optional strPhone As String = "" _
                        , Optional strContact As String = "" _
                        , Optional strBankAccountName As String = "", Optional strBankName As String = "" _
                        , Optional strAccountNumber As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objCustData As XElement

        ResetProperties(FunctionName.UpdateCus)

        objCustData =
            <Customers>
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
            </Customers>
        Dim strTemp As String = "<![CDATA[<Customers><Customer><Name> Kim Yen test </Name><Code>binh</Code><TaxCode></TaxCode><Address> Thôn Ngũ Hồ, xã Thiện Kế, Bình Xuyên, Vĩnh Phúc </Address><BankAccountName/><BankName/><BankNumber/><Email></Email><Fax/><Phone></Phone><ContactPerson/><RepresentPerson/><CusType>0</CusType></Customer></Customers>]]>"

        objRequest2 =
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <UpdateCus xmlns="http://tempuri.org/">
                        <XMLCusData><%= strTemp %></XMLCusData>
                        <username><%= strWsUser %></username>
                        <pass><%= strWsPass %></pass>
                        <convert>0</convert>
                    </UpdateCus>
                </soap:Body>
            </soap:Envelope>

        'objRequest2 =
        '    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
        '        <soap:Body>
        '            <UpdateCus xmlns="http://tempuri.org/">
        '                <XMLCusData><%= objCustData.ToString %></XMLCusData>
        '                <username><%= strWsUser %></username>
        '                <pass><%= strWsPass %></pass>
        '                <convert>0</convert>
        '            </UpdateCus>
        '        </soap:Body>
        '    </soap:Envelope>
        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/UpdateCus")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "UpdateCusResult")

        ConvertResponseCode2Desc(FunctionName.UpdateCus, mstrReponseCode)
        If IsNumeric(mstrReponseCode) AndAlso mstrReponseCode > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function HandleInvoiceErrors(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String, strResponse As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        End If

        ResetProperties(FunctionName.HandleInvoiceErrors)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <HandleInvoiceErrors xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <mtd><%= strResponse %></mtd>
                              </HandleInvoiceErrors>
                          </soap:Body>
                      </soap:Envelope>
        '<AttachFile>10</AttachFile>
        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/HandleInvoiceErrors")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "HandleInvoiceErrorsResponse")

        ConvertResponseCode2Desc(FunctionName.HandleInvoiceErrors, mstrReponseCode)
        If mstrReponseCode.StartsWith("ERR:") Then
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        Else
            Return True
        End If

    End Function
    Public Function listInvFromNoToNo(strPortalServiceUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strPattern As String, strSerial As String _
                        , intFromNbr As Integer, intToNumber As Integer) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.listInvFromNoToNo)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <listInvFromNoToNo xmlns="http://tempuri.org/">
                                 <invFromNo><%= intFromNbr.ToString.PadLeft(7, "0") %></invFromNo>
                                 <invToNo><%= intToNumber.ToString.PadLeft(7, "0") %></invToNo>
                                 <invPattern><%= strPattern %></invPattern>
                                 <invSerial><%= strSerial %></invSerial>
                                 <userName><%= strWsUser %></userName>
                                 <userPass><%= strWsPass %></userPass>
                             </listInvFromNoToNo>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strPortalServiceUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/listInvFromNoToNo")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody
        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strPortalServiceUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "listInvFromNoToNoResponse")

        ConvertResponseCode2Desc(FunctionName.listInvFromNoToNo, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function MassUpdateCus2(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , lstCustData As List(Of XElement)) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objCustData As XElement
        Dim strTmp As String
        ResetProperties(FunctionName.UpdateCus)

        objCustData =
            <Customers>
            </Customers>
        For Each objXelement As XElement In lstCustData
            objCustData.Add(objXelement)
        Next

        strTmp = objCustData.ToString
        'strTmp = "<Customers><Customer><Name><![CDATA[Kim Yen test]]></Name><Code>binh</Code><TaxCode></TaxCode><Address><![CDATA[Thôn Ngũ Hồ, xã Thiện Kế, Bình Xuyên, Vĩnh Phúc]]></Address><BankAccountName/><BankName/><BankNumber/><Email></Email><Fax/><Phone></Phone><ContactPerson/><RepresentPerson/><CusType>0</CusType></Customer></Customers>"

        objRequest2 =
            <soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
                <soap12:Body>
                    <UpdateCus xmlns="http://tempuri.org/">
                        <XMLCusData><%= strTmp %></XMLCusData>
                        <username><%= strWsUser %></username>
                        <pass><%= strWsPass %></pass>
                        <convert>0</convert>
                    </UpdateCus>
                </soap12:Body>
            </soap12:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/UpdateCus")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "UpdateCusResult")

        ConvertResponseCode2Desc(FunctionName.UpdateCus, mstrReponseCode)
        If IsNumeric(mstrReponseCode) AndAlso mstrReponseCode > 0 Then
            Return True
        Else
            Append2TextFile("Error:" & FunctionName.UpdateCus & vbNewLine & "URL:" & strWsUrl _
                            & vbNewLine & "Request" & vbNewLine & ReformatXml(mstrLastRequest) _
                            & vbNewLine & "Response" & vbNewLine & ReformatXml(mstrLastResponse))
            Return False
        End If
    End Function
    Public Function PublishInvFkey(strWsUrl As String, strStaffAccount As String, strStaffPass As String _
                                , strFkeys As String, strWsUser As String, strWsPass As String _
                               , Optional strPattern As String = "", Optional strSerial As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.PublishInvFkey)
        objRequest =
            <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                <soap:Body>
                    <PublishInvFkey xmlns="http://tempuri.org/">
                        <Account><%= strStaffAccount %></Account>
                        <ACpass><%= strStaffPass %></ACpass>
                        <lsFkey><%= strFkeys %></lsFkey>
                        <username><%= strWsUser %></username>
                        <password><%= strWsPass %></password>
                        <pattern><%= strPattern %></pattern>
                        <serial><%= strSerial %></serial>
                    </PublishInvFkey>
                </soap:Body>
            </soap:Envelope>


        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/PublishInvFkey")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText


        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "PublishInvFkeyResponse")

        ConvertResponseCode2Desc(FunctionName.PublishInvFkey, mstrReponseCode)
        If mstrResponseDesc.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

        Return True
    End Function
    Public Function ReceivedInvoiceErrors(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String, strResponse As String _
                        , objInput As XElement) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        End If

        ResetProperties(FunctionName.ReceivedInvoiceErrors)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <ReceivedInvoiceErrors xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xml><%= objInput.ToString %></xml>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <mtd><%= strResponse %></mtd>
                              </ReceivedInvoiceErrors>
                          </soap:Body>
                      </soap:Envelope>
        '<AttachFile>10</AttachFile>
        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ReceivedInvoiceErrors")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ReceivedInvoiceErrorsResponse")

        ConvertResponseCode2Desc(FunctionName.ReceivedInvoiceErrors, mstrReponseCode)
        If mstrReponseCode.StartsWith("ERR:") Then
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        Else
            Return True
        End If

    End Function
    Public Function SendInvNoticeErrors(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String, strPattern As String _
                        , objInvoiceList As XElement) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        End If

        ResetProperties(FunctionName.SendInvNoticeErrors)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <SendInvNoticeErrors xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xml><%= objInvoiceList.ToString %></xml>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <convert>0</convert>
                              </SendInvNoticeErrors>
                          </soap:Body>
                      </soap:Envelope>
        '<pattern><%= strPattern %></pattern>
        '<AttachFile>10</AttachFile>
        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/SendInvNoticeErrors")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "SendInvNoticeErrorsResponse")

        ConvertResponseCode2Desc(FunctionName.SendInvNoticeErrors, mstrReponseCode)
        If mstrReponseCode.StartsWith("ERR:") Then
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        Else
            Return True
        End If

    End Function
    'Public Function ImportAndPublishInv(strStaffAccount As String, strStaffPass As String _
    '                    , strCustId As String, strCustFullName As String _
    '                    , strAddress As String, strPhone As String _
    '                    , strTaxCode As String, strFOP As String _
    '                    , strInvoiceKey As String, strKindOfService As String _
    '                    , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
    '                    , Optional strPattern As String = "", Optional strSerial As String = "" _
    '                    , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
    '                    , Optional strBuyer As String = "") As Boolean

    '    Dim objRequest As New ImportAndPublishInvRequest
    '    Dim objContent As New ImportAndPublishInvRequestBody

    '    Dim objXElement As XElement
    '    Dim objInv As XElement
    '    Dim objInvoice As XElement
    '    Dim objXProd As XElement

    '    Dim decVat0 As Decimal
    '    Dim decVat5 As Decimal
    '    Dim decVat10 As Decimal
    '    Dim decVatable0 As Decimal
    '    Dim decVatable5 As Decimal
    '    Dim decVatable10 As Decimal

    '    Dim decTotalDiscount As Decimal
    '    Dim decTotalVat As Decimal
    '    Dim decTotalVatable As Decimal
    '    Dim decTotalInv As Decimal
    '    Dim intRefundMultiplier As Integer = 1

    '    If intKindOfInvoice = 3 Then
    '        intRefundMultiplier = -1
    '    End If

    '    Dim objResponse As ImportAndPublishInvResponse
    '    Dim objSvc As New PublishServiceSoapClient


    '    ResetProperties(FunctionName.ImportAndPublishInv)

    '    objContent.username = mstrUserName
    '    objContent.password = mstrPass
    '    objContent.Account = strStaffAccount
    '    objContent.ACpass = strStaffPass

    '    If strPattern = "" Then
    '        strPattern = "01GTKT0/001"
    '    End If
    '    If strSerial = "" Then
    '        strSerial = "AA/" & Format(Now, "yy") & "E"
    '    End If

    '    objContent.pattern = strPattern
    '    objContent.serial = strSerial
    '    objContent.convert = 0

    '    objXElement = CreateInvData(strCustId, strCustFullName _
    '                    , strAddress, strPhone, strTaxCode, strFOP, strInvoiceKey, strKindOfService _
    '                    , intKindOfInvoice, lstProducts _
    '                    , strPattern, strSerial, decTax, decCharge, strBuyer)

    '    objContent.xmlInvData = objXElement.ToString
    '    objRequest.Body = objContent

    '    Try
    '        objResponse = objSvc.E_InvoicePublish_PublishServiceSoap_ImportAndPublishInv(objRequest)
    '        mstrReponseCode = objResponse.Body.ImportAndPublishInvResult

    '        'If objResponse.Body.ImportAndPublishInvResult.ToString.StartsWith("ERR") Then
    '        Dim sw1 = New StreamWriter(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RQ.xml")
    '        Dim xs1 As New XmlSerializer(objRequest.GetType)
    '        xs1.Serialize(sw1, objRequest)
    '        sw1.Close()

    '        Dim xmlDoc As New XmlDocument
    '        xmlDoc.LoadXml(objRequest.Body.xmlInvData)
    '        xmlDoc.Save(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RQ_content.xml")

    '        If objResponse IsNot Nothing Then
    '            Dim sw2 = New StreamWriter(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RP.xml")
    '            Dim xs2 As New XmlSerializer(objResponse.GetType)
    '            xs2.Serialize(sw2, objResponse)
    '            sw2.Close()
    '            'xmlDoc.LoadXml(objResponse.Body.ImportAndPublishInvResult)
    '            'xmlDoc.Save(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RP.xml")
    '        End If
    '        'End If

    '        Return True
    '    Catch ex As Exception

    '        Dim sw1 = New StreamWriter(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RQ.xml")
    '        Dim xs1 As New XmlSerializer(objRequest.GetType)
    '        xs1.Serialize(sw1, objRequest)
    '        sw1.Close()

    '        If objResponse IsNot Nothing Then
    '            Dim sw2 = New StreamWriter(My.Application.Info.DirectoryPath & "\" & Format(Now, "HHmmss_") & mstrFuntionName & "RQ.xml")
    '            Dim xs2 As New XmlSerializer(objResponse.GetType)
    '            xs2.Serialize(sw2, objResponse)
    '            sw2.Close()
    '        End If
    '        Return False
    '    End Try
    'End Function
    Public Function GetCertInfo(strWsUrl As String _
                        , strWsUser As String, strWsPass As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.GetCertInfo)
        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <GetCertInfo xmlns="http://tempuri.org/">
                                 <userName><%= strWsUser %></userName>
                                 <password><%= strWsPass %></password>
                             </GetCertInfo>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/GetCertInfo")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & mstrLastRequest)
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & mstrLastResponse)
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "GetCertInfoResult")

        ConvertResponseCode2Desc(FunctionName.getHashInv, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            mobjXmlDoc.LoadXml(mstrReponseCode)
            mobjCertInfo.OwnCA = GetWsResponse(mobjXmlDoc, "OwnCA")
            mobjCertInfo.OrganizationCA = GetWsResponse(mobjXmlDoc, "OrganizationCA")
            mobjCertInfo.SerialNumber = GetWsResponse(mobjXmlDoc, "SerialNumber")
            mobjCertInfo.ValidFrom = GetWsResponse(mobjXmlDoc, "ValidFrom")
            mobjCertInfo.ValidTo = GetWsResponse(mobjXmlDoc, "ValidTo")
            Return True
        Else
            Return False
        End If

    End Function
    Public Function getHashInv(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strSerialCert As String, strFkey As String, strPattern As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        Dim strSoapBody As String

        Dim objRequest As XElement
        Dim objFkey As XElement


        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.getHashInv)
        objFkey = <Invoices>
                      <Inv>
                          <key><%= strFkey %></key>
                      </Inv>
                  </Invoices>

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <getHashInv xmlns="http://tempuri.org/">
                                 <Account><%= strStaffAccount %></Account>
                                 <ACpass><%= strStaffPass %></ACpass>
                                 <username><%= strWsUser %></username>
                                 <password><%= strWsPass %></password>
                                 <serialCert><%= strSerialCert %></serialCert>
                                 <xmlFkeyInv><%= objFkey.ToString %></xmlFkeyInv>
                                 <pattern><%= strPattern %></pattern>
                             </getHashInv>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/getHashInv")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & mstrLastRequest)
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & mstrLastResponse)
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "getHashInvResult")

        ConvertResponseCode2Desc(FunctionName.getHashInv, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            mobjXmlDoc.LoadXml(mstrReponseCode)
            mobjInvToken.KeyInv = GetWsResponse(mobjXmlDoc, "key")
            mobjInvToken.IdInv = GetWsResponse(mobjXmlDoc, "idInv")
            mobjInvToken.HashValue = GetWsResponse(mobjXmlDoc, "hashValue")
            Return True
        Else
            Return False
        End If

    End Function
    Public Function getHashInvWithToken(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strSerialCert As String, strInvData As String, intInvType As Integer _
                        , strToken As String, strPattern As String, strSerial As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement


        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.getHashInvWithToken)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <getHashInvWithToken xmlns="http://tempuri.org/">
                                 <Account><%= strStaffAccount %></Account>
                                 <ACpass><%= strStaffPass %></ACpass>
                                 <xmlInvData><%= strInvData %></xmlInvData>
                                 <username><%= strWsUser %></username>
                                 <password><%= strWsPass %></password>
                                 <serialCert><%= strSerialCert %></serialCert>
                                 <type><%= intInvType %></type>
                                 <invToken><%= strToken %></invToken>
                                 <pattern><%= strPattern %></pattern>
                                 <serial><%= strSerial %></serial>
                                 <convert>0</convert>
                             </getHashInvWithToken>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/getHashInvWithToken")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & mstrLastRequest)
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & mstrLastResponse)
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "getHashInvWithTokenResult")

        ConvertResponseCode2Desc(FunctionName.getHashInvWithToken, mstrReponseCode)
        If Not mstrReponseCode.StartsWith("ERR:") Then
            mobjXmlDoc.LoadXml(mstrReponseCode)
            mobjInvToken.KeyInv = GetWsResponse(mobjXmlDoc, "key")
            mobjInvToken.IdInv = GetWsResponse(mobjXmlDoc, "idInv")
            mobjInvToken.HashValue = GetWsResponse(mobjXmlDoc, "hashValue")
            Return True
        Else
            Return False
        End If

    End Function
    Public Function publishInvWithToken(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , objE_InvConnect As clsE_InvConnect, strSerialCert As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = ""
                        ) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If



        strEmail = Replace(strEmail, ";", ",")

        objXElement = CreateInvData(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strInvoiceKey, strKindOfService _
                        , intKindOfInvoice, lstProducts _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail)

        If Not getHashInvWithToken(objE_InvConnect.WsUrl, objE_InvConnect.UserName, objE_InvConnect.UserPass _
                                            , objE_InvConnect.AccountName, objE_InvConnect.AccountPass _
                                            , strSerialCert, objXElement.ToString, intKindOfInvoice _
                                            , "", strPattern, strSerial) Then
            Return False
        End If

        ResetProperties(FunctionName.publishInvWithToken)

        Dim strSign As String = SignHashVNPT(mobjInvToken.HashValue, strSerialCert)

        objXElement =
            <Invoices>
                <SerialCert><%= strSerialCert %></SerialCert>
                <Inv>
                    <key><%= mobjInvToken.KeyInv %></key>
                    <idInv><%= mobjInvToken.IdInv %></idInv>
                    <signValue><%= strSign %></signValue>
                </Inv>
            </Invoices>


        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                          <soap:Body>
                              <publishInvWithToken xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                              </publishInvWithToken>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/publishInvWithToken")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & mstrLastRequest)
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & mstrLastResponse)
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "publishInvWithTokenResult")

        ConvertResponseCode2Desc(FunctionName.publishInvWithToken, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function AdjustActionAssignedNo(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strNewFKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strOldFkey As String, strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strOldPattern As String = "", Optional strOldSerial As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        End If

        ResetProperties(FunctionName.AdjustActionAssignedNo)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvDataAdjust(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strNewFKey, strKindOfService _
                        , intKindOfInvoice, lstProducts, strAdjustType _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <AdjustActionAssignedNo xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <pass><%= strWsPass %></pass>
                                  <fkey><%= strOldFkey %></fkey>
                                  <AttachFile>10</AttachFile>
                                  <convert>0</convert>
                                  <pattern><%= strOldPattern %></pattern>
                                  <serial><%= strOldSerial %></serial>
                              </AdjustActionAssignedNo>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/AdjustActionAssignedNo")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "AdjustActionAssignedNoResponse")

        ConvertResponseCode2Desc(FunctionName.AdjustActionAssignedNo, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function AdjustInvoiceAction(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strNewFKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strOldFkey As String, strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" Then
            If Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pblnLogXml = True
            End If
        End If

        ResetProperties(FunctionName.AdjustInvoiceAction)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvDataAdjust(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strNewFKey, strKindOfService _
                        , intKindOfInvoice, lstProducts, strAdjustType _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <AdjustInvoiceAction xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <pass><%= strWsPass %></pass>
                                  <fkey><%= strOldFkey %></fkey>
                                  <AttachFile>10</AttachFile>
                                  <convert>0</convert>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                              </AdjustInvoiceAction>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/AdjustInvoiceAction")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "AdjustInvoiceActionResponse")

        ConvertResponseCode2Desc(FunctionName.AdjustInvoiceAction, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        End If

    End Function
    Public Function AdjustInvoiceNoPublish(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strOldFKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strNewFkey As String, strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strOldPattern As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" Then
            If Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pblnLogXml = True
            End If
        End If

        ResetProperties(FunctionName.AdjustInvoiceNoPublish)

        strEmail = Replace(strEmail, ";", ",")
        'If strOldPattern = "" Then
        objXElement = CreateInvDataAdjust(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strNewFkey, strKindOfService _
                        , intKindOfInvoice, lstProducts, strAdjustType _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)
        'Else
        '    objXElement = CreateInvDataAdjustNoInv(strCustId, strCustFullName _
        '                , strAddress, strPhone, strTaxCode, strFOP, strNewFkey, strKindOfService _
        '                , intKindOfInvoice, lstProducts, strAdjustType _
        '                , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)
        'End If


        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <AdjustInvoiceNoPublish xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <pass><%= strWsPass %></pass>
                                  <convert>0</convert>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                                  <fkey><%= strOldFKey %></fkey>
                                  <OldPattern><%= strOldPattern %></OldPattern>
                              </AdjustInvoiceNoPublish>
                          </soap:Body>
                      </soap:Envelope>
        '<fkey><%= strOldFKey %></fkey>
        '<oldPattern><%= strOldPattern %></oldPattern>
        '<oldSerial><%= strOldSerial %></oldSerial>
        '<oldNo><%= intOldInvNo %></oldNo>
        '<strOldArisingDate><%= Format(dteOldIssueDate, "dd/MM/yyyy") %></strOldArisingDate>
        '<AttachFile>10</AttachFile>
        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/AdjustInvoiceNoPublish")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "AdjustInvoiceNoPublishResponse")

        ConvertResponseCode2Desc(FunctionName.AdjustInvoiceNoPublish, mstrReponseCode)
        If mstrReponseCode.StartsWith("ERR:") Then
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        Else
            Return True
        End If

    End Function
    Public Function AdjustInvoiceNote(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strNote As String, strOldFKey As String, strPattern As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement = <DieuChinhHD>
                                          <Description><%= strNote %></Description>
                                      </DieuChinhHD>

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        End If

        ResetProperties(FunctionName.AdjustInvoiceNote)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <AdjustInvoiceNote xmlns="http://tempuri.org/">
                                  <account><%= strStaffAccount %></account>
                                  <accPass><%= strStaffPass %></accPass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <userName><%= strWsUser %></userName>
                                  <userPass><%= strWsPass %></userPass>
                                  <fkey><%= strOldFKey %></fkey>
                                  <AttachFile>11</AttachFile>
                                  <pattern><%= strPattern %></pattern>
                              </AdjustInvoiceNote>
                          </soap:Body>
                      </soap:Envelope>
        '<AttachFile>10</AttachFile>
        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/AdjustInvoiceNote")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "AdjustInvoiceNoteResponse")

        ConvertResponseCode2Desc(FunctionName.AdjustInvoiceNote, mstrReponseCode)
        If mstrReponseCode.StartsWith("ERR:") Then
            Return False
        Else
            Return True
        End If

    End Function
    Public Function AdjustReplaceInvWithToken(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String, strOldInvNo As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional intVatDiscount As Integer = 0) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.ImportAndPublishInv)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvData(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strInvoiceKey, strKindOfService _
                        , intKindOfInvoice, lstProducts _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, intVatDiscount)
        objXElement = <Invoices>
                          <SerialCert><%= mobjCertInfo.SerialNumber %></SerialCert>
                          <PatternOld><%= strPattern %></PatternOld>
                          <SerialOld><%= strSerial %></SerialOld>
                          <NoOlde><%= strOldInvNo %></NoOlde>
                          <Inv>
                              <key> fkey hóa đơn mới </key>
                              <idInv> id hóa đơn mới trên hệ thống vnpt </idInv>
                              <signValue> chuỗi ký </signValue>
                          </Inv>
                      </Invoices>

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <ImportAndPublishInv xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                                  <convert>0</convert>
                              </ImportAndPublishInv>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ImportAndPublishInv")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ImportAndPublishInvResponse")

        ConvertResponseCode2Desc(FunctionName.ImportAndPublishInv, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function

    Public Function AdjustWithoutInv(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strNewFKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strOldFkey As String, strAdjustType As String _
                        , strOldPattern As String, strOldSerial As String, intOldInvNo As Integer _
                        , dteOldIssueDate As Date _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml Then
            If MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pblnLogXml = True
            End If
        End If

        ResetProperties(FunctionName.AdjustWithoutInv)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvDataAdjust(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strNewFKey, strKindOfService _
                        , intKindOfInvoice, lstProducts, strAdjustType _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)
        'objXElement = CreateInvDataAdjustNoInv(strCustId, strCustFullName _
        '                , strAddress, strPhone, strTaxCode, strFOP, strNewFKey, strKindOfService _
        '                , intKindOfInvoice, lstProducts, strAdjustType _
        '                , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)
        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <AdjustWithoutInv xmlns="http://tempuri.org/">
                                  <account><%= strStaffAccount %></account>
                                  <accPass><%= strStaffPass %></accPass>
                                  <invXml><%= objXElement.ToString %></invXml>
                                  <userName><%= strWsUser %></userName>
                                  <userPass><%= strWsPass %></userPass>
                                  <oldPattern><%= strOldPattern %></oldPattern>
                                  <oldSerial><%= strOldSerial %></oldSerial>
                                  <oldNo><%= intOldInvNo %></oldNo>
                                  <strOldArisingDate><%= Format(dteOldIssueDate, "dd/MM/yyyy") %></strOldArisingDate>
                                  <convert>0</convert>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                                  <relatedInvType>3</relatedInvType>
                                  <feature></feature>
                              </AdjustWithoutInv>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/AdjustWithoutInv")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "AdjustWithoutInvResponse")

        ConvertResponseCode2Desc(FunctionName.AdjustWithoutInv, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function CancelInvoiceWithToken(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strPattern As String, strSerial As String, intInvNo As Integer
                        ) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.CancelInvoiceWithToken)

        objXElement =
            <Invoices>
                <Inv>
                    <Serial><%= strSerial %></Serial>
                    <InvNo><%= intInvNo %></InvNo>
                </Inv>
            </Invoices>

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                          <soap:Body>
                              <CancelInvoiceWithToken xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlData><%= objXElement.ToString %></xmlData>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <pattern><%= strPattern %></pattern>
                              </CancelInvoiceWithToken>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/CancelInvoiceWithToken")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & mstrLastRequest)
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & mstrLastResponse)
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "CancelInvoiceWithTokenResult")

        ConvertResponseCode2Desc(FunctionName.CancelInvoiceWithToken, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function ImportAndPublishInv(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional intVatDiscount As Integer = 0) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.ImportAndPublishInv)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvData(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strInvoiceKey, strKindOfService _
                        , intKindOfInvoice, lstProducts _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, intVatDiscount)


        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <ImportAndPublishInv xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                                  <convert>0</convert>
                              </ImportAndPublishInv>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ImportAndPublishInv")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ImportAndPublishInvResponse")

        ConvertResponseCode2Desc(FunctionName.ImportAndPublishInv, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function ImportInv(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.ImportInv)

        strEmail = Replace(strEmail, ";", ",")

        objXElement = CreateInvData(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strInvoiceKey, strKindOfService _
                        , intKindOfInvoice, lstProducts _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                          <soap:Body>
                              <ImportInv xmlns="http://tempuri.org/">
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <convert>0</convert>
                              </ImportInv>
                          </soap:Body>
                      </soap:Envelope>

        'objRequest2 = <soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
        '                  <soap12:Body>
        '                      <ImportInv xmlns="http://tempuri.org/">
        '                          <Account><%= strStaffAccount %></Account>
        '                          <ACpass><%= strStaffPass %></ACpass>
        '                          <xmlInvData><%= objXElement.ToString %></xmlInvData>
        '                          <username><%= strWsUser %></username>
        '                          <password><%= strWsPass %></password>
        '                          <pattern><%= strPattern %></pattern>
        '                          <serial><%= strSerial %></serial>
        '                          <convert>0</convert>
        '                      </ImportInv>
        '                  </soap12:Body>
        '              </soap12:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ImportInv")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & mstrLastRequest)
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & mstrLastResponse)
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ImportInvResponse")

        ConvertResponseCode2Desc(FunctionName.ImportInv, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function ImportInvByPattern(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strInvoiceKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strSerialCert As String = "", Optional intVatDiscount As Integer = 0) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.ImportInvByPattern)

        strEmail = Replace(strEmail, ";", ",")

        objXElement = CreateInvData(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strInvoiceKey, strKindOfService _
                        , intKindOfInvoice, lstProducts _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail, strSerialCert)


        objRequest2 = <soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
                          <soap12:Body>
                              <ImportInvByPattern xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <password><%= strWsPass %></password>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                                  <convert>0</convert>
                              </ImportInvByPattern>
                          </soap12:Body>
                      </soap12:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ImportInvByPattern")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ImportInvByPatternResponse")

        ConvertResponseCode2Desc(FunctionName.ImportInvByPattern, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function deleteInvoiceByFkey(strWsUrlPublishService As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strInvId As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.deleteInvoiceByFkey)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <deleteInvoiceByFkey xmlns="http://tempuri.org/">
                                 <lstFkey><%= strInvId %></lstFkey>
                                 <username><%= strWsUser %></username>
                                 <password><%= strWsPass %></password>
                                 <Account><%= strStaffAccount %></Account>
                                 <ACpass><%= strStaffPass %></ACpass>
                             </deleteInvoiceByFkey>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strWsUrlPublishService, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/deleteInvoiceByFkey")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "deleteInvoiceByFkeyResult")

        ConvertResponseCode2Desc(FunctionName.deleteInvoiceByFkey, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else
            Append2TextFile("Error:" & FunctionName.UpdateCus & vbNewLine & "URL:" & strWsUrlPublishService _
                            & vbNewLine & "Request" & vbNewLine & ReformatXml(mstrLastRequest) _
                            & vbNewLine & "Response" & vbNewLine & ReformatXml(mstrLastResponse))
            Return False
        End If

    End Function
    Public Function deleteInvoiceByID(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strInvId As String) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest As XElement

        ResetProperties(FunctionName.deleteInvoiceByID)

        objRequest = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                         <soap:Body>
                             <deleteInvoiceByID xmlns="http://tempuri.org/">
                                 <lstID><%= strInvId %></lstID>
                                 <username><%= strWsUser %></username>
                                 <password><%= strWsPass %></password>
                                 <Account><%= strStaffAccount %></Account>
                                 <ACpass><%= strStaffPass %></ACpass>
                             </deleteInvoiceByID>
                         </soap:Body>
                     </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/deleteInvoiceByID")

        'Send the SOAP request
        strSoapBody = objRequest.ToString
        mstrLastRequest = strSoapBody

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "deleteInvoiceByID")

        ConvertResponseCode2Desc(FunctionName.deleteInvoiceByID, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else

            Return False
        End If

    End Function
    Public Function ReplaceInvoiceAction(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strNewFKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strOldFkey As String, strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml AndAlso MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            pblnLogXml = True
        Else
            pblnLogXml = False
        End If

        ResetProperties(FunctionName.ReplaceInvoiceAction)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvDataReplace(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strNewFKey, strKindOfService _
                        , intKindOfInvoice, lstProducts, strAdjustType _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <ReplaceInvoiceAction xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <pass><%= strWsPass %></pass>
                                  <fkey><%= strOldFkey %></fkey>
                                  <Attachfile>10</Attachfile>
                                  <convert>0</convert>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                              </ReplaceInvoiceAction>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ReplaceInvoiceAction")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ReplaceInvoiceActionResponse")

        ConvertResponseCode2Desc(FunctionName.ReplaceInvoiceAction, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Function ReplaceInvoiceNoPublish(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strNewFKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strOldFkey As String, strAdjustType As String _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "" _
                        , Optional strOldPattern As String = "", Optional strOldSerial As String = "" _
                        , Optional intOldInvNo As Integer = 0, Optional dteOldIssueDate As Date = Nothing) As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement


        ResetProperties(FunctionName.ReplaceInvoiceNoPublish)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvDataReplace(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strNewFKey, strKindOfService _
                        , intKindOfInvoice, lstProducts, strAdjustType _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <ReplaceInvoiceNoPublish xmlns="http://tempuri.org/">
                                  <Account><%= strStaffAccount %></Account>
                                  <ACpass><%= strStaffPass %></ACpass>
                                  <xmlInvData><%= objXElement.ToString %></xmlInvData>
                                  <username><%= strWsUser %></username>
                                  <pass><%= strWsPass %></pass>
                                  <fkey><%= strOldFkey %></fkey>
                                  <convert>0</convert>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                                  <OldPattern></OldPattern>
                              </ReplaceInvoiceNoPublish>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ReplaceInvoiceNoPublish")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ReplaceInvoiceNoPublishResponse")

        ConvertResponseCode2Desc(FunctionName.ReplaceInvoiceNoPublish, mstrReponseCode)
        If mstrReponseCode.StartsWith("ERR:") Then
            If Not pblnLogXml Then
                Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
                Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
            End If
            Return False
        Else
            Return True
        End If

    End Function
    Public Function ReplaceWithoutInv(strWsUrl As String _
                        , strWsUser As String, strWsPass As String _
                        , strStaffAccount As String, strStaffPass As String _
                        , strCustId As String, strCustFullName As String _
                        , strAddress As String, strPhone As String _
                        , strTaxCode As String, strFOP As String _
                        , strNewFKey As String, strKindOfService As String _
                        , intKindOfInvoice As Integer, lstProducts As List(Of clsProduct) _
                        , strOldFkey As String, strAdjustType As String _
                        , strOldPattern As String, strOldSerial As String, intOldInvNo As Integer _
                        , dteOldIssueDate As Date _
                        , Optional strPattern As String = "", Optional strSerial As String = "" _
                        , Optional decTax As Decimal = 0, Optional decCharge As Decimal = 0 _
                        , Optional strBuyer As String = "", Optional strEmail As String = "") As Boolean

        Dim objHttp As New MSXML2.XMLHTTP60
        'Dim objXmlDoc As New Xml.XmlDocument
        Dim strSoapBody As String

        Dim objRequest2 As XElement
        Dim objXElement As XElement

        If My.Computer.Name = "5-247" AndAlso Not pblnLogXml Then
            If MsgBox("LogXML?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                pblnLogXml = True
            End If
        End If

        ResetProperties(FunctionName.ReplaceWithoutInv)

        strEmail = Replace(strEmail, ";", ",")
        objXElement = CreateInvDataAdjustNoInv(strCustId, strCustFullName _
                        , strAddress, strPhone, strTaxCode, strFOP, strNewFKey, strKindOfService _
                        , intKindOfInvoice, lstProducts, strAdjustType _
                        , strPattern, strSerial, decTax, decCharge, strBuyer, strEmail,, 0)

        objRequest2 = <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://www.w3.org/2003/05/soap-envelope">
                          <soap:Body>
                              <ReplaceWithoutInv xmlns="http://tempuri.org/">
                                  <account><%= strStaffAccount %></account>
                                  <accPass><%= strStaffPass %></accPass>
                                  <invXml><%= objXElement.ToString %></invXml>
                                  <userName><%= strWsUser %></userName>
                                  <userPass><%= strWsPass %></userPass>
                                  <oldPattern><%= strOldPattern %></oldPattern>
                                  <oldSerial><%= strOldSerial %></oldSerial>
                                  <oldNo><%= intOldInvNo %></oldNo>
                                  <strOldArisingDate><%= Format(dteOldIssueDate, "dd/MM/yyyy") %></strOldArisingDate>
                                  <convert>0</convert>
                                  <pattern><%= strPattern %></pattern>
                                  <serial><%= strSerial %></serial>
                                  <relatedInvType>3</relatedInvType>
                                  <feature></feature>
                              </ReplaceWithoutInv>
                          </soap:Body>
                      </soap:Envelope>

        objHttp.open("POST", strWsUrl, False)
        'False = do not respond immediately

        objHttp.setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        objHttp.setRequestHeader("Content-Length", "length")
        objHttp.setRequestHeader("SOAPAction", "http://tempuri.org/ReplaceWithoutInv")

        'Send the SOAP request
        strSoapBody = objRequest2.ToString
        mstrLastRequest = strSoapBody

        If pblnLogXml Then
            Append2TextFile("E_InvLog:" & strWsUrl & vbNewLine & "Request:" & vbNewLine & ReformatXml(mstrLastRequest))
        End If

        objHttp.send(strSoapBody)
        mstrLastResponse = objHttp.responseText
        If pblnLogXml Then
            Append2TextFile("Response:" & vbNewLine & ReformatXml(mstrLastResponse))
        End If

        mobjXmlDoc.LoadXml(mstrLastResponse)

        mstrReponseCode = GetWsResponse(mobjXmlDoc, "ReplaceWithoutInvResponse")

        ConvertResponseCode2Desc(FunctionName.ReplaceWithoutInv, mstrReponseCode)
        If mstrReponseCode.StartsWith("OK") Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Property ReponseCode As String
        Get
            Return mstrReponseCode
        End Get
        Set(value As String)
            mstrResponseDesc = ConvertResponseCode2Desc(mstrFuntionName, value)
            mstrReponseCode = value
        End Set
    End Property
    Public Property FuntionName As String
        Get
            Return mstrFuntionName
        End Get
        Set(value As String)
            mstrFuntionName = value
        End Set
    End Property
    Public Property ResponseDesc As String
        Get
            Return mstrResponseDesc
        End Get
        Set(value As String)
            mstrResponseDesc = value
        End Set
    End Property
    Public Property LastResponse As String
        Get
            Return mstrLastResponse
        End Get
        Set(value As String)
            mstrLastResponse = value
        End Set
    End Property
    Public Property LastRequest As String
        Get
            Return mstrLastRequest
        End Get
        Set(value As String)
            mstrLastRequest = value
        End Set
    End Property
    Public Property InvToken As clsInvToken
        Get
            Return mobjInvToken
        End Get
        Set(value As clsInvToken)
            mobjInvToken = value
        End Set
    End Property
    Public Property CertInfo As clsCertInfo
        Get
            Return mobjCertInfo
        End Get
        Set(value As clsCertInfo)
            mobjCertInfo = value
        End Set
    End Property
End Class
