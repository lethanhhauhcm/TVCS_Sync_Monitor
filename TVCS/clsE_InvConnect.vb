Public Class clsE_InvConnect
    Private mstrWsUrl As String
    Private mstrUserName As String
    Private mstrUserPass As String
    Private mstrAccountName As String
    Private mstrAccountPass As String
    Private mstrTvc As String
    Private mstrPortalServiceUrl As String
    Private mstrBusinessServiceUrl As String
    Public Sub New(blnTt78 As Boolean, strTvc As String)
        Select Case strTvc
            Case "TVPM SGN"
                mstrWsUrl = "https://tranvietphat-hcm-tt78admin.vnpt-invoice.com.vn/PublishService.asmx"
                mstrBusinessServiceUrl = "https://tranvietphat-hcm-tt78admin.vnpt-invoice.com.vn/BusinessService.asmx"
                mstrPortalServiceUrl = "https://tranvietphat-hcm-tt78admin.vnpt-invoice.com.vn/PortalService.asmx"
                mstrAccountName = "tranvietphat-hcmadmin"
                mstrAccountPass = "Einv@oi@vn#pt20"
                mstrUserPass = "Einv@oi@vn#pt20"
                mstrUserName = "tranvietphathcmservice"
            Case "TVTR"
                If pblnTestInv Then
                    mstrWsUrl = "https://tranviethcm-tt78admindemo.vnpt-invoice.com.vn/PublishService.asmx"
                    mstrBusinessServiceUrl = "https://tranviethcm-tt78admindemo.vnpt-invoice.com.vn/BusinessService.asmx"
                    mstrPortalServiceUrl = "https://tranviethcm-tt78admindemo.vnpt-invoice.com.vn/PortalService.asmx"
                    mstrAccountName = "tranviethcmadmin"
                    mstrAccountPass = "Einv@oi@vn#pt20"
                    mstrUserPass = "Einv@oi@vn#pt20"
                    mstrUserName = "tranviethcmservice"
                Else
                    mstrWsUrl = "https://tranviethcm-tt78admin.vnpt-invoice.com.vn/PublishService.asmx"
                    mstrBusinessServiceUrl = "https://tranviethcm-tt78admin.vnpt-invoice.com.vn/BusinessService.asmx"
                    mstrPortalServiceUrl = "https://tranviethcm-tt78admin.vnpt-invoice.com.vn/PortalService.asmx"
                    mstrAccountName = "tranviethcmadmin"
                    mstrAccountPass = "Einv@oi@vn#pt20"
                    mstrUserName = "tranviethcmservice"
                    mstrUserPass = "Einv@oi@vn#pt20"
                End If

            Case "TVTR HAN"
                mstrWsUrl = "https://0301069809-001-tt78cadmin.vnpt-invoice.com.vn/PublishService.asmx"
                mstrBusinessServiceUrl = "https://0301069809-001-tt78cadmin.vnpt-invoice.com.vn/BusinessService.asmx"
                mstrPortalServiceUrl = "https://0301069809-001-tt78cadmin.vnpt-invoice.com.vn/PortalService.asmx"
                mstrAccountName = "0301069809-001_admin"
                mstrAccountPass = "Einv@oi@vn#pt20"
                mstrUserName = "0301069809001service"
                mstrUserPass = "Einv@oi@vn#pt20"
            Case Else
                MsgBox("Chưa có thông tin kết nối VNPT!")
        End Select

        mstrTvc = strTvc
    End Sub

    Public Property WsUrl As String
        Get
            Return mstrWsUrl
        End Get
        Set(value As String)
            mstrWsUrl = value
        End Set
    End Property
    Public Property UserName As String
        Get
            Return mstrUserName
        End Get
        Set(value As String)
            mstrUserName = value
        End Set
    End Property
    Public Property UserPass As String
        Get
            Return mstrUserPass
        End Get
        Set(value As String)
            mstrUserPass = value
        End Set
    End Property
    Public Property AccountName As String
        Get
            Return mstrAccountName
        End Get
        Set(value As String)
            mstrAccountName = value
        End Set
    End Property
    Public Property AccountPass As String
        Get
            Return mstrAccountPass
        End Get
        Set(value As String)
            mstrAccountPass = value
        End Set
    End Property

    Public Property Tvc As String
        Get
            Return mstrTvc
        End Get
        Set(value As String)
            mstrTvc = value
        End Set
    End Property
    Public Property PortalServiceUrl As String
        Get
            Return mstrPortalServiceUrl
        End Get
        Set(value As String)
            mstrPortalServiceUrl = value
        End Set
    End Property
    Public Property BusinessServiceUrl As String
        Get
            Return mstrBusinessServiceUrl
        End Get
        Set(value As String)
            mstrBusinessServiceUrl = value
        End Set
    End Property
End Class
