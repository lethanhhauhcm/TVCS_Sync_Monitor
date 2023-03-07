Public Class clsProduct
    Private mstrSeq As String
    Private mstrProdCode As String
    Private mstrProdNo As String
    Private mstrProdName As String
    Private mstrProdUnit As String
    Private mintProdQuantity As Integer
    Private mdecProdPrice As Decimal
    Private mdecAmount As Decimal
    Private mdecDiscountRate As Decimal
    Private mdecDiscountAmount As Decimal
    Private mdecVatRate As Decimal
    Private mdecVatAmount As Decimal
    Private mdecTotalPrice As Decimal
    Private mdecGrandTotal As Decimal
    Private mstrExtra1 As String
    Private mstrExtra2 As String
    Private mstrRemark As String
    Private mintIsSum As Integer
    Private mintTchat As Integer
    Public Property ProdName As String
        Get
            Return mstrProdName
        End Get
        Set(value As String)
            mstrProdName = value
        End Set
    End Property
    Public Property ProdUnit As String
        Get
            Return mstrProdUnit
        End Get
        Set(value As String)
            mstrProdUnit = value
        End Set
    End Property
    Public Property ProdQuantity As Integer
        Get
            Return mintProdQuantity
        End Get
        Set(value As Integer)
            mintProdQuantity = value
        End Set
    End Property
    Public Property ProdPrice As Decimal
        Get
            Return mdecProdPrice
        End Get
        Set(value As Decimal)
            mdecProdPrice = value
        End Set
    End Property
    Public Property Amount As Decimal
        Get
            Return mdecAmount
        End Get
        Set(value As Decimal)
            mdecAmount = value
        End Set
    End Property

    Public Property DiscountRate As Decimal
        Get
            Return mdecDiscountRate
        End Get
        Set(value As Decimal)
            mdecDiscountRate = value
        End Set
    End Property
    Public Property DiscountAmount As Decimal
        Get
            Return mdecDiscountAmount
        End Get
        Set(value As Decimal)
            mdecDiscountAmount = value
        End Set
    End Property
    Public Property VatRate As Decimal
        Get
            Return mdecVatRate
        End Get
        Set(value As Decimal)
            mdecVatRate = value
        End Set
    End Property
    Public Property VatAmount As Decimal
        Get
            Return mdecVatAmount
        End Get
        Set(value As Decimal)
            mdecVatAmount = value
        End Set
    End Property
    Public Property TotalPrice As Decimal
        Get
            Return mdecTotalPrice
        End Get
        Set(value As Decimal)
            mdecTotalPrice = value
        End Set
    End Property
    Public Property GrandTotal As Decimal
        Get
            Return mdecGrandTotal
        End Get
        Set(value As Decimal)
            mdecGrandTotal = value
        End Set
    End Property
    Public Property ProductCode As String
        Get
            Return mstrProdCode
        End Get
        Set(value As String)
            mstrProdCode = value
        End Set
    End Property

    Public Property Extra1 As String
        Get
            Return mstrExtra1
        End Get
        Set(value As String)
            mstrExtra1 = value
        End Set
    End Property
    Public Property Extra2 As String
        Get
            Return mstrExtra2
        End Get
        Set(value As String)
            mstrExtra2 = value
        End Set
    End Property
    Public Property Remark As String
        Get
            Return mstrRemark
        End Get
        Set(value As String)
            mstrRemark = value
        End Set
    End Property
    Public Property ProdNo As String
        Get
            Return mstrProdNo
        End Get
        Set(value As String)
            mstrProdNo = value
        End Set
    End Property
    Public Property IsSum As Integer
        Get
            Return mintIsSum
        End Get
        Set(value As Integer)
            mintIsSum = value
        End Set
    End Property
    Public Property TChat As Integer
        Get
            Return mintTChat
        End Get
        Set(value As Integer)
            mintTChat = value
        End Set
    End Property

    Public Property Seq As String
        Get
            Return mstrSeq
        End Get
        Set(value As String)
            mstrSeq = value
        End Set
    End Property
End Class
