Public Class clsInvToken
    Private mstrHashValue As String
    Private mstrKeyInv As String
    Private mstrIdInv As String

    Public Property HashValue As String
        Get
            Return mstrHashValue
        End Get
        Set(value As String)
            mstrHashValue = value
        End Set
    End Property
    Public Property KeyInv As String
        Get
            Return mstrKeyInv
        End Get
        Set(value As String)
            mstrKeyInv = value
        End Set
    End Property
    Public Property IdInv As String
        Get
            Return mstrIdInv
        End Get
        Set(value As String)
            mstrIdInv = value
        End Set
    End Property
End Class
