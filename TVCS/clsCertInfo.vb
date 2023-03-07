Public Class clsCertInfo
    Private mstrOwnCA As String
    Private mstrOrganizationCA As String
    Private mstrSerialNumber As String
    Private mdteValidFrom As Date
	Private mdteValidTo As Date

	Public Property OwnCA As String
		Get
			Return mstrOwnCA
		End Get
		Set(value As String)
			mstrOwnCA = value
		End Set
	End Property
	Public Property OrganizationCA As String
		Get
			Return mstrOrganizationCA
		End Get
		Set(value As String)
			mstrOrganizationCA = value
		End Set
	End Property
	Public Property SerialNumber As String
		Get
			Return mstrSerialNumber
		End Get
		Set(value As String)
			mstrSerialNumber = value
		End Set
	End Property
	Public Property ValidFrom As Date
		Get
			Return mdteValidFrom
		End Get
		Set(value As Date)
			mdteValidFrom = value
		End Set
	End Property
	Public Property ValidTo As Date
		Get
			Return mdteValidto
		End Get
		Set(value As Date)
			mdteValidto = value
		End Set
	End Property
End Class
