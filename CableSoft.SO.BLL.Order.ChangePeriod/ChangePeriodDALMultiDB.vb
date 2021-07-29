Public Class ChangePeriodDALMultiDB
    Inherits ChangePeriodDAL
    Implements IDisposable
    Public Sub New()
    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
End Class
