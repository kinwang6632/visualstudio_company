Public Class BillingAPI235DALMultiDB
    Inherits BillingAPI235DAL
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
End Class
