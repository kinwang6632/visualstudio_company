Public Class CutDateDALMultiDB
    Inherits CutDateDAL
    Implements IDisposable
    Private _Disposed As Boolean

    Public Sub New()
    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
End Class
