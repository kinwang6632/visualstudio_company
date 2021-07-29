Public Class DynamicUpdateDALMultiDB
    Inherits DynamicUpdateDAL
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function getSEQNo(ByVal SourceField As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT sf_getsequenceno('" & SourceField & "')"
            Case Else
                Return MyBase.getSEQNo(SourceField)
        End Select

    End Function
End Class
