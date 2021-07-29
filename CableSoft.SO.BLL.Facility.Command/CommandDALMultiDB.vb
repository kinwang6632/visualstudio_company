Public Class CommandDALMultiDB
    Inherits CommandDAL
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function GetSeqNo(ByVal OwnerName As String, ByVal SourceField As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT sf_getsequenceno('" & OwnerName & SourceField & "')"
            Case Else
                Return MyBase.GetSeqNo(OwnerName, SourceField)
        End Select

    End Function
End Class
