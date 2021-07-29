Public Class BillingAPI602DALMultiDB
    Inherits BillingAPI602DAL
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function GetAllowanceNo() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "select to_char( now()  , 'yyyymm' ) ||  lpad(sf_getsequenceno('s_inv014_allowance') ||'', 6, '0' )"
            Case Else
                Return MyBase.GetAllowanceNo
        End Select

    End Function
End Class
