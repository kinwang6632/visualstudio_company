Public Class BillingAPI259DALMultiDB
    Inherits BillingAPI259DAL
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)

    End Sub
    Friend Overrides Function GetContactDetailData() As String
        Dim strSQL As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                strSQL = String.Format("Select * From (Select * From SO006A Where SeqNo = {0}0 Order By AutoSerialNo) as A limit 1", Sign)
            Case Else
                strSQL = MyBase.GetContactDetailData
        End Select

        Return strSQL
    End Function
End Class
