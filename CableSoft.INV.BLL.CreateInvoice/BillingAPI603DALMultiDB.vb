Public Class BillingAPI603DALMultiDB
    Inherits BillingAPI603DAL
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function updInv014() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = Nothing
                result = String.Format(" UPDATE INV014 SET ISOBSOLETE = 'Y', " &
                   " OBSOLETEID = {0}0, " &
                  " OBSOLETEREASON = (SELECT DESCRIPTION FROM INV006 " &
                                                        " WHERE IDENTIFYID1 = '1' AND IDENTIFYID2=0 AND ITEMID= {0}1 " &
                                                        " AND 1=1 )," &
                 " UPTTIME = now(),UPTEN = {0}2 " &
                 " WHERE IDENTIFYID1 = '1'  AND IDENTIFYID2 = 0 " &
                 " AND COMPID = {0}3  AND ALLOWANCENO = {0}4", Sign)
                Return result
            Case Else
                Return MyBase.updInv014
        End Select

    End Function
End Class
