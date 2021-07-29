Public Class BillingAPI601DALMultiDB
    Inherits BillingAPI601DAL
    Implements IDisposable
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function DropInv007() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = Nothing
                result = String.Format("UPDATE INV007  " &
                                        "  SET ISOBSOLETE = 'Y',OBSOLETEID = {0}0, " &
                                        "  OBSOLETEREASON = (SELECT Description FROM INV006 WHERE IdentifyId1 = '1' AND IdentifyId1 = 0 AND ItemId={0}1), " &
                                       "   CANMODIFY = 'N', UPTTIME = now(), UPTEN = {0}2   WHERE IDENTIFYID1 = '1' " &
                                       "    AND IDENTIFYID2 = 0    AND COMPID = {0}3   AND INVID = {0}4", Sign)

                Return result
            Case Else
                Return MyBase.DropInv007
        End Select


    End Function
End Class
