Public Class BillingAPI600DALMultiDB
    Inherits BillingAPI600DAL
    Implements IDisposable
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function getInvSeq() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT LPAD(sf_getsequenceno('S_INV016_SEQ'),7,0)"
            Case Else
                Return MyBase.getInvSeq
        End Select

    End Function
    Friend Overrides Function InsertINV049() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = Nothing
                result = String.Format("Insert into INV049 (SEQ,COMPID,CUSTID,BUSINESSID,TITLE," &
                                            "ZIPCODE,INVADDR,MAILADDR,BEASSIGNEDINVID,SALEAMOUNT," &
                                            "TAXAMOUNT,INVAMOUNT,ISVALID,HOWTOCREATE,UPTTIME, " &
                                            "UPTEN,SHOULDBEASSIGNED,ChargeTitle,ChargeDate, " &
                                            "Memo1,Memo2,LoveNum,A_CarrierId1,A_CarrierId2, " &
                                            "CarrierType,CarrierId1,CarrierId2,TaxRate,TaxType,InvoiceKind," &
                                            "ImportSaleAmount,ImportTaxAmount,ImportInvAmount) VALUES ({0}0,{0}1,{0}2,{0}3,{0}4, " &
                                            "{0}5,{0}6,{0}7,'N',{0}8, " &
                                            "{0}9,{0}10,'Y',3,now()," &
                                            "{0}11,'Y',{0}12,{0}13, " &
                                            "{0}14,{0}15,{0}16,{0}17,{0}18," &
                                            "{0}19,{0}20,{0}21,{0}22,{0}23,1," &
                                            "{0}24,{0}25,{0}26 )", Sign)
                Return result
            Case Else
                Return MyBase.InsertINV049
        End Select

    End Function
End Class
