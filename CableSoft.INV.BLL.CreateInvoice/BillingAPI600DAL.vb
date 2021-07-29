Imports CableSoft.BLL.Utility
Public Class BillingAPI600DAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryInv099() As String
        Return String.Format("Select * from INV099 where " &
                    "Compid = {0}0 And UseKind = 3 And UploadFlag = 1  " &
                    " And yearmonth= {0}1", Sign)
    End Function
    Friend Function GetEmpName() As String
        Dim strSQL As String
        strSQL = String.Format("Select EmpName From CM003 Where EmpNo = {0}0", Sign)
        Return strSQL
    End Function
    Friend Overridable Function getInvSeq() As String
        Return "SELECT LPAD(S_INV016_SEQ.NEXTVAL,7,0) FROM DUAL"
    End Function
    Friend Overridable Function InsertINV049() As String
        Dim result As String = Nothing
        result = String.Format("Insert into INV049 (SEQ,COMPID,CUSTID,BUSINESSID,TITLE," &
                                            "ZIPCODE,INVADDR,MAILADDR,BEASSIGNEDINVID,SALEAMOUNT," &
                                            "TAXAMOUNT,INVAMOUNT,ISVALID,HOWTOCREATE,UPTTIME, " &
                                            "UPTEN,SHOULDBEASSIGNED,ChargeTitle,ChargeDate, " &
                                            "Memo1,Memo2,LoveNum,A_CarrierId1,A_CarrierId2, " &
                                            "CarrierType,CarrierId1,CarrierId2,TaxRate,TaxType,InvoiceKind," &
                                            "ImportSaleAmount,ImportTaxAmount,ImportInvAmount) VALUES ({0}0,{0}1,{0}2,{0}3,{0}4, " &
                                            "{0}5,{0}6,{0}7,'N',{0}8, " &
                                            "{0}9,{0}10,'Y',3,SYSDATE," &
                                            "{0}11,'Y',{0}12,{0}13, " &
                                            "{0}14,{0}15,{0}16,{0}17,{0}18," &
                                            "{0}19,{0}20,{0}21,{0}22,{0}23,1," &
                                            "{0}24,{0}25,{0}26 )", Sign)
        Return result
    End Function
    Friend Function InsertINV050() As String
        Dim result As String = Nothing
        result = String.Format("Insert Into INV050(SEQ,BILLID,BILLIDITEMNO,TAXTYPE,CHARGEDATE," & _
                                   "ITEMID,DESCRIPTION,QUANTITY,UNITPRICE,TAXRATE," & _
                                   "TAXAMOUNT,TOTALAMOUNT,STARTDATE,ENDDATE,CHARGEEN," & _
                                   "SHOULDBEASSIGNED,LINKTOMIS,SERVICETYPE ) Values ({0}0,{0}1,{0}2,{0}3,{0}4, " & _
                                   "{0}5,{0}6,{0}7,{0}8,{0}9," & _
                                   "{0}10,{0}11,{0}12,{0}13,{0}14," & _
                                   "'Y',{0}15,{0}16)", Sign)


        Return result
    End Function
    
    Friend Function QueryINV003Param() As String
        Return String.Format("SELECT INVPARAM FROM INV003 WHERE COMPID = {0}0", Sign)
    End Function
    Friend Function QueryINV005Name() As String
        Return String.Format("Select Description From INV005 where ItemId ={0}0 And CompID = {0}1", Sign)
    End Function
    Friend Function QueryINV005TaxCode() As String
        Return String.Format("Select TaxCode From INV005 where ItemId ={0}0 And CompID = {0}1", Sign)
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下方的 Finalize()。
            ' TODO: 將大型欄位設為 null。
        End If
        disposedValue = True
    End Sub

    ' TODO: 只有當上方的 Dispose(disposing As Boolean) 具有要釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Visual Basic 加入這個程式碼的目的，在於能正確地實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
        Dispose(True)
        ' TODO: 覆寫上列 Finalize() 時，取消下行的註解狀態。
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
