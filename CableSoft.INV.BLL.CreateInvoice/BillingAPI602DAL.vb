Imports CableSoft.BLL.Utility
Public Class BillingAPI602DAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryInv007() As String
        Return String.Format("SELECT A.CUSTID, A.BUSINESSID, A.INVFORMAT, " &
                " A.INVDATE, A.INVAMOUNT, " &
                " NVL(A.TAXTYPE,0) AS TAXTYPE, NVL(A.ISOBSOLETE,'N') AS ISOBSOLETE " &
                "  FROM INV007 A " &
                " WHERE A.IDENTIFYID1 = '1' " &
                " AND A.IDENTIFYID2 = 0 " &
                " AND A.INVID = {0}0  AND A.COMPID = {0}1", Sign)
    End Function
    Friend Overridable Function GetAllowanceNo() As String
        Return "select to_char( sysdate, 'yyyymm' ) || " &
                    " lpad( s_inv014_allowance.nextval, 6, '0' ) " &
                    " from dual "
    End Function
    Friend Function QueryInv014A() As String
        Return String.Format("SELECT * FROM INV014A WHERE ALLOWANCENO = {0}0", Sign)
    End Function
    Friend Function QueryExcludeBill(ByVal aOwner As String, ByVal aBillItem As String) As String
        Dim result As String = Nothing
        result = String.Format("SELECT A.SERVICETYPE,A.BILLNO,A.ITEM, " & _
                                "A.CITEMCODE,A.CITEMNAME,A.PTCODE,A.PTNAME, " & _
                                "A.REALAMT,A.NOTE FROM " & aOwner & ".SO034 A," & aOwner & ".SO001 B, " & aOwner & ".CD019 C " & _
                                " WHERE A.CUSTID = B.CUSTID AND A.CITEMCODE = C.CODENO AND A.PREINVOICE = 3 " & _
                                " AND A.GUINO = {0}0 AND A.COMPCODE = {0}1 " & _
                                " And BillNo || ITEM NOT IN(" & aBillItem & ")", Sign)
        Return result
    End Function
    Friend Function QuerySO034(ByVal aOwner As String) As String
        Dim result As String = Nothing
        result = String.Format("SELECT A.SERVICETYPE,A.BILLNO,A.ITEM, " & _
                               "A.CITEMCODE,A.CITEMNAME,A.PTCODE,A.PTNAME, " & _
                               "A.REALAMT,A.NOTE FROM " & aOwner & ".SO034 A," & aOwner & ".SO001 B, " & aOwner & ".CD019 C " & _
                               " WHERE A.CUSTID = B.CUSTID AND A.CITEMCODE = C.CODENO AND A.PREINVOICE = 3 " & _
                               " AND A.GUINO = {0}0 AND A.COMPCODE = {0}1", Sign)

        Return result
    End Function
    Friend Function QueryIsDataLocked() As String
        Return String.Format("SELECT NVL(COUNT(*),0) COUNT FROM INV018 WHERE IDENTIFYID1 = '1' " & _
                " AND IDENTIFYID2 = 0 AND COMPID = {0}0   " & _
                " AND YEARMONTH = {0}1 AND NVL(ISLOCKED,'N') = 'Y' ", Sign)
    End Function
    Friend Function updInv014A() As String
        Dim result As String = Nothing
        result = String.Format("Update INV014A SET InvAmount = {0}0,UptTime={0}1,UptEn={0}2 " & _
                              " Where AllowanceNo = {0}3 And seq= 0 ", Sign)
        Return result
    End Function
    Friend Function insInv014A() As String
        Dim result As String = Nothing
        result = String.Format("INSERT INTO INV014A ( ServiceTypeID,PaperNo, Seq,ItemId, " & _
                               "Description,PayTypeId,PayTypeDesc, " & _
                                "InvAmount, UptTime, UptEn, " & _
                                " AllowanceNo) " & _
                                " Values ( {0}0,{0}1,{0}2,{0}3, " & _
                                "{0}4,{0}5,{0}6, " & _
                                "{0}7,{0}8,{0}9, " & _
                                "{0}10) ", Sign)
       
        Return result
    End Function
    Friend Function UpdInv014() As String
        Dim result As String = Nothing
        result = String.Format("UPDATE INV014 set InvAmount = {0}0, " & _
                                    "UPTTIME = {0}1,UPTEN = {0}2 ,SaleAmount = {0}3,TaxAmount={0}4 " & _
                                    " Where IDENTIFYID1= '1' " & _
                                    " AND IDENTIFYID2 = 0 AND INVID = {0}5 AND ALLOWANCENO = {0}6 AND COMPID= {0}7 ", Sign)
        Return result
    End Function
    Friend Function updSO034(ByVal aOwner As String) As String
        Dim result As String = Nothing
        result = String.Format("update " & aOwner & ".so034   set " &
                " preinvoice = 4, " &
                " invoicetime = {0}0, " &
                " note = {0}1 " &
                " where  1 = 1 " &
                " and billno = {0}2 " &
                " and item = {0}3 " &
                " and compcode = {0}4 " &
                " and preinvoice = 3 ", Sign)
        Return result
    End Function
    Friend Function InsInv014() As String
        Dim result As String = Nothing
        result = String.Format("INSERT INTO INV014 ( IDENTIFYID1, IDENTIFYID2, COMPID," & _
                                               " CUSTID, PAPERDATE, BUSINESSID,  " & _
                                               " YEARMONTH, INVID, SEQ, " & _
                                               " INVFORMAT, INVDATE, TAXTYPE, " & _
                                               " SALEAMOUNT, TAXAMOUNT, INVAMOUNT,  " & _
                                               " UPTTIME, UPTEN, ALLOWANCENO, " & _
                                               " SOURCE,PaperNo,UPLOADFLAG )  " & _
                                               "  values ( '1', 0, {0}0, " & _
                                                "{0}1,{0}2,{0}3," & _
                                                "{0}4,{0}5,{0}6," & _
                                                "{0}7,{0}8,{0}9," & _
                                                "{0}10,{0}11,{0}12," & _
                                                "{0}13,{0}14,{0}15, " & _
                                                "{0}16,{0}17,{0}18)", Sign)

        Return result
    End Function
    Friend Function QueryDualInv014() As String
        Dim result As String = Nothing
        result = String.Format("SELECT ALLOWANCENO " & _
                                    " From INV014 " & _
                                    " WHERE IDENTIFYID1 = '1' AND IDENTIFYID2 = 0 " & _
                                    " AND INVID = {0}0 AND NVL(UPLOADFLAG,0) = 0 " & _
                                    " AND PAPERDATE = {0}1 AND COMPID = {0}2", Sign)
        Return result
    End Function
    Friend Function QuerySumAmt() As String
        Dim result As String = Nothing
        result = String.Format("SELECT NVL(SUM(A.INVAMOUNT),0) AS SUM FROM INV014 A " &
                             " WHERE A.IDENTIFYID1 ='1' AND A.IDENTIFYID2 = 0 " &
                             " AND A.INVID = {0}0 AND NVL(A.ISOBSOLETE,'N') = 'N' " &
                             " AND COMPID = {0}1", Sign)
        Return result
    End Function
    Friend Function QueryUpdlimtDate() As String
        Return String.Format("SELECT NVL(UPDLIMTDATE,-1) FROM INV001 A " &
                             " WHERE A.IDENTIFYID1 ='1' AND A.IDENTIFYID2 = 0  AND COMPID ={0}0", Sign)
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
