Imports CableSoft.BLL.Utility
Public Class BillingAPI252DAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub

    Function QueryDVRTryMustPair() As String
        Dim aSQL As String = String.Format("Select Count(*) From SO004,CD043 " & _
                                           "Where SO004.SEQNO={0}0 And " & _
                                           " SO004.MODELCODE = CD043.CODENO  And  CD043.DVRTRYMUSTPAIR = 1", Sign)
        Return aSQL
    End Function
    Public Function QueryWorkerName() As String
        Dim aSQL = String.Format("select empName from CM003 where  empno = {0}0", Sign)
        Return aSQL
    End Function
    Function QueryFaciSeqno() As String
        Dim aSQL As String = String.Format("Select SEQNO From SO004 " & _
                                    " Where facisno = {0}0 " & _
                                    " And ServiceType = {0}1 " & _
                                    " And CustId = {0}2", Sign)
        Return aSQL        
    End Function
    Friend Function QuerySO006() As String
        Return String.Format("Select * From SO006 Where SEQNo = {0}0", Sign)
    End Function
    Friend Function QuertyServiceCode() As String
        Dim aRet As String

        aRet = String.Format("Select CodeNo,Description,Nvl(RefNo,0) RefNo,WorkUnit," & _
                             "Nvl(GroupNo,0) GroupNo, Nvl(ReserveDay,0) ReserveDay " & _
                           " From CD006 Where CodeNo = {0}0 ", Sign)
        Return aRet
    End Function
    Friend Function QueryGroupCode() As String
        Dim aRet As String = String.Format("Select CODENO,Description,1 Flag From CD003 A " & _
                                        " Where Exists (Select * From CD002CM003 B " & _
                                        " Where A.CodeNo = B.EmpNo And ServCode = {0}0 " & _
                                        " And Type = 2) And Nvl(A.StopFlag,0) = 0 ORDER BY CODENO ", Sign)
        Return aRet
    End Function
    Friend Function QueryAcceptEn() As String
        Return String.Format("Select EmpNo,EmpName From CM003 Where EmpNo ={0}0", Sign)
    End Function
    Friend Function getSO009Reinstaddrno() As String
        '#8818 因為有些舊資料有SIGNDATE沒有FINTIME,RETURNCODE,所以再麻煩RD UPDATE時多串fintime is null and returncode is null for debby
        Dim strSQL As String = String.Format("Select ReInstAddrNo,ReInstAddress,servcode," & _
                                             " strtcode,SalesCode,SalesName from SO009 " & _
                                             " Where PrCode in (Select CodeNo From CD007 where refno = 3) " & _
                                             " And ServiceType = {0}0 And Custid = {0}1 And signdate is null " & _
                                             " And FinTime is Null And ReturnCode is Null ", Sign)
        Return strSQL
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
