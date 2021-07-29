Public Class CopyOrderDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)

    End Sub
    Friend Function ChkReturnCode() As String
        Return String.Format("SELECT RETURNCODE,CUSTID FROM SO105 WHERE ORDERNO ={0}0", Sign)
    End Function
    Friend Function UpdateFaciSeqNo(ByVal TableName As String) As String
        Dim aSQL As String = String.Format("Update " & TableName & " Set FaciSeqNo={0}0 " & _
                                         " Where OrderNo={0}1 AND FaciSeqNo = {0}2", Sign)
        Return aSQL
    End Function
    Friend Function GetCustId() As String
        Return String.Format("SELECT CUSTID FROM SO105 WHERE ORDERNO = {0}0", Sign)
    End Function
    Friend Function NoCopySO105Detail() As String
        Dim aRet As String = Nothing
        aRet = String.Format("SELECT distinct * FROM SO105DETAIL WHERE WORKERTYPE = 'I' " &
                           " AND ORDERNO = {0}0 " &
                           " AND NVL(SNO,'X') IN (SELECT SNO FROM SO007 WHERE MAINSNO IS NOT NULL " &
                           " AND INSTCODE IN (SELECT CODENO FROM CD005 WHERE REFNO = 12 ))", Sign)
        'aRet = String.Format("Select * From SO105DETAIL Where SNO = '201706IC2173420' and orderno={0}0", Sign)
        Return aRet
    End Function
    Friend Function GetOrder() As ArrayList
        Dim arraySQL As New ArrayList
        '訂單資料檔
        arraySQL.Add(String.Format("Select * From SO105 Where OrderNo = {0}0", Sign))
        '訂單資料檔-訂購人員資料
        arraySQL.Add(String.Format("Select * From SO105D1 Where OrderNo = {0}0", Sign))
        '訂單資料檔-產品檔
        arraySQL.Add(String.Format("Select * From SO105B Where OrderNo = {0}0", Sign))
        '訂單資料檔-商贈品檔
        arraySQL.Add(String.Format("Select * From SO105C Where OrderNo = {0}0", Sign))
        '訂單資料檔-設備檔
        arraySQL.Add(String.Format("Select * From SO105D Where OrderNo = {0}0", Sign))
        '訂單資料檔-派工資料
        arraySQL.Add(String.Format("Select * From SO105DETAIL Where OrderNo = {0}0", Sign))
        '結清收費記錄檔
        arraySQL.Add(String.Format("Select * From SO105I Where OrderNo = {0}0", Sign))
        '訂單管理-產品
        arraySQL.Add(String.Format("Select * From SO105J Where OrderNo = {0}0", Sign))
        '訂單管理-續約產品
        arraySQL.Add(String.Format("Select * From SO105K Where OrderNo = {0}0", Sign))
        Return arraySQL
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
