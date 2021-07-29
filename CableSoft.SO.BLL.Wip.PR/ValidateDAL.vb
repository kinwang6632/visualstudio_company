Public Class ValidateDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable
    Public Sub New()
        
    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
        
    End Sub

    Friend Function GetCust001() As String
        Return String.Format("Select * From SO001 Where Custid = {0}0 ", Sign)
    End Function

    Friend Function GetCust002() As String
        Return String.Format("Select * From SO002 Where Custid = {0}0 and ServiceType = {0}1 ", Sign)
    End Function

    Private Function GetReInstCom(ByVal COMOwner As String) As String
        If String.IsNullOrEmpty(COMOwner) = False AndAlso COMOwner.EndsWith(".") = False Then
            COMOwner &= "."
        End If
        Return COMOwner
    End Function

    Friend Function GetCOMInterface(ByVal COMOwner As String) As String
        Dim strSQL As String
        COMOwner = GetReInstCom(COMOwner)
        strSQL = String.Format("Select * From {1}SO313 Where OCustId = {0}0 And OSNo = {0}1 And OCompCode = {0}2", Sign, COMOwner)
        Return strSQL
    End Function

    Friend Function ChkMustReOpenCommand() As String
        Return String.Format("Select Count(*) From CD089 A,SO007 B Where A.CodeNo = B.InstCode And WorkerType = 'I' And SNo = {0}0", Sign)
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
