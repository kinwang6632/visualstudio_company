Public Class ManualNoLanguage
    Implements IDisposable
    Public Property NoPermission As String = "無權限操作此功能"
    Public Property DualError As String = "號碼有重覆,請檢查"
    Public Property ExceedBegin As String = "起始號碼超出截止號碼"
    Public Property hadUse As String = "有其它單據使用中"
    Public Property NoFoundNo As String = "找不到單據"
    Public Property NoneSO127 As String = "此手開單號{0}不存在"
    Public Property HasManualNo As String = "此手開單號{0}已有對應之單據編號:{1}"
    Public Property HadAbandon As String = "此手開單號{0}已作廢"
    Public Property CannotDelete As String = "此組單據號碼,已有對應之回單資料,所以無法刪除"
    Public Property AddClientInfo As String = "手開單新增"
    Public Property EditClientInfo As String = "手開單編輯"
    Public Property VoidClientInfo As String = "手開單作廢"
    Public Property EditNoClientInfo As String = "修改手開單號"
    Public Property ReUseClientInfo As String = "手開單銷單續用"
    Public Property DelClientInfo As String = "手開單刪除"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                NoPermission = Nothing
                DualError = Nothing
                ExceedBegin = Nothing
                hadUse = Nothing
                NoFoundNo = Nothing
                NoneSO127 = Nothing
                HasManualNo = Nothing
                CannotDelete = Nothing
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
