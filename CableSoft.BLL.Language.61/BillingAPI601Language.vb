Public Class BillingAPI601Language
    Implements IDisposable
    Public noInvid As String = "發票檔內查無此張發票:{0}"
    Public hasCancel As String = "此張發票:{0}, 已被作廢,不可再被作廢。"
    Public hasLock As String = "此張發票所屬年月: {0}, 已經鎖帳,無法異動。"
    Public hasAllowance As String = "此張發票:{0}, 發票系統內已經產生折讓單, 不可作廢。"
    Public resultMsg As String = "發票號碼:{0}"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                noInvid = Nothing
                hasCancel = Nothing
                hasLock = Nothing
                hasAllowance = Nothing
                resultMsg = Nothing
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
