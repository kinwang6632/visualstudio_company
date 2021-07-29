Public Class BillingAPI516Language
    Implements IDisposable
    Public mustMediaCode As String = "介紹媒介需有值"
    Public mustAccountName As String = "帳號所有人"
    Public cannotSendDate As String = "銀行代碼:[{0}]不可傳送件日期"
    Public cannotSnactionDate As String = "銀行代碼:[{0}]不可傳核准日期"
    Public noFoundACHTNO As String = "ACH交易代碼無值"
    Public cannotACHTNO As String = "非ACH銀行,不允許有ACHTNO"
    Public cannotStopDate As String = "新增狀態停用日期不允許有值"
    Public cannotDeAuthorize As String = "新增不允許取消授權"
    Public noFoundIntroName As String = "介紹人不存在"
    Public musAchtNo As String = "ACH交易代碼需傳值"
    Public noFoundInv As String = "找不到發票資訊檔"
    Public cannotModiACHAccount As String = "ACH、郵局類,已有送件日期不可修改帳號"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                cannotSnactionDate = Nothing
                mustMediaCode = Nothing
                mustAccountName = Nothing
                cannotSendDate = Nothing
                noFoundACHTNO = Nothing
                cannotACHTNO = Nothing
                cannotStopDate = Nothing
                cannotDeAuthorize = Nothing
                noFoundIntroName = Nothing
                noFoundInv = Nothing
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
