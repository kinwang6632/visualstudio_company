Public Class BillingAPI514Language
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
    Public ACHTNONotSameBank As String = "ACHTNO與銀行別不符"
    Public ACHTNONotSameBank2 As String = "{0} ACHTNO 與銀行別不符!"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                cannotSnactionDate = Nothing
                mustMediaCode = Nothing
                mustAccountName = Nothing
                cannotSendDate = Nothing
                noFoundACHTNO = Nothing
                cannotACHTNO = Nothing
                cannotStopDate = Nothing
                cannotDeAuthorize = Nothing
                noFoundIntroName = Nothing
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
