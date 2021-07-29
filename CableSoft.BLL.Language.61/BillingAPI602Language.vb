Public Class BillingAPI602Language
    Implements IDisposable
    Public DiscountLargerInv As String = "折讓金額不可>原發票金額"
    Public SumDiscountLargerInv As String = "折讓金額累計不可>原發票金額"
    Public onlyCanDrop As String = "當期發票全額折讓只能用作廢方式"
    Public noFoundInvId As String = "客服系統找不到此發票!"
    Public notSameCust As String = "折讓金額與客服系統不符!"
    Public noLargerZero As String = "客服系統退費折讓金額不可>0。"
    Public beLock As String = "此筆資料所屬公司別已經鎖帳,無法存檔!"
    Public retMsg As String = "發票號碼{0},折讓單號{1}"
    Public noteMsg As String = "折讓單號:{0}"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                DiscountLargerInv = Nothing
                SumDiscountLargerInv = Nothing
                onlyCanDrop = Nothing
                noFoundInvId = Nothing
                notSameCust = Nothing
                noLargerZero = Nothing
                beLock = Nothing
                retMsg = Nothing
                noteMsg = Nothing
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
