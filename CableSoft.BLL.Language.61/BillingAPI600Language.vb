Public Class BillingAPI600Language
    Implements IDisposable
    Public NoInvChr As String = "取不到發票字軌"
    Public getMultiInvChr As String = "取到多筆發票字軌"
    Public noINV003 As String = "找不到INV003設定檔"
    Public noSysID As String = "SysID未設定"
    Public needServiceType As String = "來源資料品名{0},服務別需有值"
    Public needItem As String = "來源資料品名{0},收費單號項次需有值"
    Public differItemSource As String = "來源資料品名{0}與現有資料不符"
    Public differAmount As String = "發票金額與明細不符!"
    Public noAnyInv As String = "無任何發票資料產生"
    Public resultMsg As String = "客戶編號:{0},發票號碼:{1}"
    Public excuteSFFail As String = "Excute store function failed"
    Public differItemTaxCode As String = "來源資料品名{0}稅別不符"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                NoInvChr = Nothing
                getMultiInvChr = Nothing
                noINV003 = Nothing
                noSysID = Nothing
                needServiceType = Nothing
                needItem = Nothing
                differItemSource = Nothing
                differAmount = Nothing
                noAnyInv = Nothing
                resultMsg = Nothing
                excuteSFFail = Nothing
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
