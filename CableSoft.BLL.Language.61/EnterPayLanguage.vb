Public Class EnterPayLanguage
    Implements IDisposable
    Public Property SNo As String = "工單單號"
    Public Property NotFoundNo As String = "無此手開單號,請查證!!"
    Public Property HaveCancelNo As String = "此手開單號已作廢不可使用,請查證!!"
    Public Property HaveBillNo As String = "該手開單已有收費資料,請查證!!"
    Public Property HaveDiscount As String = "有第2層優惠不能作廢!!"
    Public Property HasBillNo As String = "收費資料已收"
    Public Property HasCancel As String = "收費資料已作廢"
    Public Property UpdBillError As String = "回填週期性收費資料有誤"
    Public Property NotFoundBillNo As String = "無此單據編號或此單據已收款!!"
    Public Property HasInput As String = "此單據已登錄過！請重新輸入！"
    Public Property PayForCounter As String = "此單據櫃臺已收！請重新輸入！"
    Public Property MustSTReason As String = "實收金額與出單金額不相等,須有短收原因!"
    Public Property CustidNotNormal As String = "此客戶狀態已非正常收視戶, 請檢查"
    Public Property SNoLenError As String = "單據長度錯誤!!"
    Public Property CanNotCancel As String = "該單據編號有資料已收或作廢, 不允許取消登錄!!"
    Public Property DateIsIllegal As String = "日期不合法！"
    Public Property OverToday As String = "此日期超過今天日期！"
    Public Property OverSafeDay As String = "此日期已超過系統設定的安全期限！"
    Public Property HadClosed As String = "之前已做過日結，不可使用之前日期入帳"
    Public Property ClientInfoString = "收費單登錄"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                SNo = Nothing
                NotFoundNo = Nothing
                HaveCancelNo = Nothing
                HaveBillNo = Nothing
                HaveDiscount = Nothing
                HasBillNo = Nothing
                HasCancel = Nothing
                UpdBillError = Nothing
                NotFoundBillNo = Nothing
                HasInput = Nothing
                MustSTReason = Nothing
                CustidNotNormal = Nothing
                SNoLenError = Nothing
                CanNotCancel = Nothing
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
