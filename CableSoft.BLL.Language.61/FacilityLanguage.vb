Public Class FacilityLanguage
    Implements IDisposable
    Public Property CanNotEdit As String = "BPCode 有值,不允許修改!"
    Public Property CanNotDel As String = "BPCode 有值,不允許刪除!"
    Public Property MustServiceType As String = "服務別為必要欄位需有值!"
    Public Property MustFaciName As String = "設備名稱為必要欄位需有值!"
    Public Property MustBuyType As String = "買賣方式為必要欄位需有值!"
    Public Property MustWorkType As String = "派工類別為必要欄位需有值!"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                CanNotEdit = Nothing
                CanNotDel = Nothing
                MustServiceType = Nothing
                MustFaciName = Nothing
                MustBuyType = Nothing
                MustWorkType = Nothing
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
