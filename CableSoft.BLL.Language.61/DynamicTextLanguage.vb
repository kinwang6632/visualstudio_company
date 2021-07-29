Public Class DynamicTextLanguage
    Implements IDisposable

    Public Property NoSO1101AData As String = "{0} 代號找不到任何SO1109A資料!"
    Public Property NoSO1101BData As String = "{0} ProgramId 找不到任何SO1109B資料!"
    Public Property NoAnyData As String = "無任何資料"
    Public Property RunOK As String = "執行完成!!"
    Public Property RunTotalRecord As String = "共執行：{0} 筆 "
    Public Property RunSucessRecord As String = "成功：{1} 筆"
    Public Property RunFailRecord As String = "失敗：{2} 筆"
    Public Property RunSucessAmt As String = "成功總金額合計：{3}"
    Public Property RunSpendTime As String = "共花費：{4} 秒"
    Public Property RunSpendTime2 As String = "共花費：{3} 秒"
    Public Property NoFundMaster As String = "找不到SysProgramId代號"
    Public Property NoPermission As String = "無權限操作此功能"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                NoSO1101AData = Nothing
                NoSO1101BData = Nothing
                NoAnyData = Nothing
                RunOK = Nothing
                RunTotalRecord = Nothing
                RunSucessRecord = Nothing
                RunFailRecord = Nothing
                RunSucessAmt = Nothing
                RunSpendTime = Nothing
                RunSpendTime2 = Nothing
                NoFundMaster = Nothing
                NoPermission = Nothing
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
