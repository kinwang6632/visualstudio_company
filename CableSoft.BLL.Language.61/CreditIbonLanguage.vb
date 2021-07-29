
Public Class CreditIBonLanguage
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Shared Property fStatusNot As String = "未處理"
    Public Shared Property fStatusYes As String = "已處理"
    Public Shared Property fNotSMS As String = "與SMS不符"
    Public Shared Property fNotAmount As String = "金額不符"
    Public Shared Property fCancelFlag As String = "已作廢"
    Public Shared Property fUCCode As String = "收費已收"
    Public Shared Property fNoCharge As String = "查無費用"
    Public Shared Property fCmdError As String = "命令失敗"

    Public Shared Property NoConnStr As String = "未傳入連線字串!"
    Public Shared Property sNotConnMsg As String = "無法連線:第 {0} 次!{1}"
    Public Shared Function NotConnMsg(Count As Integer) As String
        Return String.Format(sNotConnMsg, Count, ControlChars.CrLf)
    End Function
    Public Shared Property sNotConnStrMsg As String = " {0} ,連線字串: {1} {2}"
    Public Shared Function NotConnStrMsg(ErrLog As String, ConnStr As String) As String
        Return String.Format(sNotConnStrMsg, ErrLog, ConnStr, ControlChars.CrLf)
    End Function



    'Public Shared Property sSuccessMsg As String = "共更新{1}成功資料{0}筆!{1}異常資料{2}筆!"
    'Public Shared Function SuccessMsg(SCount As Integer, ECount As Integer) As String
    '    Return String.Format(sSuccessMsg, SCount, ControlChars.CrLf, ECount)
    'End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                fStatusNot = Nothing
                fStatusYes = Nothing
                fNotSMS = Nothing
                fNotAmount = Nothing
                fCancelFlag = Nothing
                fUCCode = Nothing
                NoConnStr = Nothing
                sNotConnMsg = Nothing
                sNotConnStrMsg = Nothing
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時,才覆寫 Finalize()。
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
