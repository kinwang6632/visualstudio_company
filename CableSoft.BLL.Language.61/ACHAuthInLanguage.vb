Public Class ACHAuthInLanguage
    Implements IDisposable
    Public Property FormatError As String = "格式錯誤 !"
    Public Property ReplyFormatError As String = "回覆格式錯誤 !"
    Public Property ACHCustIdNotInDB As String = "此筆ACHCustid 未存在於資料庫"
    Public Property NotFoundSO106A As String = "找不到SO106A對應資料,該筆資料可能已授權或取消授權 "
    Public Property UpdSO106AError As String = "未成功更新SO106! 可能為該筆資料已有核准日期..請確認用戶識別碼、帳戶、核准日期或送出日期是否錯誤 !"
    Public Property InsSO002AError As String = "未成功新增SO002A"    
    Public Property UpdSO003Error As String = "未成功更新SO003"
    Public Property UpdNoneSO003Error As String = "未成功更新非週期SO003"
    Public Property UpdSO003CError As String = "未成功更新SO003C"
    Public Property InsSO004Error As String = "未成功新增SO004"
    Public Property UpdSO106AError2 As String = "未成功更新SO106A"
    Public Property UpdSO106Error As String = "未成功更新SO106!請確認用戶識別碼、帳戶、核准日期或送出日期是否錯誤 !"
    Public Property StopSO003Error As String = "停用SO003資料錯誤!"
    Public Property StopNonePeriodError As String = "停用非週期SO003資料錯誤!"
    Public Property StopSO003CError As String = "停用SO003C資料錯誤!"
    Public Property RunTotalRecord As String = "已完成資料筆數共{0}筆,"
    Public Property RunErrorRecord As String = "問題筆數共{1}筆,"
    Public Property RunSpendTime As String = "共花費:{2}秒"
    Public Property NoPermission As String = "無權限操作此功能"
    Public Property AuthClientInfo As String = "ACH授權"
    Public Property CancelClientInfo As String = "ACH取消授權"
    Public Property OldClientInfo As String = "新增舊有已簽約"
    Public Property GetReStatusP As String = "P:已發送授權書及授權扣款檔"
    Public Property GetReStatusR As String = "R:先收到回覆訊息但未收到授權書"
    Public Property GetReStatusY As String = "Y:先收到回覆訊息後收到授權書"
    Public Property GetReStatusM As String = "M:先收到授權書但未收到回覆訊息"
    Public Property GetReStatusS As String = "S:先收到授權書後收到回覆訊息"
    Public Property GetReStatusC As String = "C:已收到舊件轉檔回覆訊息"
    Public Property GetReStatusD As String = "D:已收到取消授權扣款回覆訊息"
    Public Property failDate As String = " 失敗日期:"
    Public Property AuthErrorMsg1 As String = " 印鑑不符!! "
    Public Property AuthErrorMsg2 As String = "無此帳號 !! "
    Public Property AuthErrorMsg3 As String = "委繳戶統編不存在 !! "
    Public Property AuthErrorMsg4 As String = "資料重覆 !! "
    Public Property AuthErrorMsg5 As String = "原交易不存在 !! "
    Public Property AuthErrorMsg6 As String = "電子資料與授權書內容不符 !! "
    Public Property AuthErrorMsg7 As String = "帳戶已結清 !! "
    Public Property AuthErrorMsg8 As String = "印鑑不清 !! "
    Public Property AuthErrorMsgA As String = "未收到授權書 !!"
    Public Property AuthErrorMsgB As String = "用戶號碼錯誤 !!"
    Public Property AuthErrorMsgC As String = "靜止戶 !!"
    Public Property AuthErrorMsgD As String = "未收到聲明書 !!"
    Public Property AuthErrorMsg9 As String = "其他不成功原因 !! "
    Public Property GetErrMsg As String = " ID: {0}; 帳號:{1}; 失敗原因:{2}"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                FormatError = Nothing
                ReplyFormatError = Nothing
                ACHCustIdNotInDB = Nothing
                NotFoundSO106A = Nothing
                UpdSO106AError = Nothing
                InsSO002AError = Nothing
                UpdSO003Error = Nothing
                InsSO004Error = Nothing
                UpdSO106AError2 = Nothing
                UpdSO106Error = Nothing
                StopSO003Error = Nothing
                RunTotalRecord = Nothing
                RunErrorRecord = Nothing
                RunSpendTime = Nothing
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
