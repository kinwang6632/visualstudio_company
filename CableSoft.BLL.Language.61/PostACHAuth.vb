Public Class PostACHAuth
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Property NoACHCustId As String = "ID: {0} 帳號 : {1} 沒有ACH用戶號碼"
    Public Property ProcResult As String = "需處理筆數共{0}筆," & vbCrLf & vbCrLf & _
            "已完成資料筆數共{1}筆," & vbCrLf & vbCrLf & _
            "問題筆數共{2}筆," & vbCrLf & vbCrLf & _
            "共花費:{3}秒"
    Public Property NoData As String = "無任何資料可提出！"
    Public Property NoFoundCitem As String = "ID: {0} 帳號 : {1} 沒有指定扣款週期收費項目"
    Public Property chkFileFail As String = "該檔案不符合提回格式！"
    Public Property NoExistsACHCustId As String = "此筆ACHCustid 未存在於資料庫"
    Public Property NoFoundSO106A As String = "找不到SO106A對應資料,該筆資料可能已授權或取消授權"
    Public Property CommonErr As String = "ID: {0}; 帳號:{1}; 失敗原因:{2}"
    Public Property UpdNoneSO003Error As String = "未成功更新非週期SO003"
    Public Property UpdSO003CError As String = "未成功更新SO003C"
    Public Property UpdSO003Error As String = "未成功更新SO003"
    Public Property UpdSO106AError As String = "未成功更新SO106A"
    Public Property UpdNote As String = "失敗日期:"
    Public Property UpdSO106 As String = "未成功更新SO106"
    Public Property ClearSO106Fail As String = "未成功清除SO106"
    Public Property UpdCancelAuthNote As String = "日期: {0} {1}"
    Public Property StopSO003Error As String = "停用SO003錯誤"
    Public Property StopNonePeriodError As String = "停用非週期錯誤"
    Public Property StopSO003CError As String = "停用SO003C錯誤"
    Public Property ResumeDataStatus As String = "郵局誤終止扣款,已回復為申請!"
    Public Property ResumeCitemCode As String = "回復收費項目有誤!"
    Public Property NoPermission As String = "無權限操作此功能"
    Public Property NoCanedit As String = "參數尚未設定，無法使用"
    Public Property resultCount As String = "成功筆數 :{0} 筆 " & Environment.NewLine & "錯誤筆數:{1} 筆"
    Public Property ClientInfoString As String = "富邦ACH提回"
    Public Property ApplyErrMsg03 As String = "已終止代繳"
    Public Property ApplyErrMsg06 As String = "凍結戶或警示戶"
    Public Property ApplyErrMsg07 As String = "業務支票專戶"
    Public Property ApplyErrMsg08 As String = "帳號錯誤"
    Public Property ApplyErrMsg09 As String = "終止戶"
    Public Property ApplyErrMsg10 As String = "身分證號不符"
    Public Property ApplyErrMsg11 As String = "轉出戶"
    Public Property ApplyErrMsg12 As String = "拒絕往來戶"
    Public Property ApplyErrMsg13 As String = "無此用戶編號"
    Public Property ApplyErrMsg14 As String = "用戶編號已存在"
    Public Property ApplyErrMsg16 As String = "管制帳戶"
    Public Property ApplyErrMsg17 As String = "掛失戶"
    Public Property ApplyErrMsg18 As String = "異常交易帳戶"
    Public Property ApplyErrMsg19 As String = "用戶編號非英數字元"
    Public Property ApplyErrMsg91 As String = "用戶編號非英數字元"
    Public Property ApplyErrMsg98 As String = "其他"
    Public Property otherErrMsg1 As String = "局帳號不符"
    Public Property otherErrMsg2 As String = "戶名不符"
    Public Property otherErrMsg3 As String = "身分證號不符"
    Public Property otherErrMsg4 As String = "印鑑不符"
    Public Property otherErrMsg9 As String = "印鑑不符"
    Public Property noErrCode = "找不到錯誤代碼"
    Public Property GetReStatusP As String = "P:已發送授權書及授權扣款檔"
    Public Property GetReStatusR As String = "R:先收到回覆訊息但未收到授權書"
    Public Property GetReStatusY As String = "Y:先收到回覆訊息後收到授權書"
    Public Property GetReStatusM As String = "M:先收到授權書但未收到回覆訊息"
    Public Property GetReStatusS As String = "S:先收到授權書後收到回覆訊息"
    Public Property GetReStatusC As String = "C:已收到舊件轉檔回覆訊息"
    Public Property GetReStatusD As String = "D:已收到取消授權扣款回覆訊息"
    Public Property GetReStatusOther As String = "系統日期:{0}郵局核印成功!"
    Public Property getCancelStatus1 As String = "已提回終止扣款!"
    Public Property getCancelStatus2 As String = "郵局終止扣款!"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                NoACHCustId = Nothing
                ProcResult = Nothing
                NoData = Nothing
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
