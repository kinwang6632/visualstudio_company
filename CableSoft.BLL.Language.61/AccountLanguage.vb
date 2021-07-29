Public Class AccountLanguage
    Implements IDisposable

    Public Property NoPKField As String = "SO106 NO PKField"
    Public Property MustField As String = "缺少必要欄位"
    Public Property AchSNDouble As String = "[申請書單號] 資料重複 !"
    Public Property CMDataError As String = "收費資料檢核有問題"
    Public Property Authorizeing As String = "ACH授權中不允許停用 !"
    Public Property MustAccountTable As String = "無Account Table"
    Public Property MustProductTable As String = "無ChooseProduct Table"
    Public Property MustStopField As String = "無停用欄位"
    Public Property MustStopDateField As String = "無停用日期欄位"
    Public Property MustMasterId As String = "無唯一序號"
    Public Property NoSO001Data As String = "找不到SO001資料"
    Public Property NoSO002Data As String = "找不到SO002資料"
    Public Property InsertSO106Error As String = "SO106 Insert Error"
    Public Property UpdateSO106Error As String = "SO106 Update Error"
    Public Property InsertSO106LogError As String = "新增SO106LOG失敗"
    Public Property UpdateSO106AError As String = "更新SO106A失敗"
    Public Property InsertSO106AError As String = "新增SO106A失敗"
    Public Property NoDataUpdate As String = "無任何資料可異動"
    Public Property NoProposer As String = "無申請人欄位 !"
    Public Property NoProposerData As String = "無申請人資料 !"
    Public Property NoPropDate As String = "無銀行核印日期欄位 !"
    Public Property NoPropDateData As String = "無銀行核印日期資料 !"
    Public Property NoCMCodeField As String = "無收費方式欄位 !"
    Public Property NoCMCodeData As String = "無收費方式資料 !"
    Public Property NoPTCodeField As String = "無付款種類欄位 !"
    Public Property NoPtCodeData As String = "無付款種類資料 !"
    Public Property MustACHT As String = "無ACH授權交易別資料 !"
    Public Property AddClientInfo As String = "帳號資料新增"
    Public Property EditClientInfo As String = "帳號資料修改"
    Public Property noAccountNameField As String = "無帳號所有人欄位！"
    Public Property noAccountNameData As String = "無帳號所有人資料！"
    Public Property noAccountNameIDField As String = "無帳號所有人身份證號欄位！"
    Public Property noAccountNameIDData As String = "無帳號所有人身份證號資料！"
    Public Property AccountIdLimit As String = "帳號長度必須是 = {0}"
    Public Property noBankCodeField As String = "無銀行代碼欄位！"
    Public Property noBankCodeData As String = "無銀行代碼資料！"
    Public Property noCardCodeField As String = "無信用卡別欄位！"
    Public Property noCardCodeData As String = "無信用卡資料！"
    Public Property noStopYMField As String = "無信用卡有效期限欄位！"
    Public Property noStopYMData As String = "無信用卡有效期限資料！"
    Public Property noFoundCardCode As String = "找不到信用卡代碼"
    Public Property noFoundCardType As String = "信用卡種類不存在!!"
    Public Property VisaHeader As String = "VISA卡 開頭數字必須為 4"
    Public Property VisaLenLimit As String = "VISA卡 長度必須為 16"
    Public Property MasterHeader As String = "MASTER卡 開頭數字必須為5"
    Public Property MasterLenLimit As String = "MASTER卡 長度必須為16"
    Public Property JCBHeader As String = "JCB卡 開頭數字必須為 3"
    Public Property JCBLenLimit As String = "JCB卡 長度必須為16"
    Public Property AmericaLimit As String = "美國運通卡 長度必須為15"
    Public Property BigLimit As String = "大來卡 長度必須為14"
    Public Property ptCash As String = "現金"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                NoPKField = Nothing
                MustField = Nothing
                AchSNDouble = Nothing
                CMDataError = Nothing
                Authorizeing = Nothing
                MustAccountTable = Nothing
                MustProductTable = Nothing
                MustStopField = Nothing
                MustStopDateField = Nothing
                MustMasterId = Nothing
                NoSO001Data = Nothing
                NoSO002Data = Nothing
                InsertSO106Error = Nothing
                UpdateSO106Error = Nothing
                InsertSO106LogError = Nothing
                UpdateSO106AError = Nothing
                InsertSO106AError = Nothing
                NoDataUpdate = Nothing
                NoProposer = Nothing
                NoProposerData = Nothing
                NoPropDate = Nothing
                NoPropDateData = Nothing
                NoCMCodeField = Nothing
                NoCMCodeData = Nothing
                NoPTCodeField = Nothing
                NoPtCodeData = Nothing
                MustACHT = Nothing
                noAccountNameField = Nothing
                noAccountNameData = Nothing
                noAccountNameIDField = Nothing
                noAccountNameIDData = Nothing
                AccountIdLimit = Nothing
                noBankCodeField = Nothing
                noBankCodeData = Nothing
                noCardCodeField = Nothing
                noCardCodeData = Nothing
                noStopYMField = Nothing
                noStopYMData = Nothing
                noFoundCardCode = Nothing
                noFoundCardType = Nothing
                VisaHeader = Nothing
                VisaLenLimit = Nothing
                MasterHeader = Nothing
                MasterLenLimit = Nothing
                JCBHeader = Nothing
                JCBLenLimit = Nothing
                AmericaLimit = Nothing
                BigLimit = Nothing
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
