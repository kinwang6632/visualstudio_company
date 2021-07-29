Public Class MaintainLanguage
    Implements IDisposable
    Public Property CancelCust As String = "註銷戶無法產生派工單!"
    Public Property NotInstallCust As String = "非裝機中或復機中客戶無法產生派工單 !"
    Public Property NotUseServiceType As String = "沒有可用的服務別!"
    Public Property CloseCanNotEdit As String = "已日結不可修改資料!"
    Public Property CloseCanNotCancel As String = "已日結不可作廢資料!"
    Public Property GetWipCalculateDataErr As String = "執行GetWipCalculateData錯誤!--"
    Public Property AcceptDateFmtErr As String = "受理時間日期格式不正確!"
    Public Property MustCustId As String = "需有客戶編號或設備流水號!"
    Public Property ChangeWipErr As String = "ChangeWip 失敗!"
    Public Property ChangeChargeErr As String = "ChangeCharge 失敗!"
    Public Property ChangeCommandDataErr As String = "ChangeCommandData 失敗!"
    Public Property SendCmdNotRetValue As String = "SendCmd Error: SendCmd有錯,但無任何回傳值!"
    Public Property ChangeFacilityErr As String = "ChangeFacility 失敗!"
    Public Property ChangePRFacilityErr As String = "ChangePRFacility 失敗!"
    Public Property ChangeChangeFacilityErr As String = "ChangeChangeFacility 失敗!"
    Public Property ChangeResvDetailErr As String = "ChangeResvDetail 失敗!"
    Public Property ChangeResvLogErr As String = "ChangeResvLog 失敗!"
    Public Property ChangeResvTempPoint As String = "ChangeResvTempPoint 失敗!"
    Public Property DelResvPoint As String = "DelResvPoint 失敗!"
    Public Property IsNullField As String = "本欄位是否為空值!"
    Public Property DateFmtErr As String = "日期不合法!"
    Public Property DateExceed As String = "此日期超過今天日期!"
    Public Property DateExceedSave As String = "此日期已超過系統設定的安全期限!"
    Public Property DateHasClose As String = "之前已做過日結,不可使用之前日期入帳"
    Public Property NoMaintainData As String = "無維修工單資料!"
    Public Property MustChangeFaci As String = "需指定變更設備!"
    Public Property MustFixNum1 As String = "請設定故障代號一!"
    Public Property MustFixNum2 As String = "請設定故障代號二!"
    Public Property FixingArea As String = "已發生區域故障，是否要派維修工單?!"
    Public Property AddClientInfo As String = "維修新增"
    Public Property EditClientInfo As String = "維修修改"
    Public Property Mantaining As String = " 該客戶狀態已是拆機中且該客戶已是維修中狀態, 是否確認新增?"
    Public Property Demolishing As String = "該客戶狀態已是拆機中, 是否確認新增?"
    Public Property NoInstalling = "非裝機中或復機中客戶無法產生派工單！"
    Public Property OnMaintain = "該客戶為裝機中,復機中或維修中狀態, 是否確認新增?"
    Public Property NotNormal = "該客戶非正常客戶, 是否確認新增?"


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                NotNormal = Nothing
                OnMaintain = Nothing
                NoInstalling = Nothing
                Demolishing = Nothing
                Mantaining = Nothing
                CancelCust = Nothing
                NotInstallCust = Nothing
                NotUseServiceType = Nothing
                CloseCanNotEdit = Nothing
                CloseCanNotCancel = Nothing
                GetWipCalculateDataErr = Nothing
                AcceptDateFmtErr = Nothing
                MustCustId = Nothing
                ChangeWipErr = Nothing
                ChangeChargeErr = Nothing
                ChangeCommandDataErr = Nothing
                SendCmdNotRetValue = Nothing
                ChangeFacilityErr = Nothing
                ChangePRFacilityErr = Nothing
                ChangeChangeFacilityErr = Nothing
                ChangeResvDetailErr = Nothing
                ChangeResvLogErr = Nothing
                ChangeResvTempPoint = Nothing
                DelResvPoint = Nothing
                IsNullField = Nothing
                DateFmtErr = Nothing
                DateExceed = Nothing
                DateExceedSave = Nothing
                DateHasClose = Nothing
                NoMaintainData = Nothing
                MustChangeFaci = Nothing
                MustFixNum1 = Nothing
                MustFixNum2 = Nothing
                FixingArea = Nothing
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
