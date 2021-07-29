
Public Class CounterPayLanguage
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Shared Property ProcessKind1 As String = "櫃檯繳款新增"
    Public Shared Property ProcessKind2 As String = "櫃檯繳款修改"
    Public Shared Property ProcessKind3 As String = "櫃檯繳款取消登錄"
    Public Shared Property ProcessKind4 As String = "櫃檯繳款結轉"
    Public Shared Property ProcessKind5 As String = "櫃檯信用卡刷卡繳費"

    Public Shared Property ChkBillOK As String = "無此單據編號或此單據已登錄過,請核對!"
    Public Shared Property BillOver As String = "此單據已登錄過!請重新輸入!"
    Public Shared Property CustNotOk As String = "此客戶狀態已非正常收視戶!"
    Public Shared Property NoParaData As String = "未傳入參數資料!"
    Public Shared Property NoChargeData As String = "未傳入收費資料!"
    Public Shared Property AddChargeTmpErr As String = "新增櫃臺收費登錄暫存資料失敗!"
    Public Shared Property UpdRealChargeErr As String = "更新正式應收資料失敗!"
    Public Shared Property ChargePayOK As String = "此筆資料已入實收或作廢,不可取消!"
    Public Shared Property ChargeCancelOK As String = "該筆收費資料已被作廢,不異動該收費資料!"
    Public Shared Property ChargePaymentOK As String = "此筆資料已線上刷卡,不可取消!"
    Public Shared Property ChargeTmpNoData As String = "櫃檯繳款暫存檔中查無資料!"
    Public Shared Property ChkNoData As String = "應收資料檔查無資料"
    Public Shared Property ChargeClose As String = "已入帳或作廢"
    Public Shared Property ChargeEdit As String = "該筆收費資料已被修改"
    Public Shared Property NoDelCharge As String = "此筆資料已入實收或作廢，不可取消"

    Public Shared Property sSuccessMsg As String = "共更新{1}成功資料{0}筆!{1}異常資料{2}筆!"
    Public Shared Function SuccessMsg(SCount As Integer, ECount As Integer) As String
        Return String.Format(sSuccessMsg, SCount, ControlChars.CrLf, ECount)
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                ChkBillOK = Nothing
                BillOver = Nothing
                CustNotOk = Nothing
                NoParaData = Nothing
                NoChargeData = Nothing
                AddChargeTmpErr = Nothing
                UpdRealChargeErr = Nothing
                ChargePayOK = Nothing
                ChargeCancelOK = Nothing
                ChargePaymentOK = Nothing
                ChargeTmpNoData = Nothing
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
