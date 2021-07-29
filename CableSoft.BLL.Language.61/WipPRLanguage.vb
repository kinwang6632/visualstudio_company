Public Class WipPRLanguage

    'Private _LanguageUtility As New UtilityLanguage
    Public Shared Property SO004DKind = "更換"
    Public Shared Property OtherErrorCode As Integer = -199
    Public Shared Property OtherErrorMessage As String = "其它例外失敗"
#Region "PRLanguage"
    Public Shared Property NoPriv As String = "無權限!!"
    Public Shared Property ChkCustStatusError As String = "判斷客戶狀態異常!"
    Public Shared Property ChkCustIsCancel As String = "註銷戶無法產生派工單!"
    Public Shared Property WipPRDataNothing As String = "無任何拆機單資料!"
    Public Shared Property colClsTimeNothing As String = "無日結欄位可判斷!"
    Public Shared Property colClsTimeIsNotNull As String = "已日結不可修改資料!"
    Public Shared Property colClsTimeIsCancel As String = "已日結不可作廢資料!"

    Public Shared Property showStatusName As String = "停拆移{0}"
    Public Shared Function GetInvoiceName(ByVal InvoiceType As Integer) As String
        Dim RetString As String = String.Empty
        Select Case InvoiceType
            Case 1
                RetString = "裝機"
            Case 2
                RetString = "拆機"
            Case 3
                RetString = "維修"
            Case 4
                RetString = "暫存"
            Case Else
                RetString = ""
        End Select
        Return RetString
    End Function
    Public Shared Function GetStatusName(ByVal intEditMode As Integer) As String
        Dim RetString As String = String.Empty
        Select Case intEditMode
            Case 1
                RetString = "修改"
            Case 2
                RetString = "新增"
            Case 3
                RetString = "刪除"
            Case Else
                RetString = "瀏覽"
        End Select
        Return RetString
    End Function
    Public Shared Function SetWorkingName(ByVal intEditMode As Integer) As String
        Return String.Format(showStatusName, GetStatusName(intEditMode))
    End Function

    Public Shared Property setCustomerUcanAdd As String = "此客戶狀態為:{0},是否確定要新增?"
    Public Shared Function CustomerUcanAdd(ByVal StatusName As String) As String
        Return String.Format(setCustomerUcanAdd, StatusName)
    End Function

    Public Shared Property setCustomerIsWap As String = "該客戶已是[{0}]派工中狀態, 是否確認新增?"
    Public Shared Function CustomerIsWap(ByVal StatusName As String) As String
        Return String.Format(setCustomerIsWap, StatusName)
    End Function

    Public Shared Property setCustomerIsPR As String = "此客戶狀態為:{0},只能派拆設備,是否確定新增?"
    Public Shared Function CustomerIsPR(ByVal StatusName As String) As String
        Return String.Format(setCustomerIsPR, StatusName)
    End Function
#End Region

#Region "SaveDataLanguage"
    Public Shared Property FaciPRFromMove As String = "移機順拆"
    Public Shared Property WipPRandReturn As String = "拆機順退單"
    Public Shared Property WipPRRemoveAndReturn As String = "同區移機順退單"
    Public Shared Property PRtoCustNothing As String = "無對應到客戶基本資料"
    Public Shared Property UpdBuildingError As String = "異動大樓資料異常失敗!!"
    Public Shared Property WipPRandFin As String = "拆機復機順完工"
    Public Shared Property RelationWipIsFin As String = "關聯工單已完工,此工單不得退單!!"
    Public Shared Property RelationWipIsReturn As String = "關聯工單已退單,此工單不得完工!!"

    Public Shared Property setNotCodeRef As String = "{0}工單代碼無參考號{1}資料!!"
    Public Shared Property retrieve As String = "取回"
    Public Shared Property remove As String = "拆除"
    Public Shared Function NotCodeRef(WipType As Integer, WipRefNo As Integer) As String
        Return String.Format(setNotCodeRef, GetInvoiceName(WipType), WipRefNo)
    End Function
#End Region

#Region "ValidateLanguage"
    Public Shared Property DataUpdateError As String = "資料異常失敗"
    Public Shared Property CustIsCancel As String = "註銷戶無法產生派工單"
    Public Shared Property WipNotAddPR As String = "不可直接新增移拆單"
    Public Shared Property WipNotAddPR1 As String = "不可直接新增停(分)機單"
    Public Shared Property WipPRNotPayNow As String = "現付制不可直接派拆機單"
    Public Shared Property WipPRNotCancel As String = "如果沒有則不可做退單(Wip.ReturnCode 有值)"
    Public Shared Property FinTimeError As String = "完工時間必須大於前次裝機時間!"
    Public Shared Property WipPRHaveRemove As String = "移機設備中有派設備更換工單,請先將其退單或完工再派移機"
    Public Shared Property OtherWipFintTime As String = "關聯工單已完工,此工單不得退單!"

    Public Shared Property colNotNullPRCode As String = "必要欄位:停/拆機類別!"
    Public Shared Property colNotNullPRReason As String = "必要欄位:停/拆機原因!"
    Public Shared Property colNotNullResvTime As String = "必要欄位:預約時間!"
    Public Shared Property colNotNullReInstAddrNo As String = "必要欄位:新裝機地址"
    Public Shared Property colNotNullNewChargeAddrNo As String = "必要欄位:新收費地址"
    Public Shared Property colNotNullNewMailAddrNo As String = "必要欄位:新郵寄地址"
    Public Shared Property colNotNullNewTel1 As String = "必要欄位:新電話1"
    Public Shared Property colNotNullOldAddrNo As String = "新裝機地址不可等於原裝機地址"
    Public Shared Property colNotNullReInstDate As String = "必要欄位:復機預約時間"
    Public Shared Property WipRunStatus0 As String = "退單"
    Public Shared Property WipRunStatus1 As String = "完工"

    Public Shared Property CusidIsCancel As String = "註銷戶無法產生派工單"
    Public Shared Property OAddressReturnCannotAccept As String = "新地址移入單已退單,不可派工!!"
    Public Shared Property OAddressReturnCannotFinish As String = "新地址移入單已退單,移拆單只可做退單不可完工!!"
    Public Shared Property OAddressReturnCannotReturn As String = "新地址移入單已完工,移拆單只可做完工不可退單!!"
    Public Shared Property OAddressMustReOpenCommand As String = "如該設備已做過CM開關機動作! 現在退單請記得自行重做CM開關機,以利設備順利使用"
#End Region

End Class
