Public Class Invoice
    Implements IDisposable

    Public Property INVFORMAT1 As String = "雲端"
    Public Property INVFORMAT2 As String = "手二"
    Public Property INVFORMAT3 As String = "手三"
    Public Property UPLOADFLAGYes As String = "是"
    Public Property UPLOADFLAGNo As String = "否"
    Public Property QueryWhere1 As String = "客編"
    Public Property QueryWhere2 As String = "姓名"
    Public Property QueryWhere3 As String = "聯絡人"
    Public Property QueryWhere4 As String = "統編"
    Public Property QueryWhere5 As String = "電話"
    Public Property InvoiceKind1 As String = "電子計算機發票"
    Public Property InvoiceKind2 As String = "雲端發票"
    Public Property InvoiceKind3 As String = "全部"
    Public Property NoInvModify As String = "無發票資料, 不可修改。"
    Public Property ObsoleteYes As String = "此筆發票已作廢"
    Public Property NOMODIFY As String = "此筆發票需做過授權才可修改!"
    Public Property INVDue As String = "此筆發票已過期(該月份已鎖帳或發票號碼不在有效發票本內),不可修改。"
    Public Property NOINV099 As String = "找不到發票字軌檔"
    Public Property INV099Waring As String = "本次預計開立發票張數為 {0} 張, 剩餘可用發票張數為 {1} 張, 發票張數不足, 請重新選擇發票本。"
    Public Property SystemIDNull As String = "檢核發票檢查碼為空白 無法存檔"
    Public Property ExcessDetail As String = "發票明細筆數已超出換開筆數設定, 設定的筆數限制為 {0} 筆。"
    Public Property TotalCreate As String = "此次共開立 :[{0}] 共 {1} 張發票"
    Public Property GridOrder1 As String = "收費日期"
    Public Property GridOrder2 As String = "郵遞區號"
    Public Property GridOrder3 As String = "收費日期 + 郵遞區號"
    Public Property GridOrder4 As String = "收費日期 + 郵遞區號 + 客編"
    Public Property HowtoCreate1 As String = "預開"
    Public Property HowtoCreate2 As String = "後開"
    Public Property TAXTYPE1 As String = "應稅"
    Public Property TAXTYPE2 As String = "零稅率"
    Public Property TAXTYPE3 As String = "免稅"
    Public Property CannotBatchCreate As String = "預估開立發票張數共 {0} 張, " & Environment.NewLine &
                                                                         "所選擇發票本可用發票張數共 {1} 張，" & Environment.NewLine &
                                                                          "發票張數不足無法開立！"
    Public Property CreateResult As String = "開立結果：(共費時  {0}   秒)"
    Public Property noBillClose As String = "未日結"
    Public Property BillClose As String = "己日結"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下方的 Finalize()。
            ' TODO: 將大型欄位設為 null。
        End If
        disposedValue = True
    End Sub

    ' TODO: 只有當上方的 Dispose(disposing As Boolean) 具有要釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Visual Basic 加入這個程式碼的目的，在於能正確地實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
        Dispose(True)
        ' TODO: 覆寫上列 Finalize() 時，取消下行的註解狀態。
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
