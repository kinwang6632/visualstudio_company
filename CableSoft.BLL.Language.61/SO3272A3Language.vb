Public Class SO3272A3Language
    Public Property ClientInfoString As String = "富邦銀行出帳"
    Public Property PrcResult As String = "已完成資料筆數共{0}筆," & vbCrLf & vbCrLf & _
                                "問題筆數共{1}筆," & vbCrLf & vbCrLf & _
                                "共花費:{2}秒"
    Public Property NoFoundComCustId As String = "單據編號：{0} 套房找不到統收戶客戶編號，請設定統收戶客戶編號"
    Public Property WatchDateWrong As String = "收視截止日不正確 --> 收視截止日大於扣款處理日：單號　{0}  客編：{1} " & _
                                   " 客戶姓名　{2}  收視截止日：{3}  繳付類別：{4}"
    Public Property CreditCardEmpty As String = "信用卡卡號空白 : 單號 {0} 客戶姓名 {1} "
    Public Property CreditCardDateEmpty = "信用卡日期不正確 : 單號 {0} 客戶姓名 {1}"
    Public Property CreditCardDue = "信用卡過期 : 單號 {0}  客戶姓名 {1}   信用卡號 {2}  到期日 {3}"
    Public Property ZeroAmt As String = "負項或是金額等於零 : 單號 {0}  客戶姓名 {1}    金額 {2}"

End Class
