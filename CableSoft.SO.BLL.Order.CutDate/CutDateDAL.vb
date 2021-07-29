Public Class CutDateDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable
    Private _Disposed As Boolean

    Public Sub New()
    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    ''' <summary>
    ''' 取得查詢可用下收切齊資料SQL[SO003,SO004,SO033] ; SQL參數:CustId
    ''' </summary>
    ''' <returns>回傳查詢可用下收切齊資料SQL</returns>
    ''' <remarks>SQL參數:CustId</remarks>
    Friend Function QueryCanCutDate() As String
        'Return String.Format("Select * From" &
        '                     " (" &
        '                     "  Select 1 as Type,A.ServiceType,A.CitemCode,A.CitemName,A.Period" &
        '                     "     ,A.Amount,A.StartDate,A.StopDate,A.ClctDate,(Select SO137.DECLARANTNAME From SO137.ID = B.ID) DeclarantName " &
        '                     "     ,A.BankCode,A.BankName,A.AccountNo,A.InvSeqNo,A.CMCode,A.CMName" &
        '                     "     ,A.PTCode,A.PTName,A.CustAllot,A.FaciSeqNo,A.Period NextPeriod" &
        '                     "     ,A.Amount NextAmt" &
        '                     "     From SO003 A,SO004 B" &
        '                     "     Where A.CustId=B.CustId(+)" &
        '                     "     And A.FaciSeqNo=B.SeqNo(+)" &
        '                     "     And A.CustId={0}0" &
        '                     "     And A.ClctDate>={0}1" &
        '                     "     And A.StopFlag<>1" &
        '                     "     And A.CitemCode In (Select CodeNo From CD019 Where Sign='+')" &
        '                     "  Union All" &
        '                     "  Select 2 as Type,A.ServiceType,A.CitemCode,A.CitemName,A.RealPeriod Period" &
        '                     "     ,A.ShouldAmt Amount,A.RealStartDate StartDate,A.RealStopDate StopDate,Null" &
        '                     "     ,(Select SO137.DECLARANTNAME From SO137.ID = B.ID) DeclarantName,A.BankCode,A.BankName,A.AccountNo,A.InvSeqNo,A.CMCode" &
        '                     "     ,A.CMName,A.PTCode,A.PTName,0 CustAllot,A.FaciSeqNo" &
        '                     "     ,Decode(Nvl(C.Period,0),0" &
        '                     "     ,Decode(Nvl(A.NextPeriod,0),0,A.RealPeriod,A.NextPeriod),C.Period) NextPeriod" &
        '                     "     ,0 NextAmt" &
        '                     "     From SO033 A,SO004 B,SO003 C" &
        '                     "     Where A.CustId=B.CustId(+)" &
        '                     "     And A.FaciSeqNo=B.SeqNo(+)" &
        '                     "     And A.CustId=C.CustId(+)" &
        '                     "     And A.FaciSeqNo=C.FaciSeqNo(+)" &
        '                     "     And A.CitemCode=C.CitemCode(+)" &
        '                     "     And A.CustId={0}2" &
        '                     "     And A.RealStopDate+1>={0}3" &
        '                     "     And A.CancelFlag<>1" &
        '                     "     And A.CitemCode In (Select CodeNo From CD019 Where Sign='+') A" &
        '                     "     And UCCode Is Not Null" &
        '                     "  Order By ServiceType,CitemCode,StopDate Desc,Type)",
        '                     Sign)
        Return String.Format("Select * From" &
                            " (" &
                            "  Select 1 as Type,A.ServiceType,A.CitemCode,A.CitemName,A.Period" &
                            "     ,A.Amount,A.StartDate,A.StopDate,A.ClctDate,(Select SO137.DECLARANTNAME From SO137 Where SO137.ID = B.ID) DeclarantName " &
                            "     ,A.BankCode,A.BankName,A.AccountNo,A.InvSeqNo,A.CMCode,A.CMName" &
                            "     ,A.PTCode,A.PTName,A.CustAllot,A.FaciSeqNo,A.Period NextPeriod" &
                            "     ,A.Amount NextAmt" &
                            "     From SO003 A left join SO004 B on  A.CustId=B.CustId And A.FaciSeqNo=B.SeqNo " &
                            "     Where 1=1" &
                            "     And A.CustId={0}0" &
                            "     And A.ClctDate>={0}1" &
                            "     And A.StopFlag<>1" &
                            "     And A.CitemCode In (Select CodeNo From CD019 Where Sign='+')" &
                            "  Union All" &
                            "  Select 2 as Type,A.ServiceType,A.CitemCode,A.CitemName,A.RealPeriod Period" &
                            "     ,A.ShouldAmt Amount,A.RealStartDate StartDate,A.RealStopDate StopDate,Null" &
                            "     ,(Select SO137.DECLARANTNAME From SO137.ID = B.ID) DeclarantName,A.BankCode,A.BankName,A.AccountNo,A.InvSeqNo,A.CMCode" &
                            "     ,A.CMName,A.PTCode,A.PTName,0 CustAllot,A.FaciSeqNo" &
                            "     ,Decode(Nvl(C.Period,0),0" &
                            "     ,Decode(Nvl(A.NextPeriod,0),0,A.RealPeriod,A.NextPeriod),C.Period) NextPeriod" &
                            "     ,0 NextAmt" &
                            "     From SO033 A left join SO004 B on A.CustId=B.CustId And And A.FaciSeqNo=B.SeqNo " &
                            "      left join SO003 C on A.CustId=C.CustId  And A.FaciSeqNo=C.FaciSeqNo And A.CitemCode=C.CitemCode " &
                            "     Where 1=1 " &
                            "     And A.CustId={0}2" &
                            "     And A.RealStopDate+1>={0}3" &
                            "     And A.CancelFlag<>1" &
                            "     And A.CitemCode In (Select CodeNo From CD019 Where Sign='+')  " &
                            "     And UCCode Is Not Null" &
                            "  Order By ServiceType,CitemCode,StopDate Desc,Type) AS A ",
                            Sign)
    End Function


#Region "IDisposable Support"

    ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not _Disposed Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        _Disposed = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
