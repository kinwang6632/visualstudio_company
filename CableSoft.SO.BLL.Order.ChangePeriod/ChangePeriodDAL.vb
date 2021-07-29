Public Class ChangePeriodDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable
    Public Sub New()
    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    ''' <summary>
    ''' 取得查詢可選繳別SQL[CD078A1,CD078,CD019] ; SQL參數:BPCode,CitemCode
    ''' </summary>
    ''' <returns>回傳查詢可選繳別SQL</returns>
    ''' <remarks>SQL參數:BPCode,CitemCode</remarks>
    Friend Function QueryBPPeriod() As String
        ' ServiceType ? StopFlag ?
        Return String.Format("Select A.*,B.Description BPName,C.Description CitemName" &
                             " From CD078A1 A,CD078 B,CD019 C" &
                             " Where A.BPCode=B.CodeNo" &
                             " And A.CitemCode=C.CodeNo" &
                             " And A.BPCode={0}0" &
                             " And A.CitemCode={0}1" &
                             " Order By A.BPCode", Sign)
    End Function

    ''' <summary>
    ''' 取得查詢繳別細項SQL[CD078A12] ; SQL參數:StepNo
    ''' </summary>
    ''' <returns>回傳查詢繳別細項SOL</returns>
    ''' <remarks>SQL參數:StepNo</remarks>
    Friend Function QueryBPPeriodLevel() As String
        Dim aSQL As String = Nothing
        Dim aRet As String = Nothing
        For i As Int32 = 1 To 12
            aSQL = String.Format("SELECT STEPNO,{1} LEVELITEM," & _
                            " MON{1} MON, PERIOD{1} PERIOD, RATETYPE{1} RATETYPE, " & _
                            " DISCOUNTAMT{1} DISCOUNTAMT,MONTHAMT{1} MONTHAMT, " & _
                            " DAYAMT{1} DAYAMT, PUNISH{1} PUNISH, COMMENT{1} DISCOUNTNOTE," & _
                            " DISCOUNTRATE{1} DISCOUNTRATE " & _
                            " FROM CD078A1 WHERE STEPNO = {0}{2}", Sign, i, i - 1)
            If i = 1 Then
                aRet = aSQL
            Else
                aRet = String.Format("{0} UNION ALL {1} ", aRet, aSQL)
            End If
        Next
        Return aRet
    End Function
    Friend Function QueryCD078A1() As String        
        Dim aSQL As String = Nothing
        aSQL = String.Format("Select Nvl(MonthAmt,0) MonthAmt From CD078A " & _
                            " Where LONGPAYflag = 1 " & _
                            " And BpCode = {0}0 And CitemCode = {0}1 ", Sign)
        Return aSQL
    End Function
    ''' <summary>
    ''' 取得查詢組合產品代碼SQL[CD078A1,CD078A12] ; SQL參數:Period,CitemCode,BPCode,StepNo
    ''' </summary>
    ''' <returns>回傳取得查詢組合產品代碼SQL</returns>
    ''' <remarks>SQL參數:Period,CitemCode,BPCode,StepNo</remarks>
    Friend Function QueryBPCode() As String
        'Return String.Format("Select B.MonthAmt,B.Period,B.Mon,A.StepNo,A.LinkKey,B.DiscountNote" &
        '                     " From CD078A1 A,CD078A12 B" &
        '                     " Where A.StepNo=B.StepNo" &
        '                     " And A.BPCode={0}0" &
        '                     " And A.CitemCode={0}1" &
        '                     " And A.Period={0}2" &
        '                     " Order By Decode(A.Period,{0}2,0,1),B.LevelItem", Sign)

        Dim aSQL As String = Nothing
        '        Select StepNo,<Loop Item> LevelItem,Mon<Loop Item> Mon, Period<Loop Item> Period, RateType<Loop Item> RateType, DiscountAmt<Loop Item> DiscountAmt,MonthAmt<LoopItem> MonthAmt, DayAmt<LoopItem> DayAmt, Punish<LoopItem> Punish, Comment<LoopItem> DiscountNote, DiscountRate <LoopItem> DiscountRate From CD078A1 Where A.CitemCode = “Charge.CitemCode” And A.BPCode = “Change.BPCode” And (A.Period = “Period” Or A.StepNo = “Change.StepNo”)
        Dim aRet As New Text.StringBuilder("")
        '2014/10/22 Jacky 增加抓取的順序, 先抓有期數, 無則取原StepNo
        Dim Table As String = String.Format("Select * From (Select RANK() OVER (PARTITION BY A.CitemCode ORDER BY Type,StepNo) as RankX, A.*  From (Select 0 Type,A.* From CD078A1 A Where A.CitemCode = {0}0 And A.BPCode = {0}1 And A.Period = {0}2 " &
                                " Union All " &
                                " Select 1 Type,A.* From CD078A1 A Where A.CitemCode = {0}0 And A.BPCode = {0}1 And A.StepNo = {0}3) A )  A Where RankX = 1", Sign)

        For i As Int32 = 1 To 12
            aSQL = String.Format("Select A.StepNo,{1} LevelItem,A.Mon{1} Mon, A.Period{1} Period, A.RateType{1} RateType, " & _
                               " A.DiscountAmt{1} DiscountAmt,A.MonthAmt{1} MonthAmt, A.DayAmt{1} DayAmt, A.Punish{1} Punish, " & _
                               " A.Comment{1} DiscountNote, A.DiscountRate{1} DiscountRate,A.LinkKey " & _
                               " From ({2}) A" & _
                               " Where A.CitemCode = {0}0 And A.BPCode ={0}1 And (A.Period ={0}2 Or A.StepNo ={0}3)",
                                Sign, i, Table)
            If i = 1 Then
                aRet.Append(aSQL)
            Else
                aRet.Append(" UNION ALL " & aSQL)
            End If
        Next


        Return aRet.ToString()
    End Function

#Region "IDisposable Support"
    Private _Disposed As Boolean


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
