Imports CableSoft.BLL.Utility
Public Class ManualNoDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryCompCode() As String
        Return String.Format("Select A.CodeNo,A.Description  " & _
                            " From CD039 A,SO026 B  " & _
                            " Where Instr(','||B.CompStr||',',','||A.CodeNo||',')>0 " & _
                           " And UserId = {0}0 Order By CodeNO", Sign)

    End Function
    Friend Function QueryPaperNum() As String
        Return String.Format("Select A.*,(select (case nvl(to_char(max(billno)),'0') when '0' then 0 else 1 end ) From so127 where A.seq=seq and compcode = A.compcode) isExist  " &
                             " From SO127 A Where " &
                      "PaperNum >= {0}0 And PaperNum <={0}1", Sign)
    End Function
    Friend Function QueryExistData() As String
        Return String.Format("Select Count(1) From SO127 Where " & _
                             "Seq = {0}0 And PaperNum={0}1 And CompCode = {0}2", Sign)
    End Function
    Friend Function chkHadUse() As String
        Return String.Format("Select Count(*) From SO127 Where ( BillNo is Not Null Or Nvl(Status,0) = 0 ) " & _
                             " And To_number(Replace(PaperNum,{0}0,'')) >= To_number({0}1) And To_number(Replace(PaperNum,{0}2,'')) <=to_number( {0}3) " & _
                             " And Seq = {0}4 And CompCode = {0}5 ", Sign)
    End Function
    Friend Function chkDual() As String
        Return String.Format("select count(1) from so126 where ( to_number(beginnum) between to_number({0}0) " & _
                                        " and to_number({0}1) " & _
                                        " Or  to_number(endnum) between to_number({0}2) and to_number({0}3) ) And Prefix = {0}4 ", Sign)
    End Function
    Friend Function UpdReUseSO126(ByVal isSameBegin As Boolean) As String
        If Not isSameBegin Then
            Return String.Format("Update SO126 Set EndNum = {0}0,UPDTIME={0}1,OPERATOR = {0}2,TotalPaperCount = {0}3 " & _
                                 " Where Seq = {0}4 And CompCode = {0}5", Sign)
        Else
            Return String.Format("Update SO126 Set UPDTIME={0}0,OPERATOR = {0}1,EmpNO = {0}2,EmpName={0}3, " & _
                                 "GetPaperDate = {0}4 Where Seq = {0}5 And CompCode = {0}6", Sign)
        End If

    End Function
    Friend Function QuerySingleSO127() As String
        Return String.Format("Select * From SO127 Where PaperNum = {0}0 And CompCode = {0}1", Sign)
    End Function
    Friend Function CanDelete() As String
        Return String.Format("Select Count(1) From SO127 Where CustId is Not Null " & _
                             " And Seq = {0}0 And PaperNum >= {0}1 And PaperNum <={0}2 " & _
                             " And CompCode = {0}3", Sign)
    End Function
    Friend Function DeleteSO127() As String
        Return String.Format("delete SO127 Where 1=1 " & _
                            " And Seq = {0}0 And PaperNum >= {0}1 And PaperNum <={0}2 " & _
                            " And CompCode = {0}3", Sign)
    End Function
    Friend Function DeleteSO126() As String
        Return String.Format("Delete SO126 Where Seq={0}0 And Prefix={0}1 " & _
                             " And to_number(BeginNum) >= to_number({0}2) " & _
                             " And to_number(EndNum) <= to_number({0}3) " & _
                             " And CompCode = {0}4", Sign)
    End Function
    Friend Function ClearSO127() As String
        Return String.Format("Update SO127 Set custid = null,custname = null,CustTEL= null, " & _
                             "BillNo = Null,RealDate = Null,OPERATOR = {0}0,UPDTIME ={0}1 " & _
                             " Where PaperNum = {0}2 And CompCode = {0}3", Sign)
    End Function
    Friend Function UpdSO127ManualNo() As String
        Return String.Format("Update SO127 Set custid = {0}0,custname = {0}1,CustTEL= {0}2, " & _
                             "BillNo = {0}3,RealDate = {0}4,OPERATOR = {0}5,UPDTIME ={0}6 " & _
                             " Where PaperNum = {0}7 And CompCode = {0}8", Sign)
    End Function
    Friend Function UpdBillManual() As String
        Return String.Format("Update SO034 Set ManualNo ={0}0 Where " & _
                             " BillNo = {0}1 And Item = {0}2 And CompCode = {0}3", Sign)
    End Function
    Friend Function abandonPaper() As String
        Return String.Format("Update SO127 Set Status = 0,OPERATOR = {0}0,UPDTIME={0}1 Where " & _
                         " PaperNum >= {0}2 And PaperNum <= {0}3 " & _
                        " And Seq = {0}4 And CompCode = {0}5", Sign)
    End Function
    Friend Function QueryBillData() As String
        Return String.Format("select a.CustId,b.CustName, b.Tel1, a.BillNo,a.ManualNo, " & _
                             " a.Item,a.CitemName,a.ShouldAmt,a.RealDate, a.RealAmt, a.ClctName,a.CMName,a.STName " & _
                             " from SO034 a, SO001 b where a.BillNo={0}0 and a.CompCode={0}1 " & _
                             " and a.CompCode=b.CompCode and a.CustID=b.CustID Order by A.CustId,A.BillNo,A.Item", Sign)
    End Function
    Friend Function UpdSO127() As String
        Return String.Format("Update SO127 Set OPERATOR = {0}0,UPDTIME = {0}1 Where Seq = {0}2", Sign)
    End Function
    Friend Function UpdSO126() As String
        Return String.Format("Update SO126 Set OPERATOR = {0}0,UPDTIME = {0}1," & _
                                            " RETURNDATE = {0}2,CLEARDATE= {0}3,NOTE={0}4 " & _
                                            " Where Seq = {0}5 And CompCode = {0}6 ", Sign)
    End Function
    Friend Overridable Function QuerySeqVal() As String
        Return "Select S_SO126_SEQ.NextVal From Dual"
    End Function
    Friend Function InsSO126() As String
        Return String.Format("Insert into SO126 (CompCode,Seq,EmpNO,EmpName,GetPaperDate,BeginNum, " & _
                                            "EndNum,TotalPaperCount,OPERATOR,UPDTIME,Prefix,RETURNDATE,CLEARDATE, " & _
                                            "NOTE) Values ( {0}0,{0}1,{0}2,{0}3,{0}4,{0}5," & _
                                                "{0}6,{0}7,{0}8,{0}9,{0}10,{0}11,{0}12,{0}13)", Sign)
    End Function
    Friend Function InsSO127() As String
        Return String.Format("Insert into SO127 (CompCode,Seq,PaperNum,EmpNO,EmpName,GetPaperDate," & _
                                    "Status,OPERATOR,UPDTIME ) Values ({0}0,{0}1,{0}2,{0}3,{0}4,{0}5,{0}6,{0}7,{0}8)", Sign)
    End Function
    Friend Function UpdReUseSO127() As String
        Return String.Format("Update SO127 Set SEQ={0}0,EmpNO={0}1,EmpName={0}2,GetPaperDate={0}3," & _
                                    "OPERATOR = {0}4,UPDTIME= {0}5 Where SEQ = {0}6 " & _
                                    " And To_number(Replace(PaperNum,{0}7,'')) >= To_number({0}8) " & _
                                    " And To_number(Replace(PaperNum,{0}9,'')) <= To_number({0}10) ", Sign)
    End Function
    Friend Overridable Function QueryData(ByVal ds As DataSet) As String
        Dim tbWhere As DataTable = ds.Tables(0)
        Dim strWhere As String = " 1 = 1 "
        If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("GETPAPERDATE1")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("GETPAPERDATE1")) Then
            strWhere = strWhere & String.Format(" And GetPaperDate >= To_Date('{0}','yyyy/mm/dd') ", tbWhere.Rows(0).Item("GETPAPERDATE1"))
        End If
        If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("GETPAPERDATE2")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("GETPAPERDATE2")) Then
            strWhere = strWhere & String.Format(" And GetPaperDate <= To_Date('{0}','yyyy/mm/dd') ", tbWhere.Rows(0).Item("GETPAPERDATE2"))
        End If
        If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("EMPNO")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("EMPNO")) Then
            strWhere = strWhere & String.Format(" And EMPNO IN ({0})", tbWhere.Rows(0).Item("EMPNO"))
        End If
        If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("PREFIX")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("PREFIX")) Then
            strWhere = strWhere & String.Format(" And PREFIX IN ({0})", tbWhere.Rows(0).Item("PREFIX"))
        End If
        If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("SEQNO")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("SEQNO")) Then
            strWhere = strWhere & String.Format("And {0} between to_number(BEGINNUM) and  to_number(ENDNUM)", tbWhere.Rows(0).Item("SEQNO"))
        End If
        Return "Select A.*,rowid From SO126 A Where " & strWhere
    End Function
    Friend Function QueryEmployee() As String
        Return "Select EmpNo CodeNo,EmpName Description From CM003  Where Nvl(StopFlag,0) = 0 Order By EmpNo"
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
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
