Public Class SaveDataDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub

#Region "拆機退單同步退取回單"
    Friend Function GetRtnWip() As String
        Return String.Format("Select RowId,A.* From SO009 A Where CustId = {0}0 And MainSNo = {0}1 And SignDate is null And PRCode In (Select CodeNo From CD007 Where RefNo = 9) Order By SNo", Sign)
    End Function

    Friend Function GetDBWip() As String
        Return String.Format("Select * From SO009 Where SNO={0}0", Sign)
    End Function

    Friend Function GetOwner(ByVal Owner As String) As String
        If Owner <> String.Empty Then
            If Owner.Substring(Owner.Length - 1, 1) <> "." Then
                Owner = Owner & "."
            End If
        End If
        Return Owner
    End Function

    Friend Function GetSO313(ByVal Owner As String) As String
        Dim ReInstOwner As String = GetOwner(Owner)
        Return String.Format("Select * From {1}SO313 Where OCompCode={0}0 and OCustId={0}1 and OSNO={0}2 ", Sign, ReInstOwner)
    End Function

    Friend Function GetSO041() As String
        Return "Select * From SO041 "
    End Function

    Friend Function GetCD039() As String
        Return String.Format("Select * From CD039 Where CodeNo= {0}0 ", Sign)
    End Function
#End Region

#Region "更新客戶基本資料"
    Friend Function GetSO001() As String
        Return String.Format("Select * From SO001 Where Custid={0}0", Sign)
    End Function

    Friend Function GetSO137() As String
        Return String.Format("Select * From SO137 Where ID={0}0", Sign)
    End Function
#End Region

#Region "取得地址資料"
    Friend Function GetNewAddress() As String
        Return String.Format("Select A.*,Decode(B.ClctMethod,1,3,2,2,3,2,1) ClctMethod From SO014 A,SO017 B Where A.MduId = B.MduId(+) And A.AddrNo = {0}0", Sign)
    End Function
    Friend Function GetSO014() As String
        Return String.Format("Select * From SO014 Where AddrNo = {0}0 ", Sign)
    End Function
    Friend Function GetCD017() As String
        Return String.Format("Select * From CD017 Where CodeNo = {0}0", Sign)
    End Function
#End Region

#Region "拆復異動資料用"
    Friend Function GetCyclePeriodInvDef() As String
        Return String.Format("Select CMCode,CMName,PTCode,PTName,BankCode,BankName,AccountNo,InvSeqNo " &
                             "From SO003 Where 1=0", Sign)
    End Function

    Friend Function GetCustomerInvDef() As String
        Return String.Format("Select B.CMCode,B.CMName,B.InvoiceType,B.InvNo,B.InvTitle,B.InvAddress," &
                             "B.InvPurposeCode,B.InvPurposeName,B.InvoiceKind,B.Email,B.DenRecCode," &
                             "B.DenRecName,B.DenRecDate,A.CustNote,A.ChargeNote,A.InstAddrNo,A.InstAddress," &
                             "A.MailAddrNo,A.MailAddress From SO001 A,SO002 B Where A.CustId = B.CustId And " &
                             "A.CompCode = B.CompCode And B.CustId = {0}0 And B.CompCode = {0}1 And B.ServiceType = {0}2", Sign)
    End Function

    Friend Function GetStopAccountNo(TableName As String) As String
        Dim strSQL As String
        Dim Field As String
        If TableName = "SO106" Then
            Field = "AccountId"
        Else
            Field = "AccountNo"
        End If
        strSQL = String.Format("Update {1} X Set StopDate={0}2,StopFlag=1 Where CustId = {0}0" &
                               " And Exists (Select 1 From SO003 A Where A.AccountNo = X.{2} And A.CustId = X.CustId And " &
                               "A.CustId = {0}0 And A.ServiceType <> {0}1 And " &
                               "Exists (Select 1 From SO002 B Where A.CustId = B.CustId And " &
                               "B.CustStatusCode in (1,6)) And (A.StopFlag = 0 Or (A.StopType = 3 And A.StopFlag = 1) ))", Sign, TableName, Field)
        Return strSQL
    End Function

    Friend Function GetUseAccountCount() As Integer
        Return String.Format("Select Count(*) From SO003 A Where A.CustId = {0}0 And A.ServiceType <> {0}1 And " &
                             "Exists (Select 1 From SO002 B Where A.CustId = B.CustId And " &
                             "A.AccountNo = B.AccountNo And B.CustStatusCode in (1,6)) And (A.StopFlag = 0 Or (A.StopType = 3 And A.StopFlag = 1) )", Sign)
    End Function

    Friend Function GetOtherWip() As String
        Return String.Format("Select * From SO009 Where Custid={0}0 and ServiceType={0}1 and " &
                             " FinTime is Null and ReturnCode is Null and PrCode in (Select CodeNo From CD007 Where RefNo in (2,5)) ", Sign)
    End Function

    Private Shared Function GetOtherFailityUtil() As String
        Return "Select A.RowId,A.*,B.Description InitPlaceName,C.Description PgName,D.DVRSize " & _
                     "From SO004 A,CD056 B,CD029 C,CD102 D Where A.FaciCode = B.CodeNo(+) AND " & _
                     "A.PgNo=C.CodeNo(+) AND A.DVRAuthSizeCode=D.CODENO(+) "
    End Function

    Friend Function GetOtherFacility() As String
        Dim strSQL As String = String.Format("{1} And SNo = {0}0 ", Sign, GetOtherFailityUtil)
        Return strSQL
    End Function

#End Region

#Region "停復異動資料用"
    Friend Function GetStopPeriodCycle() As String
        Return String.Format("Update SO003 Set StopFlag = 0,CeaseDate = Null Where CustId = {0}0 And CompCode = {0}1 And ServiceType = {0}2", Sign)
    End Function
#End Region

#Region "更新客戶促銷資料檔用"
    Friend Function GetCustomerPromData() As String
        Return String.Format("Select * From SO098 Where CustId = {0}0 And ServiceType = {0}1 And PromCode = {0}2 And BulletinCode = {0}3", Sign)
    End Function
#End Region

#Region "更新設備資料-同區移機"
    Friend Function PrMoveToFacility() As String
        Return String.Format("CustId = {0}0 And CompCode = {0}1 And ServiceType = {0}2 and SeqNo = {0}3", Sign)
    End Function
#End Region

#Region "設備是否最後一台"
    Friend Function FaciCount(ByVal FaciSEQNO As String, ByVal FaciRefNo As String) As String
        Return String.Format("Select Count(0) From SO004 A,CD022 B Where A.CustID={0}0 and A.ServiceType={0}1 and SEQNO Not in ({1})" &
                             " and A.FaciCode=B.CodeNo and B.RefNo in ({2}) and A.PrDate Is Null", Sign, FaciSEQNO, FaciRefNo)
    End Function

    Friend Function FaciToUpSO002() As String
        Return String.Format("Update SO002 Set PRTIME={0}2 Where CustId={0}0 and ServiceType={0}1", Sign)
    End Function
#End Region

#Region "移機順產生其他服務移機單"
    Friend Function GetMoveFaciData(InterDependRefNo As String) As String
        Dim strSQL As String
        Dim RefNoStr As String = Nothing
        If String.IsNullOrEmpty(InterDependRefNo) = False Then
            RefNoStr = String.Format(" And FaciCode In (Select CodeNO From CD022 Where Nvl(RefNo,0) in ({0}))", InterDependRefNo)
        End If
        strSQL = String.Format("Select * From SO004 Where CustId = {0}0 And ServiceType = {0}1 And PRDate is null{1}", Sign, RefNoStr)
        Return strSQL
    End Function
#End Region

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
