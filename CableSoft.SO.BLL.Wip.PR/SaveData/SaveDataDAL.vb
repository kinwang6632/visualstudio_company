Public Class SaveDataDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function GetStartNewProd() As String
        Return "Select nvl(startnewprod,0) as startnewprod from so041"
    End Function
    Friend Overridable Function QuerySO015(ByVal strWhere As String) As String
        Return String.Format("Select RowId ,SO015.* From SO015 Where {0} Order by  InDate Desc", strWhere)
    End Function
#Region "拆機退單同步退取回單"
    Friend Overridable Function GetRtnWip(ByVal SEQNO As String, ByVal refNo As Integer) As String
        'so009.sno=so004d.sno and so004d.seqno=so004.seqno and so004.facicode=cd022.codeno 
        '#8790 第一點 DVR取回單不應該連動退單
        Dim aSql As String
        Dim ref10Where As String = String.Empty
        If refNo = 10 Then
            ref10Where = " And SO004D.SeqNo in (Select Seqno From SO004 where SO004D.SeqNo = SO004.SeqNo " & _
                                    " And FaciCode in (Select CodeNo From CD022 Where SO004.FaciCode = CD022.CodeNo And CD022.RefNo =  9))"

            aSql = String.Format("select Rowid,A.* from so009 A where A.sno in ( " & _
                             " select SO004d.sno from so004d join so004 on SO004D.SEQNO = SO004.SEQNO  join cd022 " & _
                             "  on so004.facicode = CD022.CODENO and cd022.refno = 9 " & _
                             "  where so004d.seqno in (" & SEQNO & ") " & _
                             " and so004d.sno in (select sno from so009 where mainsno =( select  distinct mainsno from so009 where custid = {0}0 and sno = {0}1) " & _
                             " AND PRCode In (Select CodeNo  FROM CD007 Where RefNo = 9) ) " & _
                             " and (kind ={0}2 or kind = {0}3)) ", Sign)

        Else
            aSql = String.Format("Select RowId,A.* From SO009 A Where A.CustId = {0}0 " & _
                             " And A.MainSNo = {0}1 " & _
                             " And A.SNO IN (Select SNO From SO004D Where SO004D.SNO = A.SNO AND SEQNO In (" & SEQNO & ") " & ref10Where & _
                             " AND (KIND = {0}2 or KIND = {0}3))" & _
                             " And A.SignDate is null And A.PRCode In (Select CodeNo From CD007 Where RefNo = 9) Order By SNo", Sign)
        End If
        
        Return aSql
    End Function
    Friend Overridable Function GetRtnWip(ByVal refno As Integer) As String
        '#8790 第三點針對DVR拆機單退單，不會連動退DVR取回單
        Return String.Format("Select RowId,A.* From SO009 A Where A.CustId = {0}0 " & _
                          " AND A.SNO IN (SELECT SNO FROM SO004D WHERE SO004D.SNO = A.SNO " & _
                          " AND SO004D.SEQNO IN (SELECT SEQNO FROM SO004 WHERE SO004D.SEQNO = SO004.SEQNO " & _
                          " AND SO004.FACICODE IN (SELECT CODENO FROM CD022 WHERE SO004.FACICODE = CD022.CODENO AND CD022.REFNO = 9)))" & _
                           " And A.MainSNo = {0}1 " & _
                           " And A.SignDate is null And A.PRCode In (Select CodeNo From CD007 Where RefNo = 9) Order By SNo", Sign)
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

    Friend Function GetCD039(CompanyTableName As String) As String
        Return String.Format("Select * From {1} Where CodeNo= {0}0 ", Sign, CompanyTableName)
    End Function

    Friend Function GetServiceType(ByVal CanUseServiceType As String) As String
        Dim strWhere As String = String.Empty
        If Not String.IsNullOrEmpty(CanUseServiceType) Then strWhere = String.Format("Where CodeNo in ('{0}')", CanUseServiceType.Replace(",", "','"))
        Return String.Format("Select CodeNo,Description,DependService From CD046 {0} Order by CodeNo", strWhere)
    End Function

    Friend Function GetSO009MainSNO(ByVal IsFinTime As Boolean) As String
        Dim strWhere As String = "And FinTime is not Null"
        'IsFinTime邏輯比要特別，True:表完工，需要檢核是否有退單資料。 False:表退單，需要檢核是否有完工資料
        If IsFinTime Then strWhere = "And ReturnCode is not Null"
        Return String.Format("Select * From SO009 Where CustId = {0}0 And ServiceType = {0}1 And MainSNo = {0}2 And SNO <> {0}3 {1}", Sign, strWhere)
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
        'Return String.Format("Select A.*,Decode(B.ClctMethod,1,3,2,2,3,2,1) ClctMethod From SO014 A,SO017 B Where A.MduId = B.MduId(+) And A.AddrNo = {0}0", Sign)
        Return String.Format("Select A.*, (CASE B.CLCTMETHOD " &
                                          " WHEN 1 THEN 3 " &
                                         " WHEN 2 THEN 2 " &
                                         " WHEN 3 THEN 2 " &
                                         " ELSE 1 END) AS ClctMethod " &
                                    " FROM SO014 A LEFT JOIN SO017 B ON A.MduId = B.MduId " &
                                    " WHERE 1=1 AND A.AddrNo = {0}0", Sign)
    End Function
    Friend Function GetSO014() As String
        Return String.Format("Select * From SO014 Where AddrNo = {0}0 ", Sign)
    End Function
    Friend Function GetAddressData() As String
        Dim strSQL As String
        strSQL = String.Format("Select * From SO014 Where AddrNo = {0}0 And CompCode = {0}1", Sign)
        Return strSQL
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

    Friend Overridable Function GetOtherFailityUtil() As String
        'Return "Select A.RowId,A.*,B.Description InitPlaceName,C.Description PgName,D.DVRSize " &
        '             "From SO004 A,CD056 B,CD029 C,CD102 D Where A.FaciCode = B.CodeNo(+) AND " &
        '             "A.PgNo=C.CodeNo(+) AND A.DVRAuthSizeCode=D.CODENO(+) "

        Return " Select  A.rowid,A.*,B.Description InitPlaceName,C.Description PgName,D.DVRSize " &
                      "From SO004 A left Join CD056 B on A.FaciCode = B.CodeNo   " &
                      " Left Join CD029 C On A.PgNo=C.CodeNo, CD102 D Where " &
                      " 1=1 And A.DVRAuthSizeCode=D.CODENO "
    End Function

    Friend Overridable Function GetOtherFacility() As String

        Dim strSQL As String = " Select  A.rowid,A.*,B.Description InitPlaceName,C.Description PgName,D.DVRSize " &
                      "From SO004 A left Join CD056 B on A.FaciCode = B.CodeNo   " &
                      " Left Join CD029 C On A.PgNo=C.CodeNo, CD102 D Where " &
                      " 1=1 And A.DVRAuthSizeCode=D.CODENO "
        strSQL = String.Format("{1} And SNo = {0}0 ", Sign, strSQL)
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
        Return String.Format("Select * From SO004 A,CD022 B Where A.CustID={0}0 and A.ServiceType={0}1  " & _
                                " and SEQNO Not in (" & FaciSEQNO & ")" &
                                " and A.InstDate is not null " & _
                             " and A.FaciCode=B.CodeNo and B.RefNo in (" & FaciRefNo & ") and A.PrDate Is Null", Sign)
    End Function
    Friend Function isRefNo78(ByVal FaciSeqNo As String) As String
        Dim result As String = "Select nvl(count(*),0) cnt From SO004 A,CD022 B " & _
                " Where A.SEQNO in (" & FaciSeqNo & ") And A.FaciCode = B.CodeNo And B.Refno in (7,8) "
        Return result
    End Function
    Friend Function QryReInstAcrossFlag() As String
        Return String.Format("Select Nvl(ReInstAcrossFlag,0) From CD007 Where CodeNo= {0}0", Sign)
    End Function
    Friend Function FaciToUpSO002() As String
        Return String.Format("Update SO002 Set PRTIME={0}0,PR2SNO={0}1,PRCODE={0}2,PRName={0}3  " & _
                             " Where CustId={0}4 and ServiceType={0}5", Sign)

    End Function
#End Region

#Region "移機順產生其他服務移機單"
    Friend Function GetMovePRCode(FaciRefNo As String) As String
        Return String.Format("Select CodeNo,Description,RefNo From CD007 Where StopFlag=0 And RefNo=3 And Instr(','||MoveRefno||',',',{0},')>0", FaciRefNo)
    End Function
    Friend Overridable Function GetMoveFaciData(InterDependRefNo As String, strCalcFaciRefNo As String) As String
        Dim strSQL As String
        '#7922 因需求修改，比照舊版的 csAlterWip4.clsAlterWip3.AlterSO00x2 內的取設備的語法調整。

        'strSQL = String.Format("Select A.RowId,A.*,B.RefNo FaciRefNo From SO004 A,CD022 B Where A.FaciCode = B.CodeNo  And A.CustId = {0}0 And A.ServiceType = {0}1" &
        '                       " And A.PRDate is null And A.GetDate is null And (A.PRSNo is null Or Exists (Select 1 From SO004 X Where A.SeqNo = X.ReSeqNo And A.CustId = X.CustId)) And A.InstDate is not null" &
        '                       " And B.RefNo in (" & strCalcFaciRefNo & ") And B.RefNo in (" & InterDependRefNo & ") Order By Decode(B.RefNo,9,0,7,2,8,3,4),A.SeqNo", Sign)
        strSQL = String.Format("Select A.RowId,A.*,B.RefNo FaciRefNo From SO004 A,CD022 B Where A.FaciCode = B.CodeNo  And A.CustId = {0}0 And A.ServiceType = {0}1" &
                               " And A.PRDate is null And A.GetDate is null And (A.PRSNo is null Or Exists (Select 1 From SO004 X Where A.SeqNo = X.ReSeqNo And A.CustId = X.CustId)) And A.InstDate is not null" &
                               " And B.RefNo in (" & strCalcFaciRefNo & ") And B.RefNo in (" & InterDependRefNo & ") Order By (Case B.RefNo when 9 then 0 when 7 then 2 when 8 then 3 else 4 end ),A.SeqNo", Sign)
        Return strSQL
    End Function
#End Region

#Region "拆機、拆分機工單 判斷是否還有其他正常設備存在"
    Friend Function CheckFaci_ChangePrCode(ByVal RefNO As String, ByVal ServiceType As String) As String
        Return String.Format("Select * From CD007 Where ServiceType='{2}' and REFNO={1} and NVL(STOPFLAG,0)=0 and ReturnPR=1 and ReInstAcrossFlag=0 Order by CodeNO", Sign, RefNO, ServiceType)
    End Function
    Friend Function CheckFaci_ChangePrCode(ByVal RefNO As String, ByVal ServiceType As String, ReInstAcrossFlag As String, ByVal prCode As String) As String
        '#8713 add condition of functype equal to original prcode and take it for the main code by kin
        Dim result As String = String.Format("Select CodeNo,Description, 0 flag From CD007 Where ServiceType='{0}' and " & _
                             " REFNO={1} and NVL(STOPFLAG,0)=0 and ReturnPR=1 and " & _
                             " ReInstAcrossFlag={2} And  functype in (select functype from cd007 where codeno = {3}) ", ServiceType, RefNO, ReInstAcrossFlag, prCode)
        result = String.Format("{0} Union All Select CodeNo,Description, 1 flag  From CD007 Where ServiceType='{1}' and " & _
                             " REFNO={2} and NVL(STOPFLAG,0)=0 and ReturnPR=1 and " & _
                             " ReInstAcrossFlag={3} ", result, ServiceType, RefNO, ReInstAcrossFlag)
        result = String.Format("select * from ({0}) A Order by flag,codeno", result)
        Return result
            'Return String.Format("Select * From CD007 Where ServiceType='{0}' and " & _
            '                     " REFNO={1} and NVL(STOPFLAG,0)=0 and ReturnPR=1 and " & _
            '                     " ReInstAcrossFlag={2} Order by CodeNO", ServiceType, RefNO, ReInstAcrossFlag)
    End Function
    Friend Function CheckFaci_ChangePrCode(ByVal RefNO As String, ByVal ServiceType As String, ReInstAcrossFlag As String) As String
        Return CheckFaci_ChangePrCode(RefNO, ServiceType, ReInstAcrossFlag, "-9")
        'Return String.Format("Select * From CD007 Where ServiceType='{0}' and " & _
        '                     " REFNO={1} and NVL(STOPFLAG,0)=0 and ReturnPR=1 and " & _
        '                     " ReInstAcrossFlag={2} Order by CodeNO", ServiceType, RefNO, ReInstAcrossFlag)
    End Function
    Friend Function UpdateDB_ChangePrCode()
        Return String.Format("Update SO009 SET PRCODE={0}3,PRNAME={0}4 Where Custid={0}0 and SNO={0}1 and ServiceType={0}2", Sign)
    End Function
    Friend Function updMainSnoSelf() As String
        Return String.Format("update so009 set mainsno = sno where sno = {0}0", Sign)
    End Function
    Friend Function QuerySO138() As String
        Return String.Format(" Select ChargeAddrNo ,MailAddrNo From So138 Where InvSeqNo In  " &
            " (Select B.InvSeqNo From SO002A  A,SO002AD B Where A.CustId = B.CustId " &
            " And A.AccountNo = B.AccountNo And A.CustId = {0}0)", Sign)
    End Function
    Friend Function updSO138ChargeAddrNo(ByVal aStartNewProd As Integer) As String

        Dim result As String = Nothing
        If aStartNewProd = 0 Then
            result = String.Format("Update SO138 Set ChargeAddrNo = {0}0 ,ChargeAddress= {0}1 " &
                " Where InvSeqNo In (Select B.InvSeqNo From SO002A A,SO002AD B Where  " &
                " A.AccountNo = B.AccountNo And A.CustId = B.CustId And A.CustId =  {0}2 ) " &
                " And ChargeAddrNo = {0}3 ", Sign)
        Else
            result = String.Format("Update SO138 Set ChargeAddrNo = {0}0 ,ChargeAddress= {0}1 " &
                " Where InvSeqNo In (Select A.InvSeqNo From SO003C A,SO002 B Where  " &
                " A.ServiceType = B.ServiceType And  A.InvSeqNo is not null And A.CustId = B.CustId And A.CustId =  {0}2 ) " &
                " And ChargeAddrNo = {0}3 ", Sign)
        End If

        Return result
    End Function
    Friend Function updSO138MailAddrNo(ByVal aStartNewProd As Integer) As String

        Dim result As String = Nothing
        If aStartNewProd = 0 Then
            result = String.Format("Update  SO138 Set MailAddrNo = {0}0,MailAddress= {0}1 " &
            " Where InvSeqNo In (Select B.InvSeqNo From SO002A A,SO002AD B Where  " &
            " A.AccountNo = B.AccountNo And A.CustId = B.CustId And A.CustId = {0}2)  " &
            " And MailAddrNo = {0}3 ", Sign)
        Else
            result = String.Format("Update SO138 Set MailAddrNo = {0}0 ,MailAddress= {0}1 " &
               " Where InvSeqNo In (Select A.InvSeqNo From SO003C A,SO002 B Where  " &
               " A.ServiceType = B.ServiceType And A.InvSeqNo is not null And A.CustId = B.CustId And A.CustId =  {0}2 ) " &
               " And MailAddrNo = {0}3 ", Sign)
        End If

        Return result
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
