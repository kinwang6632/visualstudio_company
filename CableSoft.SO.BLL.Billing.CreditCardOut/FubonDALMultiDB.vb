Public Class FubonDALMultiDB
    Inherits FubonDAL
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function QueryCD068() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select ROW_NUMBER () OVER (ORDER BY BILLHEADFMT) AS CODENO,BILLHEADFMT AS DESCRIPTION from CD068 ORDER BY BILLHEADFMT "
            Case Else
                Return MyBase.QueryCD068
        End Select

    End Function
    Friend Overrides Function QuerySO1108A() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select * From (select Parameters from so1108a where programid = 'SO3272A3' and entryid={0}0 order by exectime desc)  A LIMIT 1", Sign)
            Case Else
                Return MyBase.QuerySO1108A
        End Select

    End Function
    Friend Overrides Function UpdUCCode(ByVal IsFubonIntegrate As Boolean) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                If IsFubonIntegrate Then
                    Return String.Format("UPDATE SO033  Set UCCode={0}0,UCName= {0}1 " &
                               ",UPDEN={0}2,UPDTIME={0}3,CMCode = {0}4,CMName = {0}5 " &
                           " Where CTID  In (Select A.CTID::text From " &
                                           "SO001 B ," &
                                           "SO002 D," &
                                           "SO033 A," &
                                           "CD013 " &
                                           " Where 1=1 " &
                                           " And A.MediabillNo ={0}6 " &
                                           "  And A.AccountNO={0}7" &
                                            " And A.CustId=B.CustId And A.CancelFlag = 0  " &
                                           "  And A.UCCode Is Not Null  " &
                                           " And  A.CustID = D.CustID   " &
                                            " And A.SERVICETYPE = D.SERVICETYPE " &
                                           " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) Not IN (3,7)  " &
                                          " And (substr(A.AccountNo, 1, 6) in (Select CodeNo From CD143 Where Length(CodeNo) = 6 ) " &
                                          "  Or substr(A.AccountNo,1,7) In (Select CodeNo From CD143 Where Length(CodeNo)= 7 )) " &
                                           " And  NVL(CD013.PAYOK,0) = 0 )", Sign)
                Else
                    Return String.Format("UPDATE SO033  Set UCCode={0}0,UCName= {0}1 " &
                               ",UPDEN={0}2,UPDTIME={0}3 " &
                           " Where CTID  In (Select A.CTID::text From " &
                                           "SO001 B ," &
                                           "SO002 D," &
                                           "SO033 A," &
                                           "CD013 " &
                                           " Where 1=1 " &
                                           " And A.MediabillNo ={0}4 " &
                                           "  And A.AccountNO={0}5" &
                                            " And A.CustId=B.CustId And A.CancelFlag = 0  " &
                                           "  And A.UCCode Is Not Null  " &
                                           " And  A.CustID = D.CustID   " &
                                            " And A.SERVICETYPE = D.SERVICETYPE " &
                                           " And A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) Not IN (3,7)  " &
                                           " And  NVL(CD013.PAYOK,0) = 0 )", Sign)
                End If
            Case Else
                Return MyBase.UpdUCCode(IsFubonIntegrate)
        End Select



    End Function
    Friend Overrides Function GetInvoiceNo2(ByVal StrTableName As String) As String
        Dim strSeq As String
        strSeq = "S_" & StrTableName
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select '" & Date.Now.ToString("yyMM") &
                     "' || Ltrim(To_Char(sf_getsequenceno('" & strSeq & "'), '0999999')) "
            Case Else
                Return MyBase.GetInvoiceNo2(StrTableName)
        End Select


    End Function
    Friend Overrides Function GetViewName() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT trim(To_Char(sf_getsequenceno('S_TMPRPT_ViewName'), '0999999')) "
            Case Else
                Return MyBase.GetViewName
        End Select

    End Function
    Friend Overrides Function QuerySO033Data(ByVal aSO033Where As String, ByVal isZero As Integer,
                                   ByVal isCrossCustCombine As Boolean, ByVal CompCode As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aSQL As String = "Select A.MediaBillNO, sum(A.ShouldAmt)  ShouldAmt," &
                 " A.AccountNO,B.CUSTNAME,B.CUSTID," &
                 " SUM(A.TCBbudget) TCBbudget , MAX(A.PTCODE)  PtCode," &
                 " MAX(A.ServiceType) ServiceType,  " &
                 GetMinPayKindWhere() & " , " & GetMinRealStopDateWhere() &
                 " FROM SO033 A,SO001 B ,SO002 D ,CD013 " &
                 "Where A.CustID  = D.CUSTID And " &
                 "A.CustId=B.CustId And A.CancelFlag = 0 And " &
                 "A.UCCode Is Not Null  And " &
                 "A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) Not In(3,7) And " &
                 "NVL(CD013.PAYOK,0)= 0 " &
                 " And " & aSO033Where &
                 " And A.SERVICETYPE = D.SERVICETYPE   " &
                 " Group by A.MediaBillNO ,  A.AccountNO  , B.CUSTNAME,B.CUSTID   " &
                 GetHavingOutZeroWhere(isZero) &
                 " ORDER BY A.MediaBillNO "

                aSQL = "Select A.MediaBillNO BILLNO ,ShouldAmt," &
                    " A.AccountNO,A.CUSTNAME,A.CUSTID,Min(A.RealStopDate) RealStopDate," &
                    " Min(A.PayKind) PayKind , " &
                    " A.TCBbudget , A.PTCODE,A.ServiceType,C.CVC2, " & GetStopYmWhere() &
                    " FROM (" & aSQL & ") A, SO106 C " &
                    " WHERE A.ACCOUNTNO=C.ACCOUNTID  " &
                    " And C.STOPFLAG<>1 And C.SnactionDate Is Not NULL " &
                    " GROUP BY A.MediaBillNO ,ShouldAmt,A.AccountNO," &
                      " A.CUSTNAME,A.CUSTID,A.TCBbudget,A.PTCODE,A.ServiceType,C.CVC2 " &
                    " ORDER BY A.MediaBillNO"
                If isCrossCustCombine Then
                    aSQL = "Select A.*," &
                                " (Case 1 When 1 Then Nvl(SO001.PROID,Null)   Else Null End ) PROID, " &
                                " (Case 1 When 1 Then " &
                                               " ( Case  When PROID Is Null Then A.CustId Else Nvl((Select MainCustId From SO018 Where SO001.PROID = SO018.MduId " &
                                                " And now() >= CStartDate And now()<= CDueDate  And COMPCODE = " & CompCode & " Limit 1  ),-1)  End ) " &
                                "  Else A.CustId  End ) MainCustId " &
                                " From (" & aSQL & ") A,SO001 " &
                        " Where A.CustId=SO001.CustId"


                End If

                Return aSQL
            Case Else
                Return MyBase.QuerySO033Data(aSO033Where, isZero, isCrossCustCombine, CompCode)

        End Select

    End Function
    Friend Overrides Function UpdLogData() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("UPDATE SO1108A SET EXECSTATUS={0}0 ,EXECMESSAGE={0}1 " &
                             " ,FINISHTIME = now(),DOWNLOADFILENAME = {0}2 " &
                        " WHERE SEQNO = {0}3", Sign)
            Case Else
                Return MyBase.UpdLogData
        End Select

    End Function
    Friend Overrides Function QueryViewSQL(ByVal viewName As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aSQL As String
                aSQL = " select A.*,(select (case " &
                                 "                when PROID is null then 1 Else " &
                                                         " ( select  count(1) from " & viewName &
                                                         " Where A.MainCustId = Custid ) " &
                                                 " End)  from dual) CustIdExistFlag  from " & viewName & " A"
                'aSQL = "Select Decode(CustIdExistFlag,1,Custid,MainCustId)  CustId,Nvl(BillNo,'X') BillNo,AccountNo,Max(PayKind) PayKind,Min(RealStopDate) RealStopDate," & _
                '                " Sum(ShouldAmt) ShouldAmt,PROID,CustIdExistFlag " & _
                '        " From ( " & aSQL & " ) A Group By Decode(CustIdExistFlag,1,Custid,MainCustId),BillNo,AccountNo,PROID,CustIdExistFlag "
                aSQL = "Select (CASE CustIdExistFlag WHEN 1 THEN Custid ELSE MainCustId END)  CustId,Nvl(BillNo,'X') BillNo,AccountNo,Max(PayKind) PayKind,Min(RealStopDate) RealStopDate," &
                        " Sum(ShouldAmt) ShouldAmt,PROID,CustIdExistFlag " &
                " From ( " & aSQL & " ) A Group By Decode(CustIdExistFlag,1,Custid,MainCustId),BillNo,AccountNo,PROID,CustIdExistFlag "
                aSQL = "Select distinct A.*, " &
                        " (select TCBbudget from " & viewName &
                                    " where a.billno=billno and  A.AccountNo = AccountNo Limit 1) TCBbudget ," &
                        " (select PTCode from " & viewName &
                                    " where a.billno=billno and  A.AccountNo = AccountNo Limit 1) PTCode ," &
                        " (select ServiceType from " & viewName &
                                    " where a.billno=billno and  A.AccountNo = AccountNo Limit 1) ServiceType ," &
                         " (select CVC2 from " & viewName &
                                    " where a.billno=billno and  A.AccountNo = AccountNo Limit 1) CVC2 ," &
                        " (select CardExpDate from " & viewName &
                                    " where a.billno=billno  and  A.AccountNo = AccountNo Limit 1) CardExpDate " &
                        " From (" & aSQL & ") A"
                Return aSQL
            Case Else
                Return MyBase.QueryViewSQL(viewName)
        End Select

    End Function
End Class
