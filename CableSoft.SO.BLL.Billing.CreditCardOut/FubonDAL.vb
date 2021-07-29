Imports CableSoft.BLL.Utility
Public Class FubonDAL
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
    Friend Function QueryPayKindName() As String
        Return "SELECT Description FROM CD112 WHERE CODENO=1 "
    End Function
    Friend Function QueryCustIDAndName() As String
        Return String.Format("SELECT  SO001.CUSTNAME , SO033.CUSTID FROM SO001,SO033 " & _
                        " WHERE SO033.CUSTID = SO001.CUSTID " & _
                        " And SO033.CompCode = SO001.CompCode And SO033.MediaBillNo={0}0", Sign)
    End Function
    Friend Overridable Function QueryCD068() As String
        Return "Select ROWNUM CODENO,BILLHEADFMT DESCRIPTION from CD068 ORDER BY ROWNUM"
    End Function
    Friend Function QueryCD018() As String
        Return String.Format("Select * from cd018 Where UPPER(PrgName) = 'CREDITCARDFUBON'  AND  COMPCODE = {0}0 order by codeno", Sign)
    End Function
    Friend Function QueryCD031() As String
        Return "select CodeNo,Description from cd031 where Nvl(stopflag,0) <>1 And Nvl(RefNo,0) <> 5 order by codeno"
    End Function
    Friend Function QueryCD031REFNO5() As String
        Return "select CodeNo,Description from cd031 where Nvl(stopflag,0) <>1 And Nvl(RefNo,0) = 5 order by codeno"
    End Function

    Friend Function QueryCD001() As String
        Return "select CodeNo,Description from cd001 where nvl(stopflag,0) <>1 order by codeno"
    End Function
    Friend Function QueryCD002() As String
        Return "select CodeNo,Description from cd002 where nvl(stopflag,0) <>1 order by codeno"
    End Function
    Friend Function QueryCD035() As String
        Return "select CodeNo,Description from cd035  order by codeno"
    End Function
    Friend Function QueryCD004() As String
        Return "select CodeNo,Description from cd004 where nvl(stopflag,0) <> 1  order by codeno"
    End Function
    Friend Function QueryCM003() As String
        Return "select empno codeno,empname description from cm003 where nvl(stopflag,0) <> 1 order by empno"
    End Function
    Friend Function QueryCD013() As String
        Return "select codeno,description from cd013 Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PAYOK,0)=0 and Nvl(stopflag,0)<> 1"
    End Function
    Friend Function QueryCD019() As String
        Return String.Format("select codeno,description from cd019 where " &
                        " codeno in (select citemcode from cd068a where billheadfmt={0}0) order by codeno", Sign)

    End Function
    Friend Function BuildCondition(ByVal dsCondition As DataSet, ByVal CompCode As String,
                                   ByRef strSO033Where As String) As String

        Dim result As String = " A.CompCode = " & CompCode
        With dsCondition.Tables("Condition").Rows(0)
            If Not DBNull.Value.Equals(.Item("ShouldDate1")) AndAlso Not String.IsNullOrEmpty(.Item("ShouldDate1")) Then
                result = result & " AND A.ShouldDate >= To_Date('" & .Item("ShouldDate1").ToString.Replace("/", "").Replace(" ", "") & "','YYYYMMDD') "
            End If
            If Not DBNull.Value.Equals(.Item("ShouldDate2")) AndAlso Not String.IsNullOrEmpty(.Item("ShouldDate2")) Then
                result = result & " AND A.ShouldDate < To_Date('" &
                            .Item("ShouldDate2").ToString.Replace("/", "").Replace(":", "").Replace(" ", "") & "','YYYYMMDD') + INTERVAL '1' DAY "
            End If
            If Not DBNull.Value.Equals(.Item("CreateTime1")) AndAlso Not String.IsNullOrEmpty(.Item("CreateTime1")) Then
                result = result & " AND A.CreateTime >= To_Date('" &
                    .Item("CreateTime1").ToString.Replace("/", "").Replace(":", "").Replace(" ", "") & "','YYYYMMDDHH24MISS') "
            End If
            If Not DBNull.Value.Equals(.Item("CreateTime2")) AndAlso Not String.IsNullOrEmpty(.Item("CreateTime2")) Then
                result = result & " AND  A.CreateTime < To_Date('" &
                    .Item("CreateTime2").ToString.Replace("/", "").Replace(":", "").Replace(" ", "") & "','YYYYMMDDHH24MISS') + INTERVAL '1' DAY"
            End If
            If Not DBNull.Value.Equals(.Item("CMCode")) AndAlso Not String.IsNullOrEmpty(.Item("CMCode")) Then
                result = String.Format(" {0} AND A.CMCode IN ({1}) ", result, .Item("CMCode"))
            Else
                ' #8537 By Kin 2019/12/17 
                If Integer.Parse(.Item("IsFubonIntegrate").ToString) = 0 Then
                    result = String.Format(" {0} AND A.CMCode  In (select codeno from cd031 where nvl(Stopflag,0) =0 And nvl(refno,0) <> 5 )", result)
                End If
            End If
            If Not DBNull.Value.Equals(.Item("AreaCode")) AndAlso Not String.IsNullOrEmpty(.Item("AreaCode")) Then
                result = String.Format(" {0} AND A.AreaCode IN ({1}) ", result, .Item("AreaCode"))
            End If
            If Not DBNull.Value.Equals(.Item("ServCode")) AndAlso Not String.IsNullOrEmpty(.Item("ServCode")) Then
                result = String.Format(" {0} AND A.ServCode IN ({1}) ", result, .Item("ServCode"))
            End If
            If Not DBNull.Value.Equals(.Item("ClctEn")) AndAlso Not String.IsNullOrEmpty(.Item("ClctEn")) Then
                result = String.Format(" {0} AND A.ClctEn IN ({1}) ", result, .Item("ClctEn"))
            End If
            If Not DBNull.Value.Equals(.Item("OldClctEn")) AndAlso Not String.IsNullOrEmpty(.Item("OldClctEn")) Then
                result = String.Format(" {0} AND A.OldClctEn IN ({1}) ", result, .Item("OldClctEn"))
            End If
            If Not DBNull.Value.Equals(.Item("PayKind")) AndAlso Not String.IsNullOrEmpty(.Item("PayKind")) Then
                result = String.Format(" {0} AND A.PayKind IN ({1}) ", result, .Item("PayKind"))
            End If
            If Not DBNull.Value.Equals(.Item("CustId")) AndAlso Not String.IsNullOrEmpty(.Item("CustId")) Then
                result = String.Format(" {0} AND A.CustId IN ({1}) ", result, .Item("CustId"))
            End If
            If Not DBNull.Value.Equals(.Item("BILLNOTYPE")) AndAlso Not String.IsNullOrEmpty(.Item("BILLNOTYPE")) Then
                result = String.Format(" {0} AND SubStr(A.BillNo,7,1) IN ({1}) ", result, .Item("BILLNOTYPE"))
            End If
            If Not DBNull.Value.Equals(.Item("CreateEn")) AndAlso Not String.IsNullOrEmpty(.Item("CreateEn")) Then
                result = String.Format(" {0} AND A.CreateEn IN ({1}) ", result, .Item("CreateEn"))
            End If
            If Not DBNull.Value.Equals(.Item("MDUIDE")) AndAlso Not String.IsNullOrEmpty(.Item("MDUIDE")) Then
                If Integer.Parse(.Item("ISOTHER").ToString) = 1 Then
                    result = String.Format("{0} And (A.MduId IN ( {1} )  Or A.MduId Is Null) ", result, .Item("MDUIDE"))
                Else
                    result = String.Format("{0} And A.MduId IN ({1})", result, .Item("MDUIDE"))
                End If
            Else
                If Not DBNull.Value.Equals(.Item("MDUIDN")) AndAlso Not String.IsNullOrEmpty(.Item("MDUIDN")) Then
                    If Integer.Parse(.Item("ISOTHER").ToString) = 1 Then
                        result = String.Format(" {0} And (Not A.MduId {1} Or A.MduId Is Null)", result, .Item("MDUIDN"))
                    Else
                        result = String.Format(" {0} And Not A.MduId IN ({1})", result, .Item("MDUIDN"))
                    End If
                End If
            End If
            If Integer.Parse(.Item("ISOTHER").ToString) = 0 Then
                If DBNull.Value.Equals(.Item("MDUIDN")) AndAlso DBNull.Value.Equals(.Item("MDUIDE")) Then
                    result = String.Format(" {0} And A.MduId Is Null", result)
                End If
            End If
            'Cancel the conditions #8537 By Kin 2019/12/17 
            'If Integer.Parse(.Item("IsFubonIntegrate").ToString) = 1 Then
            '    result = String.Format("{0} And (substr(A.AccountNo,1,6) in (Select CodeNo From CD143 Where Length(CodeNo)= 6 ) " &
            '                           " Or substr(A.AccountNo,1,7) In (Select CodeNo From CD143 Where Length(CodeNo)= 7 ))", result)
            'End If
            If Not DBNull.Value.Equals(.Item("UCCODE")) AndAlso Not String.IsNullOrEmpty(.Item("UCCODE")) Then
                result = String.Format(" {0} And A.UCCODE In ({1}) ", result, .Item("UCCODE"))
            End If
            If Not DBNull.Value.Equals(.Item("BillHeadFmt")) AndAlso Not String.IsNullOrEmpty(.Item("BillHeadFmt")) Then
                result = String.Format(" {0} And A.CitemCode In (Select CitemCode From CD068A " &
                                                                    " Where BillHeadFmt = '{1}') ", result, .Item("BillHeadFmt"))


            End If
            If Not DBNull.Value.Equals(.Item("BANKCODE")) AndAlso Not String.IsNullOrEmpty(.Item("BANKCODE")) Then
                result = String.Format(" {0} AND A.BANKCODE IN ({1}) ", result, .Item("BANKCODE"))
            End If
            If Not DBNull.Value.Equals(.Item("ExcI")) AndAlso Not String.IsNullOrEmpty(.Item("ExcI")) Then
                If Integer.Parse(.Item("ExcI").ToString) = 1 Then
                    result = String.Format(" {0} AND substr(A.billno,7,1) || A.BillMark <> 'I'  ", result)
                End If
            End If
            strSO033Where = result
            If Not DBNull.Value.Equals(.Item("CustStatusCode")) AndAlso Not String.IsNullOrEmpty(.Item("CustStatusCode")) Then
                result = String.Format(" {0} AND D.CustStatusCode IN ({1}) ", result, .Item("CustStatusCode"))
            End If
            If Not DBNull.Value.Equals(.Item("ClassCode1")) AndAlso Not String.IsNullOrEmpty(.Item("ClassCode1")) Then
                result = String.Format(" {0} AND B.ClassCode1 IN ({1}) ", result, .Item("ClassCode1"))
            End If
            If Not DBNull.Value.Equals(.Item("AMduId")) AndAlso Not String.IsNullOrEmpty(.Item("AMduId")) Then
                result = String.Format(" {0} AND B.AMduId IN ({1}) ", result, .Item("AMduId"))
            End If


        End With
        Return result
    End Function
    Friend Function QuerySO041() As String
        Return "SELECT NVL(PayKindDefault,0) PayKindDefault FROM SO041"
    End Function
    Friend Function QueryBillType() As String
        Return "select 'B' CodeNo,'收費單' Description from dual " & _
                    " union " & _
                    " select 'T' CodeNo,'臨時收費單' Description from dual " & _
                    " union " & _
                    " select 'I' CodeNo,'裝機單' Description from dual " & _
                    " union " & _
                    " select 'M' CodeNo,'維修單' Description from dual " & _
                    " union " & _
                    " select 'P' CodeNo,'停拆移機單' Description from dual "
    End Function
    Friend Function GetStopYmWhere() As String
        'Return " Max(Decode(Length(C.StopYm),5," & _
        '      " Substr(C.StopYm,2,4) || '0' || Substr(C.StopYm,1,1), " & _
        '      " Substr(C.StopYm,3,4) || Substr(C.StopYm,1,2))) CardExpDate"

        Return "Max( (CASE LENGTH(C.STOPYM) " &
                 " WHEN 5 THEN Substr(to_char(C.StopYm),2,4) || '0' || Substr(to_char(C.StopYm),1,1) " &
                  " Else Substr(to_char(C.StopYm),3,4) || Substr(to_char(C.StopYm),1,2) " &
                 " End )) As CardExpDate "
    End Function
    Friend Function GetMinRealStopDateWhere() As String
        'Return " MIN(DECODE(NVL(A.PAYKIND,0),1, " & _
        '        "DECODE(RealStopDate,NULL,TO_DATE('10000101','YYYYMMDD'),REALSTOPDATE)," & _
        '        "TO_DATE('10000101','YYYYMMDD'))) REALSTOPDATE "

        Return "MIN( (CASE NVL(A.PAYKIND, 0)  WHEN 1 THEN  " &
                        " Case REALSTOPDATE " &
                          "  WHEN NULL THEN TO_DATE('10000101','YYYYMMDD') " &
                          " Else REALSTOPDATE  End " &
                       " Else TO_DATE('10000101','YYYYMMDD')  End )) REALSTOPDATE "
    End Function
    Friend Function QueryIsCrossCustCombine() As String
        Return "Select Nvl(CrossCustCombine,0) From SO041"
    End Function
    Friend Overridable Function QuerySO033Data(ByVal aSO033Where As String, ByVal isZero As Integer,
                                   ByVal isCrossCustCombine As Boolean, ByVal CompCode As String) As String
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
                                        " And SYSDATE >= CStartDate And SYSDATE<= CDueDate And ROWNUM = 1 And COMPCODE = " & CompCode & "  ),-1)  End ) " &
                        "  Else A.CustId  End ) MainCustId " &
                        " From (" & aSQL & ") A,SO001 " &
                " Where A.CustId=SO001.CustId"


        End If

        Return aSQL
    End Function
    Friend Overridable Function GetInvoiceNo2(ByVal StrTableName As String) As String
        Dim strSeq As String
        strSeq = "S_" & StrTableName
        Return "Select '" & Date.Now.ToString("yyMM") &
                     "' || Ltrim(To_Char(" & strSeq & ".NextVal, '0999999')) FROM Dual"

    End Function
    Friend Function UpdMediabillNo(ByVal aWhere As String) As String
        Return String.Format("UPDATE SO033 A SET " &
                                  " MediaBillNo  ={0}0  WHERE   " &
                                  "A.BillNo = {0}1   AND  " & aWhere & " AND " &
                                  "A.AccountNO ={0}2", Sign)
    End Function
    Friend Function QueryMediaIsNull(ByVal aWhere As String, ByVal isZero As Integer) As String
        Return "SELECT A.BillNO, A.MediaBillNO,A.AccountNO," &
                        GetMinPayKindWhere() & " , " & GetMinRealStopDateWhere() & " , " & GetStopYmWhere() &
                        " FROM SO033 A,SO001 B ,SO002 D,SO106 C  " &
                        ",CD013 " &
                        " Where  A.CustID  = D.CUSTID AND " &
                        " A.CustId=B.CustId And A.CancelFlag = 0 And " &
                        " A.UCCode Is Not Null  And " &
                        " A.UCCode=CD013.CodeNo And Nvl(CD013.REFNO,0) NOT IN(3,7) AND  " &
                        " NVL(CD013.PAYOK,0) = 0 AND " &
                        aWhere & " AND " &
                        " A.SERVICETYPE =D.SERVICETYPE  AND   " &
                        " A.AccountNO = C.AccountID   AND A.MediaBillNO IS NULL " &
                        " AND C.StopFlag <> 1 AND C.SnactionDate Is Not Null " &
                        " Group by A.BillNo,A.MediaBillNo,A.AccountNo " &
                        GetHavingOutZeroWhere(isZero)
    End Function
    Friend Function GetUCCodeWhere(ByVal aWhere As String)
        Return " A.CustID  = C.CUSTID AND " & _
                      "A.CustId=B.CustId And A.CancelFlag = 0 And " & _
                      "A.UCCode Is Not Null  " & _
                    aWhere & " AND " & _
                      "A.AccountNO = C.AccountID AND A.CustID = D.CustID  AND " & _
                      "A.SERVICETYPE = D.SERVICETYPE " & _
                      " AND C.StopFlag<> 1 AND C.SnactionDate IS NOT NULL "
    End Function
    Friend Overridable Function QueryViewSQL(ByVal viewName As String) As String
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
                                    " where a.billno=billno and  A.AccountNo = AccountNo And rownum<=1) TCBbudget ," &
                        " (select PTCode from " & viewName &
                                    " where a.billno=billno and  A.AccountNo = AccountNo And rownum<=1) PTCode ," &
                        " (select ServiceType from " & viewName &
                                    " where a.billno=billno and  A.AccountNo = AccountNo And rownum<=1) ServiceType ," &
                         " (select CVC2 from " & viewName &
                                    " where a.billno=billno and  A.AccountNo = AccountNo And rownum<=1) CVC2 ," &
                        " (select CardExpDate from " & viewName &
                                    " where a.billno=billno  and  A.AccountNo = AccountNo And rownum<=1) CardExpDate " &
                        " From (" & aSQL & ") A"
        Return aSQL
    End Function
    Friend Overridable Function QuerySO1108A() As String
        Return String.Format("Select * From (select Parameters from so1108a where programid = 'SO3272A3' and entryid={0}0 order by exectime desc) Where rownum =1", Sign)
    End Function
    Friend Overridable Function UpdLogData() As String
        Return String.Format("UPDATE SO1108A SET EXECSTATUS={0}0 ,EXECMESSAGE={0}1 " &
                             " ,FINISHTIME = SYSDATE,DOWNLOADFILENAME = {0}2 " &
                        " WHERE SEQNO = {0}3", Sign)
    End Function
    Friend Overridable Function UpdUCCode(ByVal IsFubonIntegrate As Boolean) As String
        If IsFubonIntegrate Then
            Return String.Format("UPDATE SO033  Set UCCode={0}0,UCName= {0}1 " &
                               ",UPDEN={0}2,UPDTIME={0}3,CMCode = {0}4,CMName = {0}5 " &
                           " Where RowId  In (Select A.RowId From " &
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
                           " Where RowId  In (Select A.RowId From " &
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


    End Function
    Friend Function CreateView(ByVal viewName As String, ByVal aSQL As String) As String
        Return "Create View " & viewName & " As (" & aSQL & ")"
    End Function
    Friend Overridable Function GetViewName() As String
        Return "SELECT trim(To_Char(S_TMPRPT_ViewName.NextVal, '0999999')) FROM Dual"
        'Return "TMP_" & GetRsValue("SELECT trim(To_Char(" & TableOwnerName & "S_TMPRPT_ViewName.NextVal, '0999999')) FROM Dual", cn)
    End Function
    Friend Function GetMinPayKindWhere() As String
        Return " MAX(Nvl(A.PAYKIND,0)) PAYKIND "
    End Function
    Friend Function GetHavingOutZeroWhere(ByVal isZero As Integer) As String
        If isZero = 0 Then
            Return " Having Sum(A.ShouldAmt)>0 "
        Else
            Return ""
        End If

    End Function
    Friend Function QueryCD112() As String
        Return "select codeno,description from cd112 where stopflag<> 1 order by codeno "
    End Function
    Friend Function QuerySO202() As String
        Return "select mduid codeno,name description from so202  order by mduid"
    End Function
    Friend Function QuerySO017() As String
        Return "select mduid codeno,name description from so017  order by mduid"
    End Function
    Friend Function QueryUpdUCCode() As String
        Return "Select * From CD013 Where RefNo=4 And StopFlag<>1 Order By CodeNo Desc"
    End Function
    Friend Function dropView(ByVal viewName As String) As String
        Return "Drop View " & viewName
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
