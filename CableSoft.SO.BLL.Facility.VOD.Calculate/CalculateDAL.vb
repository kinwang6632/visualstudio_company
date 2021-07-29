Imports CableSoft.BLL.Utility

Public Class CalculateDAL
    Inherits DALBasic
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Public Function QueryClosePara() As String
        Return String.Format("Select * From SO062 WHERE TYPE = 4 " & _
                                    " AND COMPCODE = {0}0 AND SERVICETYPE = {0}1", Sign)
    End Function
    Friend Function QueryChargePara() As String
        Return String.Format("Select Nvl(para35,0) para35,Nvl(para36,0) para36,Nvl(Para14,0) Para14 " & _
                             " From  SO043 " & _
                             "WHERE COMPCODE = {0}0 AND SERVICETYPE = {0}1", Sign)

    End Function
    Friend Function getResultTable(ByVal viewName As String) As String
        Return String.Format("Select * From {0}", viewName)
    End Function
    Friend Function GetExcelWhere2(ByVal xlsTableName As String) As String
        If String.IsNullOrEmpty(xlsTableName) Then
            Return Nothing
        Else
            Return " ( A.FACISNO || A.VODACCOUNTID NOT IN ( SELECT FACISNO || VODACCOUNTID FROM " & xlsTableName & ")) "
        End If        
    End Function
    Friend Function GetExcelWhere1(ByVal xlsTableName As String, ByVal forceBill As Boolean, Para35 As Integer) As String
        If String.IsNullOrEmpty(xlsTableName) Then
            Return Nothing
        End If
        If forceBill Then
            Return " ( A.FACISNO || A.VODACCOUNTID IN ( SELECT FACISNO || VODACCOUNTID FROM " & xlsTableName & ") " & _
                              " AND (NVL(A.MUSTPAYCREDIT,0) - NVL(A.UNPAY,0)-NVL(A.OverCredit,0) ) <  -1 ) "
        Else
            Return " ( A.FACISNO || A.VODACCOUNTID IN ( SELECT FACISNO || VODACCOUNTID FROM " & xlsTableName & ") " & _
                              " AND (NVL(A.MUSTPAYCREDIT,0) - NVL(A.UNPAY,0)-NVL(A.OverCredit,0) ) <  " & Para35 & " ) "
        End If
    End Function
    Friend Function GetMvodIdWhere(ByVal MVodIdString As String) As String
        Dim result As String = Nothing
        result = "Select MVodId From SO004G H Where A.VodAccountId= H.VodAccountId "
        If Not String.IsNullOrEmpty(MVodIdString) Then
            result = String.Format(result & " AND MVODID IN ({0})", MVodIdString)
        End If
        Return result
    End Function
    Friend Function GetWhereString(ByRef dsSource As DataSet, ByVal xlsTableName As String) As String
        Dim result As String = Nothing
        Try

            result = String.Format(" A.CompCode = {0}", dsSource.Tables("Para").Rows(0).Item("CompCode"))
            result = String.Format(result & " And A.ServiceType = '{0}'", dsSource.Tables("Para").Rows(0).Item("ServiceType"))
            If Not DBNull.Value.Equals(dsSource.Tables("Para").Rows(0).Item("CustId")) Then
                If Integer.Parse(dsSource.Tables("Para").Rows(0).Item("CustId") & "") > 0 Then
                    result = String.Format(result & " And A.CustId = {0}", dsSource.Tables("Para").Rows(0).Item("CustId"))
                End If
            End If
            If Not DBNull.Value.Equals(dsSource.Tables("Para").Rows(0).Item("FaciSeqNo")) Then
                result = String.Format(result & " And A.SEQNO = '{0}'", dsSource.Tables("Para").Rows(0).Item("FaciSeqNo"))
            End If
            If Not String.IsNullOrEmpty(xlsTableName) Then
                result = String.Format(result & " And  A.CUSTID || A.VODACCOUNTID || A.FACISNO  " & _
                         " IN ( SELECT CUSTID || VODACCOUNTID || FACISNO " & _
                         " From " & xlsTableName & ")")
            End If           
            result = result & " And  A.FaciCode In ( SELECT CODENO FROM CD022 " & _
                        " WHERE NVL(StopFlag,0)=0 AND NVL(REFNO,0)= 3 )"
        Catch ex As Exception
            Throw
        End Try
        Return result
    End Function
    Friend Function GetSQL(ByRef dsSource As DataSet,
                            ByVal MVodIdWhere As String,
                            ByVal SQLWhere As String,
                            ByVal ExcelWhere1 As String,
                            ByVal ExcelWhere2 As String) As String
        Dim result As String = Nothing
        Dim Sql1 As String = Nothing
        Dim Sql2 As String = Nothing
        Dim forceBill As Boolean = dsSource.Tables("Para").Rows(0).Item("ForceAmount")
        Dim EndDate As String = dsSource.Tables("Para").Rows(0).Item("EndDate")
        Dim aShouldDate As String = Date.Today.ToString("yyyy/MM/dd")
        Dim aPresentSQL As String = "0"
        Dim aMVodPresent As String = "0"
        Dim hasOverCredit As String = Nothing
        Dim hasOverCredit2 As String = Nothing
        If Not DBNull.Value.Equals(dsSource.Tables("Para").Rows(0).Item("ShouldDate")) Then
            aShouldDate = dsSource.Tables("Para").Rows(0).Item("ShouldDate")
        End If
        If IsDate(EndDate) Then EndDate = Date.Parse(EndDate).ToString("yyyyMMdd")
        EndDate = EndDate.Replace("/", "").Replace(" ", "")


        Sql1 = "SELECT RANK() OVER (PARTITION BY B.VODACCOUNTID ORDER BY B.SEQNO DESC) RankX," & _
                " A.NOPAYCREDIT,A.PREPAY,( A.PRESENT - (" & aPresentSQL & ")) PRESENTX,0 PRESENT ,B.VODACCOUNTID,B.SERVICETYPE,B.COMPCODE," & _
                " B.CUSTID,B.SEQNO,B.FACICODE,B.FACISNO,B.RESEQNO" & _
                " FROM SO004G A,SO004 B " & _
                " WHERE A.VODACCOUNTID = B.VODACCOUNTID AND A.NOPAYCREDIT > A.PREPAY - (A.PRESENT - (" & aPresentSQL & "))" & _
                " AND B.FACICODE IN (SELECT CODENO FROM CD022 WHERE REFNO = 3 AND NVL(STOPFLAG,0) =0 ) "

        hasOverCredit = " ( select Nvl(sum(Nvl(usecredit,0)),0) From SO182  " & _
                                  " WHERE so182.creditseqno  in (select seqno " & _
                                  " From SO033VOD " & _
                                  " Where SO182.vodaccountid = SO033vod.vodaccountid " & _
                                  " AND SO033VOD.Seqno Is Not Null " & _
                                  " AND TRUNC(so033vod.INCURREDDATE)> TRUNC( TO_Date('" & EndDate & "','yyyymmdd'))) " & _
                                  " And SO182.VodAccountId = A.VodaccountId) hasOverCredit"


        hasOverCredit2 = " ( select Nvl(sum(Nvl(usecredit,0)),0) From SO182  " & _
                                  " WHERE so182.creditseqno  in (select seqno " & _
                                  " From SO033VOD " & _
                                  " Where SO182.vodaccountid = SO033vod.vodaccountid " & _
                                  " AND SO033VOD.Seqno Is Not Null " & _
                                  " AND SO033VOD.VODACCOUNTID IN (Select VodAccountId From SO004G H " & _
                                  " Where H.MvodId = ( ( " & MVodIdWhere & "))) " & _
                                  " AND TRUNC(so033vod.INCURREDDATE)> TRUNC( TO_Date('" & EndDate & "','yyyymmdd'))) " & _
                                  " And SO182.VODAccountId = A.VodAccountId ) hasOverCredit2"


        Sql1 = "SELECT (A.NOPAYCREDIT - A.PREPAY - A.PRESENT ) MustPayCredit, " & _
                    " (SELECT NVL(SUM(C.SHOULDAMT),0) FROM  SO033 C " & _
                    " WHERE C.UCCODE IS NOT NULL AND NVL(C.CANCELFLAG,0) = 0 " & _
                    " AND A.CUSTID = C.CUSTID AND (A.SEQNO = C.FACISEQNO OR A.RESEQNO = C.FACISEQNO ) " & _
                    " AND C.UCCODE NOT IN (SELECT CODENO FROM CD013 WHERE REFNO IN (3,7,8) OR PAYOK = 1) " & _
                    " AND C.CITEMCODE IN (SELECT CODENO FROM CD019 WHERE REFNO =21)) UNPAY," & _
                    " (SELECT Decode(Min(incurredDate),null,sysdate +365,Min(incurredDate)) FROM (" & _
                    "  SELECT VODACCOUNTID,INCURREDDATE,SEQNO,VALUE FROM SO033VOD D Where D.CloseBillNo is null " & _
                    " MINUS " & _
                    " SELECT F.VODACCOUNTID,F.INCURREDDATE,F.SEQNO,VALUE " & _
                    " FROM SO182SO033VOD E ,SO033VOD F " & _
                    " WHERE E.CREDITSEQNO = F.SEQNO ) " & _
                    " B WHERE B.VODACCOUNTID IN (SELECT VodAccountId From SO004G H Where MVodId = ( " & _
                    MVodIdWhere & ")) " & _
                    " AND VALUE >0 ) MINUSEDATE," & _
                    " (SELECT NVL(SUM(VALUE),0)  FROM SO033VOD B WHERE A.VODACCOUNTID = B.VODACCOUNTID " & _
                    " AND TRUNC(INCURREDDATE)>TRUNC( TO_Date('" & EndDate & "','yyyymmdd'))) OverCredit ," & _
                    hasOverCredit & ",  A.*, " & _
                    "Nvl((Select Nvl(Sum(G.NOPAYCREDIT),0) -Nvl(SUM(G.PREPAY),0) - (0 - " & aMVodPresent & ") " & _
                        " From SO004G G where MVodId = (" & MVodIdWhere & ") ),0) MustPayCredit2, " & _
                    "Nvl((Select Sum(value) From SO033Vod B " & _
                        " Where B.VodAccountId in  (Select VodAccountId From SO004G H " & _
                         " Where H.MvodId = ( ( " & MVodIdWhere & "))) " & _
                         " AND TRUNC(B.INCURREDDATE)> TRUNC(TO_DATE('" & EndDate & "','yyyymmdd'))),0) OverCredit2, " & hasOverCredit2 & _
                    " FROM ( " & Sql1 & " ) "


           
        Sql2 = Sql1

        Sql1 = "SELECT A.*,0 FLAG,B.Para35,B.Para36 FROM ( " & Sql1 & " A WHERE A.RANKX=1) A,SO043 B " & _
           " WHERE A.SERVICETYPE = B.SERVICETYPE AND (NVL(A.MUSTPAYCREDIT2,0) - NVL(A.UNPAY,0) - NVL(A.OVERCREDIT2,0)+NVL(A.HASOVERCREDIT2,0)  ) > 0 " & _
           " AND (NVL(A.MUSTPAYCREDIT,0) - NVL(A.UNPAY,0) - NVL(A.OVERCREDIT,0) +NVL(A.HASOVERCREDIT,0)) > 0 " & _
           " AND ((nvl(A.MUSTPAYCREDIT2,0) - nvl(A.UNPAY,0)-nvl(A.OverCredit2,0)+Nvl(A.HasOverCredit2,0) ) >=" & IIf(forceBill, "1", " nvl(B.PARA35,0) ") & _
           " OR ADD_MONTHS(TO_DATE(TO_CHAR(A.MINUSEDATE,'yyyymm') || '01','yyyymmdd'),NVL(B.PARA36,0)) <=To_Date('" & EndDate & "','yyyymmdd') " & _
           IIf(String.IsNullOrEmpty(ExcelWhere1), "", " OR " & ExcelWhere1) & " )" & _
           " AND " & SQLWhere

        Sql2 = "SELECT A.*,1 FLAG,B.Para35,B.Para36 FROM ( " & Sql2 & " A WHERE A.RANKX=1) A,SO043 B " & _
            " WHERE A.SERVICETYPE = B.SERVICETYPE AND (NVL(A.MUSTPAYCREDIT2,0) - NVL(A.UNPAY,0) - NVL(A.OVERCREDIT2,0)+NVL(A.HASOVERCREDIT2,0) ) > 0 " & _
            " AND ((nvl(A.MUSTPAYCREDIT2,0) - nvl(A.UNPAY,0)-nvl(A.OverCredit2,0)+nvl(A.HasOverCredit2,0) ) < " & IIf(forceBill, "-1", " nvl(B.PARA35,0) ") & _
            " AND ADD_MONTHS(TO_DATE(TO_CHAR(A.MINUSEDATE,'yyyymm') || '01','yyyymmdd'),NVL(B.PARA36,0)) > To_Date('" & EndDate & "','yyyymmdd') " & _
            IIf(String.IsNullOrEmpty(ExcelWhere2), "", " AND " & ExcelWhere2) & " )" & _
            " AND " & SQLWhere

        result = Sql1 & " UNION ALL " & Sql2
        'result = "Select Distinct A.*, To_Date( ShouldDate From ( " & result & " ) A"
        result = String.Format("Select Distinct A.*, To_Date('{0}','yyyy/mm/dd') ShouldDate,To_Date('{1}','yyyymmdd') EndDate From ( {2} ) A",
                               aShouldDate, EndDate, result)
        Return result
    End Function
    Friend Function InsDataToTempTable(ByVal TempName As String) As String
        Dim result As String = String.Format("Insert Into " & TempName & "( " & _
                                            "CustId,VODACCOUNTID,FACISNO) Values " & _
                                            "({0}0,{0}1,{0}2 )", Sign)
        Return result
    End Function
    Friend Function CreateTempXLSTable(ByVal TempName As String) As String
        Return String.Format(" Create Table {0} " & _
            "(CUSTID Number(8), VODACCOUNTID Number(20), FACISNO Varchar2(25), FACISEQNO Varchar2(25))", TempName)
    End Function
    Friend Function DeleteTmpTable(ByVal TmpName As String) As String
        Return String.Format("Drop Table {0}", TmpName)
    End Function
    Friend Function DeleteView(ByVal ViewName As String) As String
        Return String.Format("Drop View {0}", ViewName)
    End Function
    Friend Function GetTmpTableName() As String
        Return "SELECT TRIM(TO_CHAR(S_TMPRPT_ViewName.NextVal,'0999999' )) FROM DUAL"
    End Function
    Friend Function QueryDefaultCitem() As String
        Return String.Format("SELECT CodeNo,Description FROM  CD019 " & _
                            " WHERE NVL(STOPFLAG,0)=0 AND REFNO = 21 " & _
                            " AND SERVICETYPE = {0}0 " & _
                            " AND SIGN = '+'", Sign)
    End Function
    Friend Function QueryDefaultCMCode() As String
        Return String.Format("SELECT A.CMCode,B.Description CMNAME FROM SO044 A,CD031 B " & _
               " WHERE A.CMCODE = B.CODENO AND A.SERVICETYPE= {0}0", Sign)
    End Function
    Friend Function QueryDefaultUCCode() As String
        Return String.Format("SELECT A.UCCODE CodeNo ,B.Description Description FROM SO044 A,CD013 B " & _
                " WHERE A.UCCODE = B.CODENO AND A.SERVICETYPE = {0}0", Sign)
    End Function
    Friend Function QueryDefaultSalePointCode() As String
        Return String.Format("SELECT CodeNo,Description FROM CD107 " & _
            " WHERE NVL(STOPFLAG,0)=0 AND DefaultValue=1 AND COMPCODE = {0}0", Sign)
    End Function
    Friend Function QueryDefaultSO014(ByVal intPara14 As Integer) As String
        Return String.Format("SELECT A.*,B.CLASSCODE1 FROM SO014 A,SO001 B " & _
            " WHERE B.CUSTID={0}0 AND " & IIf(intPara14 = 1, "B.ChargeAddrNo", "B.InstAddrNo") & "=A.AddrNo" & _
            " AND B.COMPCODE={0}1 AND A.COMPCODE=B.COMPCODE", Sign)

    End Function
    Friend Function QueryServiceType() As String
        Return "SELECT CODENO,DESCRIPTION FROM CD046 "
    End Function

    Friend Function GetSO033Sechema() As String
        Return "Select RowId ,A.* From SO033 A Where 1 = 0 "
    End Function
    Friend Function GetSO033VODSechema() As String
        Return "Select RowId,A.* From SO033Vod A Where 1 = 0"
    End Function
    Friend Function QueryInsSO033Data() As String
        Return String.Format("Select RowId,A.* From SO033 A Where BillNo = {0}0 And Item = 1 ", Sign)
    End Function


    Friend Function QuerySO033VODSeqNo(ByVal VODAccountId As String, ByVal EndDate As Date) As String
        Dim IncurredDate As String = EndDate.ToString("yyyyMMdd")
        IncurredDate = String.Format("To_Date('{0}','yyyymmdd')", IncurredDate)

        Return String.Format(" SELECT SEQNO FROM SO033VOD " & _
                                         " WHERE VODACCOUNTID= {0} " & _
                                         " AND CLOSEBILLNO IS NULL AND VALUE > 0 " & _
                                         " AND TRUNC(INCURREDDATE) <= {1} " & _
                                         " MINUS SELECT A.CREDITSEQNO FROM SO182SO033VOD A,SO182 B " & _
                                         " WHERE A.HISCREDITSEQNO=B.SEQNO AND B.VODACCOUNTID= {0} " & _
                                         " AND ( B.CLOSEBILLNO IS NOT NULL Or B.UseCredit >= B.MinusCredit )", VODAccountId, IncurredDate)

    End Function
    Friend Function QuerySO182SeqNo(ByVal VODAccountId As String, ByVal SO033VODSEQNO As String) As String
        Dim result As String = Nothing
        result = String.Format("SELECT A.SEQNO FROM SO182 A " & _
                                      " WHERE A.VODACCOUNTID = {0} " & _
                                      "  AND NVL(A.ADDCREDIT,0)=0 AND A.CLOSEBILLNO IS NULL " & _
                                      " AND Nvl(UseCredit,0) < Nvl(MinusCredit,0) " & _
                                      " AND A.CREDITSEQNO IN ( {1} )", VODAccountId, SO033VODSEQNO)

        Return result
    End Function
    Friend Function QuerySO182SeqNo(ByVal VODAccountId As String, ByVal EndDate As Date) As String
        Dim result As String = Nothing
      
        Dim SO033VODSeqNo As String = QuerySO033VODSeqNo(VODAccountId, EndDate)
        result = String.Format("SELECT A.SEQNO FROM SO182 A " & _
                                         " WHERE A.VODACCOUNTID = {0} " & _
                                         "  AND NVL(A.ADDCREDIT,0)=0 AND A.CLOSEBILLNO IS NULL " & _
                                         " AND Nvl(UseCredit,0) < Nvl(MinusCredit,0) " & _
                                         " AND A.CREDITSEQNO IN ( {1} )", VODAccountId, SO033VODSeqNo)
      
        Return result
    End Function
    Friend Function QueryUpdSO033VODData() As String
        'Dim result As String =String.Format("Select RowId,A.* From SO033VOD A " & _
        '                                    " Where SEQNO " & _
        Dim result As String = String.Format("SELECT A.ROWID,A.* FROM SO033VOD A " & _
                                " WHERE A.SEQNO IN (SELECT CREDITSEQNO FROM SO182 " & _
                                    " WHERE ROWID={0}0 " & _
                                    " AND CLOSEBILLNO= {0}1)", Sign)

        Return result
    End Function
    Friend Function InsSO062() As String
        Dim result As String = String.Format("Insert Into SO062 " & _
                                            " (Type,TranDate,UpdEn,UpdTime,Para,CompCode,ServiceType) " & _
                                            " Values (4,{0}0,{0}1,{0}2,{0}3,{0}4,{0}5 )", Sign)
        Return result
    End Function
    Friend Function UpdateSO062() As String
        Dim result As String = String.Format("Update SO062 Set  " & _
                                            "TranDate={0}0,UpdEn={0}1,UpdTime={0}2,Para={0}3 " & _
                                            " Where CompCode = {0}4 And ServiceType = {0}5 ", Sign)
        Return result
    End Function
    Friend Function QuerySO062Data() As String
        Dim result As String = String.Format("Select TranDate,UpdEn,UpdTime " & _
                                            " From SO062 Where Type = 4 And ServiceType = {0}0 And CompCode = {0}1", Sign)
        Return result
    End Function
    Friend Function IsExistsSO062() As String
        Dim result As String = String.Format("SELECT count(*) FROM SO062 " & _
                                          " WHERE TYPE =4 AND COMPCODE={0}0 " & _
                                           " AND SERVICETYPE = {0}1", Sign)
        Return result
    End Function
    Friend Function QueryUpdSO182Data(ByVal SeqNo As String) As String
        Dim result As String = String.Format("Select RowId,A.* From SO182 A Where " & _
                                                                    "SEQNO IN ( {0} ) ", SeqNo)

        Return result
    End Function
    Friend Function UpdateSO182(ByVal SeqNo As String) As String
        Dim result As String = Nothing
        result = String.Format("Update SO182 Set CloseTime = {0}0," & _
                                            " CloseEN={0}1,CloseName={0}2, " & _
                                            " UpdEn = {0}3,UPDTime = {0}4, " & _
                                            " CloseBillNo = {0}5,CloseItem=1,CloseFlag=1 " & _
                                            " Where SeqNo In (" & SeqNo & " )", Sign)


        
        Return result
    End Function
    Friend Function UpdateSO033VOD(ByVal SeqNo As String) As String
        Dim result As String = String.Format(" UPDATE SO033VOD SET " & _
                            " CloseBillNo={0}0," & _
                            " CloseItem=1" & _
                            " WHERE SEQNO IN (" & SeqNo & " )", Sign)

        
        Return result
    End Function
    Friend Function InsertSO033Test() As String
        Dim result As String = Nothing
        result = "Insert Into SO033 (CustId,CompCode,BillNo,Item,CitemCode,CitemName,ShouldDate,OldAmt,ShouldAmt " & _
                        ",OldPeriod,RealPeriod,CMCODE,CMName,UCCODE,UCName,PTCODE,PTName" & _
                        ",ClassCode,FaciSeqNo,FaciSNO,ServiceType) " & _
                    " Values ({0},{1},'{2}',1,{3},'{4}',To_Date('{5}','yyyy/MM/dd'),{6},{7},{8},{9} " & _
                    ",{10},'{11}',{12},'{13}',{14},'{15}',{16},'{17}','{18}','{19}')"

        'result = String.Format("Insert Into SO033 (CustId,CompCode,BillNo,Item,CitemCode, " & _
        '            "ShouldDate,OldAmt,ShouldAmt,OldPeriod,RealPeriod,CMCODE,CMName, " & _
        '           "CitemName ) " & _
        '            " Values ({0}0,{0}1,{0}2,1,{0}3, " & _
        '            "{0}4,{0}5,{0}6,{0}7,{0}8,{0}9,{0}10, " & _
        '            "{0}11)", Sign)
        Return result
    End Function
    Friend Function InsertSO033() As String
        Dim result As String = Nothing
        result = String.Format("Insert Into SO033 (CustId,CompCode,BillNo,Item,CitemCode, " & _
                    "ShouldDate,OldAmt,ShouldAmt,OldPeriod,RealPeriod,CMCODE,CMName, " & _
                    "UCCODE,UCName,PTCODE,PTName,ClassCode,FaciSeqNo,FaciSNO," & _
                    "ServiceType,Note,CreateTime,UpdTime,CreateEn,UpdEn,AddrNo," & _
                    "StrtCode,MduId,ServCode,ClctAreaCode,OldClctEn,OldClctName," & _
                    "AreaCode,ClctEn,ClctName,SalePointCode,SalePointName,CloseStopDate,CitemName," & _
                    "NewUpdTime ) " & _
                    " Values ({0}0,{0}1,{0}2,1,{0}3, " & _
                    "{0}4,{0}5,{0}6,{0}7,{0}8,{0}9,{0}10, " & _
                    "{0}11,{0}12,{0}13,{0}14,{0}15,{0}16,{0}17," & _
                    "{0}18,{0}19,{0}20,{0}21,{0}22,{0}23,{0}24," & _
                    "{0}25,{0}26,{0}27,{0}28,{0}29,{0}30, " & _
                    "{0}31,{0}32,{0}33,{0}34,{0}35,{0}36,{0}37,{0}38 )", Sign)

        'result = String.Format("Insert Into SO033 (CustId,CompCode,BillNo,Item,CitemCode, " & _
        '            "ShouldDate,OldAmt,ShouldAmt,OldPeriod,RealPeriod,CMCODE,CMName, " & _
        '           "CitemName ) " & _
        '            " Values ({0}0,{0}1,{0}2,1,{0}3, " & _
        '            "{0}4,{0}5,{0}6,{0}7,{0}8,{0}9,{0}10, " & _
        '            "{0}11)", Sign)
        Return result
    End Function
    Friend Function GetSO182Sechema() As String
        Return "Select RowId,A.* From SO182 A Where 1 = 0"
    End Function
    Friend Function QueryCompCode() As String
       
            Return String.Format("Select A.CodeNo,A.Description  " & _
                              " From CD039 A,SO026 B  " & _
                              " Where Instr(','||B.CompStr||',',','||A.CodeNo||',')>0 " & _
                             " And UserId = {0}0 Order By CodeNO", Sign)        
    End Function
    Friend Function QueryCompCode(ByVal GroupId As String) As String
        If GroupId = "0" AndAlso 1 = 0 Then
            Return "Select A.CodeNo ,A.Description From CD039 A Order By CodeNo"
        Else
            Return String.Format("Select A.CodeNo,A.Description  " & _
                              " From CD039 A,SO026 B  " & _
                              " Where Instr(','||B.CompStr||',',','||A.CodeNo||',')>0 " & _
                             " And UserId = {0}0 Order By CodeNO", Sign)
        End If
    End Function
    Friend Function QueryMVodId(ByVal VodAccountIdS As String) As String
        Return String.Format("select distinct MvodId from SO004G where VodAccountId In ({0})", VodAccountIdS)
    End Function
End Class
