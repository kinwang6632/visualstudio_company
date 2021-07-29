Imports CableSoft.BLL.Utility
Public Class AccountDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryAccount() As String
        Dim aRet As String = String.Format("Select * From SO106 Where MasterId = {0}0", Sign)
        Return aRet
    End Function
    Friend Overridable Function QueryAccountDetail() As String
        Return String.Format("select rowid as ctid,a.* from so106a a where masterid = {0}0", Sign)
    End Function
    Friend Overridable Function QuerySO106Log() As String
        Return String.Format("Select A.rowid as ctid,A.* From SO106  A where Masterid = {0}0", Sign)
    End Function
    Friend Function DeleteSO106(ByVal strWhere As String) As String
        Return "DELETE FROM SO106 WHERE " & strWhere
    End Function
    Friend Function DeleteSO106A(ByVal MasterId As String, ByVal AchtNo As String) As String
        Return String.Format("DELETE FROM SO106A WHERE MasterId = {0} " &
                    " AND ACHTNO ='{1}' AND AuthorizeStatus = 4 ", MasterId,
                    AchtNo.Trim("'"))
    End Function
    Friend Function QueryStartPost() As String
        Return String.Format("Select Nvl(StartPost,0) From SO041 Where SysID ={0}0", Sign)
    End Function
    Friend Function UpdSO033() As String
        Dim aUpdSQL As String = String.Format("Update SO033 Set AccountNo = {0}0, BankCode = {0}1, " & _
                        " BankName = {0}2, CMCode ={0}3, CMName={0}4," & _
                        " PTCode = {0}5, PTName={0}6,UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = {0}9 " & _
                        " Where BillNo = {0}10 And Item = {0}11", Sign)
        Return aUpdSQL
    End Function
    Friend Function UpdOldSO033() As String
        Dim aUpdSQL As String = String.Format("Update SO033 Set AccountNo = {0}0, BankCode = {0}1, " & _
                        " BankName = {0}2, CMCode ={0}3, CMName={0}4," & _
                        " PTCode = {0}5, PTName={0}6,UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = {0}9 ,INVSEQNO = {0}10" & _
                        " Where BillNo = {0}11 And Item = {0}12", Sign)
        Return aUpdSQL
    End Function
    Friend Function ClearOldSO033() As String
        Dim aUpdSQL As String = String.Format(
                                    "UPDATE SO033 SET " &
                                    "BANKCODE=NULL" &
                                    ",BANKNAME=NULL" &
                                    ",ACCOUNTNO=NULL" &
                                    ",INVSEQNO = NULL" &
                                    ",CMCODE= {0}0" &
                                    ",CMNAME={0}1" &
                                    ",PTCODE={0}2" &
                                    ",PTNAME={0}3" &
                                    ",UCCode = {0}4" &
                                    ",UCName = {0}5" &
                                    ",UpdEn = {0}6" &
                                    ",UpdTime = {0}7 " &
                                    ",NewUpdTime = {0}8 " &
                                    " WHERE BILLNO = {0}9" &
                                    " AND ITEM = {0}10 " &
                                    " AND UCCODE IS NOT NULL " &
                                    " AND UCCODE  NOT IN (Select CodeNo " &
                                    "           From CD013 Where PayOk = 1 Or RefNo in (3,7,8))  " &
                                    " AND NVL(CANCELFLAG,0)=0", Sign)

        Return aUpdSQL
    End Function
    Friend Function ClearSO033() As String
        Dim aUpdSQL As String = String.Format(
                                    "UPDATE SO033 SET " &
                                    "BANKCODE=NULL" &
                                    ",BANKNAME=NULL" &
                                    ",ACCOUNTNO=NULL" &
                                    ",CMCODE= {0}0" &
                                    ",CMNAME={0}1" &
                                    ",PTCODE={0}2" &
                                    ",PTNAME={0}3" &
                                    ",UCCode = {0}4" &
                                    ",UCName = {0}5" &
                                    ",UpdEn = {0}6" &
                                    ",UpdTime = {0}7 " &
                                    ",NewUpdTime = {0}8 " &
                                    " WHERE BILLNO = {0}9" &
                                    " AND ITEM = {0}10 " &
                                    " AND UCCODE IS NOT NULL " &
                                    " AND UCCODE  NOT IN (Select CodeNo " &
                                    "           From CD013 Where PayOk = 1 Or RefNo in (3,7,8))  " &
                                    " AND NVL(CANCELFLAG,0)=0", Sign)

        Return aUpdSQL
    End Function
    Friend Function QueryNewChooseProduct(ByVal strProServiceID As String) As String
        Dim aRet As String = String.Format("Select * From SO003C Where MasterId = {0}0  " &
                                         " And (( InstDate is null ) Or ( PRDate is null ) Or ( InstDate >  PrDate )) " &
                                         " Or ServiceId In(" & strProServiceID & ")", Sign)
        Return aRet
    End Function
    Friend Function QueryChooseProduct() As String
        Dim aRet As String = String.Format("Select * From SO003C Where MasterId = {0}0 And CustId = {0}1 " & _
                                          " And (( InstDate is null ) Or ( PRDate is null ) Or ( InstDate >  PrDate )) ", Sign)
        Return aRet
    End Function
    Friend Function QueryNoUseAchCustId(ByVal AccountId As String, ByVal MasterId As String, ByVal WhereIn As String) As String
        Return String.Format("SELECT ACHCUSTID From SO106 " &
                                       " Where SUBSTR(LPAD(AccountId,30,'0'),25,6) = '{0}' " & _
                                       " And MasterId <> {1}" & _
                                       " AND SUBSTR(ACHCUSTID,4,6) = '{2}' " & _
                                       " AND ACHCUSTID IN ({3}) ORDER BY ACHCUSTID ",
                                       AccountId,
                                       MasterId,
                                       AccountId, WhereIn)
    End Function
    Friend Function GetACHCustId() As String
        Return String.Format(" SELECT NVL(ACHCUSTID,0) FROM SO041 " & _
                                                               " WHERE SYSID={0}0", Sign)
    End Function
    Friend Overridable Function GetSO137CustId() As String
        'Return String.Format("Select * from SO001 Where CustId In  (Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                          " where a.seqno ={0}0 And a.memberid = b.memberid  " & _
        '                                          " and b.HomeId = c.HomeId )", Sign)
        Return String.Format("Select * from SO001 Where CustId In  (Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno ={0}0 And a.seqno = b.memberid  " &
                                                " and b.HomeId = c.HomeId )", Sign)
    End Function
    Friend Overridable Function GetNewCitemCode(ByVal FaciSeqNos As String, ByVal ProductCodes As String) As String


        Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno = {0}0 And a.seqno = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
        Return String.Format("Select SeqNo,CitemCode,CitemName From SO003 " &
               " Where CustId In (" & aCustId & ") And FaciSeqNo IN (" & FaciSeqNos & ") And NVL(StopFlag,0) = 0 " &
               " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN (" & ProductCodes & "))",
               Sign)

    End Function
    Friend Overridable Function GetNewCitemCode(ByVal SeqNo As String, ByVal FaciSeqNos As String, ByVal ProductCodes As String) As String
        'Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                          " where a.seqno = " & SeqNo & " And a.memberid = b.memberid  " & _
        '                                          " and b.HomeId = c.HomeId "

        Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno = " & SeqNo & " And a.seqno = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
        Return String.Format("Select SeqNo,CitemCode,CitemName From SO003 " &
               " Where CustId In ({0}) And FaciSeqNo IN ({1}) And NVL(StopFlag,0) = 0 " &
               " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN ({2}))",
               aCustId, FaciSeqNos, ProductCodes)

    End Function
    Friend Function GetCitemCode(ByVal FaciSeqNos As String, ByVal ProductCodes As String) As String
        Return String.Format("Select SeqNo,CitemCode,CitemName From SO003 " &
            " Where CustId = {0}0 And FaciSeqNo IN (" & FaciSeqNos & ") And NVL(StopFlag,0) = 0 " &
            " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN (" & ProductCodes & "))",
             Sign)
    End Function
    Friend Function GetCitemCode(ByVal CustId As String, ByVal FaciSeqNos As String, ByVal ProductCodes As String) As String
        Return String.Format("Select SeqNo,CitemCode,CitemName From SO003 " &
                " Where CustId = {0} And FaciSeqNo IN ({1}) And NVL(StopFlag,0) = 0 " &
                " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN ({2}))",
                CustId, FaciSeqNos, ProductCodes)


    End Function
    Friend Function getOldCitemCode(ByVal CustId As String, ByVal CitemStr As String) As String
        Dim result As String = Nothing
        result = "Select SeqNo, CitemCode, CitemName From SO003 " &
                " Where CustId = " & CustId & "  And NVL(StopFlag, 0) = 0 " &
                " And SeqNo In (" & CitemStr & ")"
        Return result
    End Function
    Friend Function getSO014AddressByAddrNo() As String
        Return String.Format("Select Address From SO014 Where AddrNo = {0}0", Sign)
    End Function
    Friend Function insertSO002AD() As String
        Dim result As String = Nothing
        result = String.Format("Insert into SO002AD (AccountNo,CompCode,CustId,InvSeqNo) " & _
                               " values ({0}0,{0}1,{0}2,{0}3)", Sign)

        Return result
    End Function
    Friend Function insertSO138() As String
        Dim result As String = Nothing
        result = String.Format("Insert into SO138(InvSeqNo,ChargeTitle,InvoiceType,InvNo, " & _
            "InvTitle,InvPurposeCode,InvPurposeName,PreInvoice," & _
            "BillMailKind,DenRecCode,DenRecName,DenRecDate," & _
            "LoveNum,InvoiceKind,ApplyInvDate,ChargeAddrNo,ChargeAddress," & _
            "MailAddrNo,MailAddress,UpdTime,NEWUPDTIME,UpdEn ) Values (" & _
            "{0}0,{0}1,{0}2,{0}3," & _
            "{0}4,{0}5,{0}6,{0}7," & _
            "{0}8,{0}9,{0}10,{0}11," & _
           " {0}12,{0}13,{0}14,{0}15,{0}16," & _
           " {0}17,{0}18,{0}19,{0}20,{0}21 )", Sign)
        
        Return result
    End Function
    Friend Function getInvPurposeNameByCode() As String
        Return String.Format("Select Description From CD095 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function getDenRecNameByCode() As String
        Return String.Format("Select Description From CD110 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function UpdateACHSO003C() As String
        Return String.Format("Update SO106 Set ProServiceID = {0}0 Where MasterId= {0}1", Sign)
    End Function
    Friend Function UpdateSO003C(ByVal ServiceIds As String, ByVal strNewUpdTime As String) As String
        Return String.Format("UPDATE SO003C Set MASTERID={0}0, " &
                              " UpdEn ={0}1, " &
                              " UpdTime = {0}2 , NewUpdTime = To_Date('" & strNewUpdTime & "','yyyymmddhh24miss') " &
                                 "  WHERE ServiceId IN (" & ServiceIds & ") ",
                                             Sign)
    End Function
    Friend Function UpdateSO003C(ByVal rw As DataRow, ByVal ServiceIds As String) As String
        Return String.Format("UPDATE SO003C SET MASTERID={0}, " & _
                              " UpdEn ='{1}',  " & _
                              " UpdTime = '{2}' ,NewUpdTime = To_Date('{3}','yyyymmddhh24miss') " & _
                                 "  WHERE ServiceId IN ({4}) ",
                                             rw.Item("MASTERID"), rw.Item("UpdEn"),
                                            rw.Item("UpdTime"),
                                            CType(rw.Item("NewUpdTime"), Date).ToString("yyyyMMddHHmmss"),
                                            ServiceIds)
    End Function
    Friend Overridable Function TakeSO106SeqNo() As String
        Return "select S_SO106_MasterId.NextVal from dual"
    End Function
    Friend Function IsACHBank(ByVal blnStartPost As Boolean) As String
        If blnStartPost Then
            Return String.Format("Select Count(*) From CD018 Where CodeNo ={0}0  And (PRGNAME LIKE 'ACH%' Or PRGNAME LIKE '%POST4%') ",
                                                            Sign)
        Else
            Return String.Format("Select Count(*) From CD018 Where CodeNo ={0}0  And PRGNAME LIKE 'ACH%'",
                                                            Sign)
        End If
        
    End Function
    Friend Function IsACHBank() As String

        Return String.Format("Select Count(*) From CD018 Where CodeNo ={0}0  And PRGNAME LIKE 'ACH%'",
                                                            Sign)
    End Function
    Friend Function IsStartPos() As String
        Return "Select Nvl(StartPost,0) StartPost from SO041"
    End Function
    Friend Function GetAchHeadCode() As String
        Return String.Format("SELECT ACHHeadCode FROM SO041 WHERE SYSID = {0}0", Sign)
    End Function
    Friend Function UpdateSO004(ByVal rw As DataRow, ByVal SEQNOs As String) As String
        Return String.Format("UPDATE SO004 SET MASTERID={0} WHERE SEQNO IN ({1}) ",
                                              rw.Item("MASTERID"), SEQNOs)
    End Function
    Friend Function IsAddCancelAuth() As String
        'String.Format("select count(*) from so106a " &
        '        " where AuthorizeStatus=2 and achtno='{0}'" &
        '        " and masterid={1}", aAchtNo.Trim("'"), aMasterId)
        Return String.Format("select count(*) from so106a " &
                " where AuthorizeStatus=2 and achtno={0}0" &
                " and masterid={0}1", Sign)
    End Function
    Friend Overridable Function GetSO106RowId() As String
        Return String.Format("select rowid as ctid from so106 where masterid={0}0", Sign)
    End Function
    Friend Function GetDefUCCode() As String
        Return "Select * From CD013 Where RefNo = 1"
    End Function
    Friend Function QueryUccode() As String
        Return String.Format("SELECT CODENO,Description FROM CD013 " & _
                                          " WHERE CODENO = (SELECT UCCODE FROM SO044 " & _
                                          " WHERE SERVICETYPE={0}0 " & _
                                          " AND COMPCODE= {0}1 )" & _
                                          " AND STOPFLAG<>1", Sign)
    End Function
    Friend Function updOldSO033Uccode() As String
        Return String.Format("UPDATE  SO033 SET " & _
                                           "BANKCODE=NULL" & _
                                           ",BANKNAME=NULL" & _
                                           ",ACCOUNTNO=NULL" & _
                                           ",InvSeqNo=NULL" & _
                                           ",CMCODE= {0}0 " & _
                                           ",CMNAME= {0}1 " & _
                                           ",PTCODE= {0}2 " & _
                                           ",PTNAME= {0}3 " & _
                                           ",UCCODE= {0}4 " & _
                                           ",UCNAME= {0}5 " & _
                                           " WHERE CUSTID= {0}6 " & _
                                           " AND COMPCODE= {0}7 " & _
                                           " AND UCCODE > 0 AND CANCELFLAG=0" & _
                                           " AND BILLNO= {0}8", Sign)
    End Function
    Friend Function GetACHTDESC() As String
        'String.Format("Select ACHTDESC From CD068 " & _
        '                                    " Where ACHTNO='{0}'  " & _
        '                                    " AND ACHTYPE =1 ",
        Return String.Format("Select ACHTDESC From CD068 " & _
                                            " Where ACHTNO={0}0  " & _
                                            " AND ACHTYPE =1 ", Sign)
    End Function
    Friend Function DelSO106A() As String
        'String.Format("DELETE SO106A WHERE MasterId = {0} " &
        '            " AND ACHTNO ='{1}' AND AuthorizeStatus = 4 ", Account.Tables(FNewAccountTableName).Rows(aRowIndex).Item("MasterId"),
        '            aAchtNo.Trim("'"))
        Return String.Format("DELETE FROM SO106A WHERE MasterId = {0}0 " &
                    " AND ACHTNO ={0}1 AND AuthorizeStatus = 4 ", Sign)
    End Function
    Friend Function CreateTableSchema(ByVal TableName As String) As String
        Return "SELECT * FROM " & TableName & " WHERE 1=0"
    End Function
    Friend Function ChkIsAuthorize() As String
        Return String.Format("SELECT COUNT(*) FROM SO106A WHERE MASTERID = {0}0 AND AUTHORIZESTATUS IS NULL  ", Sign)
    End Function
    Friend Function VoidSO106Data(ByVal AccountTable As DataTable) As String
        Return String.Format("UPDATE SO106 SET STOPFLAG = 1,STOPDATE = TO_DATE('{0}','yyyymmdd') " & _
                        " WHERE MASTERID = {1}", Format(AccountTable.Rows(0).Item("StopDate"), "yyyyMMdd"),
                AccountTable.Rows(0).Item("MASTERID"))
    End Function
    Friend Function GetSO106A() As String
        Return String.Format("SELECT * FROM SO106A WHERE MASTERID = {0}0", Sign)
    End Function
    Friend Function GetAchCitem(ByVal rw As DataRow, ByVal ServiceIds As String, ByVal ACHNo As String) As String
        'Return String.Format("Select A.*,B.ACHTNO FROM SO003C A,CD046 B " & _
        '                                " Where A.ServiceType = B.CodeNo " & _
        '                                " AND A.SERVICEID IN ( {0} ) " & _
        '                                " AND B.ACHTNO = '{1}' " & _
        '                                " And (( A.InstDate is null ) Or ( A.PRdate is null ) Or ( A.InstDate > A.PRDate))",
        '                                ServiceIds, ACHNo.Trim("'"c))
        Return String.Format("Select A.*,B.ACHTNO FROM SO003C A,CD046 B " & _
                                       " Where A.ServiceType = B.CodeNo " & _
                                       " AND A.SERVICEID IN ( {0} ) " & _
                                       " And (( A.InstDate is null ) Or ( A.PRdate is null ) Or ( A.InstDate > A.PRDate))",
                                       ServiceIds)
    End Function
    Friend Function GetInvSeqNo() As String
        Return String.Format("SELECT INVSEQNO " & _
                                                        " FROM SO002 " & _
                                                        " WHERE CUSTID = {0}0 " & _
                                                        " AND INVSEQNO IS NOT NULL", Sign)
    End Function
    Friend Function QueryChooseFaci() As String
        'Dim aRet As String = String.Format("Select * From SO004 Where MasterId = {0}0 " & _
        '                                   " And CustId = {0}1", Sign)
        Dim aRet As String = String.Format("Select * From SO004 Where MasterId = {0}0 " &
                                           " And CustId = {0}1", Sign)
        Return aRet
    End Function
    Friend Function ClearSO003C(ByVal ServiceID As String, ByVal NewUpdTime As String) As String
        Dim aRet As String = String.Format("Update SO003C Set CMCode = {0}0, " &
            " CMName={0}1, PTCode = {0}2, PTName={0}3, " &
            " MasterId = Null," &
            "UpdEn = {0}4,UpdTime = {0}5, NewUpdTime = To_Date('" & NewUpdTime & "','yyyymmddhh24miss') " &
            " Where ServiceID IN (" & ServiceID & ")", Sign)
        Return aRet

    End Function
    Friend Function ClearSO003C() As String
        Dim aRet As String = "Update SO003C Set CMCode = {0}, " & _
            " CMName='{1}', PTCode = {2}, PTName='{3}', " & _
            " MasterId = Null," & _
            "UpdEn = '{4}',UpdTime = '{5}', NewUpdTime = To_Date('{6}','yyyymmddhh24miss') " & _
            " Where ServiceID IN ({7})"
        Return aRet
    End Function
    Friend Function ClearSO004() As String
        Dim aRet As String = "UPDATE SO004 SET AccountNo = Null, BankCode = Null, " &
                                         "BankName = Null,MasterId = Null," &
                                         "CHCMCODE = {0},ChCMName = '{1}'," &
                                         "ChPTCode = {2}, ChPTName='{3}'" &
                                         " WHERE SEQNO IN ({4})"

        Return aRet
    End Function
    Friend Overridable Function UpdNewNonePeriod(ByVal SeqNo As String) As String


        'Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                           " where a.seqno = {10} And a.memberid = b.memberid  " & _
        '                                           " and b.HomeId = c.HomeId "
        Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                 " where a.seqno = {10} And a.seqno = b.memberid  " &
                                                 " and b.HomeId = c.HomeId "
        Dim aSQL As String = "UPDATE SO003 SET AccountNo='{0}', " &
                       "BankCode = {1},BankName = '{2}',CMCode = {3},CMName='{4}', " &
                       "PTCode = {5},PTName = '{6}',UpdEn = '{7}',UpdTime = '{8}',NewUpdTime = To_Date('{9}','yyyymmddhh24miss')  " &
                       " WHERE 1=1 AND CUSTID In (" & aCustId & ") " &
                       " And SEQNO IN (" & SeqNo & ") " &
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode is Null)"
        Return aSQL
    End Function
    Friend Function ClearOldNoneSO003(ByVal SEQNO As String, ByVal CustId As String) As String
        Dim aSQL As String = "UPDATE SO003 SET AccountNo=NULL,INVSEQNO = NULL, " &
                      "BankCode = NULL,BankName = NULL,CMCode = {0},CMName='{1}', " &
                      "PTCode = {2},PTName = '{3}',UpdEn = '{4}',UpdTime = '{5}', NewUpdTime = To_Date('{6}','yyyymmddhh24miss') " &
                      " WHERE 1=1 AND CUSTID = " & CustId & " AND BankCode = {7} " &
                      " And COMPCODE = {8} And ACCOUNTNO = '{9}' " & _
                      IIf(String.IsNullOrEmpty(SEQNO), "", " And SEQNO IN (" & SEQNO & " ) ")
        Return aSQL
    End Function
    Friend Function UpdOldNonePeriod(ByVal SeqNo As String, ByVal CustId As String) As String
        Dim aSQL As String = "UPDATE SO003 SET AccountNo='{0}', " &
                       "BankCode = {1},BankName = '{2}',CMCode = {3},CMName='{4}', " &
                       "PTCode = {5},PTName = '{6}',UpdEn = '{7}',UpdTime = '{8}',InvSeqNo = {9}, NewUpdTime = To_Date('{10}','yyyymmddhh24miss')  " &
                       " WHERE 1=1 AND CUSTID =" & CustId &
                       " And SEQNO IN (" & SeqNo & ") "
        Return aSQL
    End Function
    Friend Overridable Function UpdNewSO003() As String
        'Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                           " where a.seqno = {12} And a.memberid = b.memberid  " & _
        '                                           " and b.HomeId = c.HomeId "
        Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                   " where a.seqno = {12} And a.seqno = b.memberid  " &
                                                   " and b.HomeId = c.HomeId "
        Dim aSQL As String = "UPDATE SO003 SET AccountNo='{0}', " &
                       "BankCode = {1},BankName = '{2}',CMCode = {3},CMName='{4}', " &
                       "PTCode = {5},PTName = '{6}',UpdEn = '{7}',UpdTime = '{8}',NewUpdTime = To_Date('{9}','yyyymmddhh24miss')  " &
                       " WHERE FaciSeqNo  IN ( {10}) AND CUSTID In (" & aCustId & ") " &
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN ({11}))"
        Return aSQL
    End Function
    Friend Function UpdSO003() As String
        Dim aSQL As String = "UPDATE SO003 SET AccountNo='{0}', " &
                       "BankCode = {1},BankName = '{2}',CMCode = {3},CMName='{4}', " &
                       "PTCode = {5},PTName = '{6}' " &
                       " WHERE FaciSeqNo  IN ( {7}) AND CUSTID = {8} " & _
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN ({9}))"
        Return aSQL
    End Function
    Friend Overridable Function ClearNoneSO003(ByVal SEQNO As String) As String

        Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno = {7} And a.seqno = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
        Dim aSQL As String = "UPDATE SO003 SET AccountNo=NULL, " &
                       "BankCode = NULL,BankName = NULL,CMCode = {0},CMName='{1}', " &
                       "PTCode = {2},PTName = '{3}',UpdEn = '{4}',UpdTime = '{5}', NewUpdTime = To_Date('{6}','yyyymmddhh24miss') " &
                       " WHERE 1=1 AND CUSTID In (" & aCustId & " )  " &
                       " And SEQNO IN (" & SEQNO & " ) " &
                       " AND CitemCode In (Select CodeNo From CD019 Where ProductCode is Null ) "

        Return aSQL
    End Function
    Friend Overridable Function ClearNewSO003() As String
        'Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                         " where a.seqno = {9} And a.memberid = b.memberid  " & _
        '                                         " and b.HomeId = c.HomeId "

        Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                 " where a.seqno = {9} And a.seqno = b.memberid  " &
                                                 " and b.HomeId = c.HomeId "
        Dim aSQL As String = "UPDATE SO003 SET AccountNo=NULL, " &
                       "BankCode = NULL,BankName = NULL,CMCode = {0},CMName='{1}', " &
                       "PTCode = {2},PTName = '{3}',UpdEn = '{4}',UpdTime = '{5}', NewUpdTime = To_Date('{6}','yyyymmddhh24miss') " &
                       " WHERE FaciSeqNo  IN ( {7}) AND CUSTID In (" & aCustId & " )  " &
                       " AND CitemCode In (Select CodeNo From CD019 Where ProductCode  IN ({8})) "

        Return aSQL



    End Function
    Friend Function ClearSO003() As String
        Dim aSQL As String = "UPDATE SO003 SET AccountNo=NULL, " &
                       "BankCode = NULL,BankName = NULL,CMCode = {0},CMName='{1}', " &
                       "PTCode = {2},PTName = '{3}' " &
                       " WHERE FaciSeqNo  IN ( {4}) AND CUSTID = {5} " & _
                       " AND CitemCode In (Select CodeNo From CD019 Where ProductCode  IN ({6})) "

        Return aSQL



    End Function
    Friend Overridable Function ChkNewSameAcc(ByVal SEQNO As String) As String
        'Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                      " where a.seqno = " & SEQNO & " And a.memberid = b.memberid  " & _
        '                                      " and b.HomeId = c.HomeId "

        Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                             " where a.seqno = " & SEQNO & " And a.seqno = b.memberid  " &
                                             " and b.HomeId = c.HomeId "
        Dim aSQL As String = String.Format("SELECT COUNT(1) CNT FROM SO106 " &
           " WHERE ACCOUNTID={0}0 " &
               " AND COMPCODE ={0}1 " &
               " AND CUSTID IN (" & aCustId & " ) " &
               " AND STOPFLAG = 0 AND STOPDATE IS NULL " &
               " AND MASTERID <> {0}2 ", Sign)
        Return aSQL
    End Function
    Friend Function ChkSameAcc() As String
        Dim aSQL As String = String.Format("SELECT COUNT(1) CNT FROM SO106 " &
            " WHERE ACCOUNTID={0}0 " & _
                " AND COMPCODE ={0}1 " & _
                " AND CUSTID = {0}2" & _
                " AND STOPFLAG = 0 AND STOPDATE IS NULL " &
                " AND MASTERID <> {0}3 ", Sign)
        Return aSQL

    End Function
    Friend Function getInvSO138Seqno() As String
        Return "select sf_getsequenceno('S_SO138_InvSeqNo') from dual"
    End Function
    Friend Function chkSameSO106() As String
        Return String.Format("Select Nvl(Count(*),0) Cnt From SO106 Where " & _
                        " AccountID= {0}0 And " &
                        "CompCode= {0}1 And " & _
                        "CUSTID= {0}2 And " & _
                        "StopFlag=0 And StopDate Is Null" & _
                        " And Masterid <> {0}3", Sign)
    End Function
    Friend Function UpdSO002A() As String
        Dim aSQL As String = String.Format("UPDATE SO002A SET " & _
                            "BANKCODE={0}0" & _
                            ",BANKNAME={0}1" & _
                            ",ID={0}2 " & _
                            ",ACCOUNTNO={0}3" & _
                            ",CARDNAME={0}4" & _
                            ",CARDEXPDATE={0}5" & _
                            ",CVC2={0}6" & _
                            ",NOTE={0}7" & _
                            ",CITEMSTR={0}8" & _
                            ",CITEMSTR2={0}9" & _                            
                            ",STOPFLAG=0,STOPDATE=NULL" & _
                            ",ADDCITEMACCOUNT = {0}10 " & _
                            " WHERE CUSTID={0}11 " & _
                            " AND ACCOUNTNO={0}12" & _
                            " AND COMPCODE={0}13 ", Sign)
        Return aSQL
    End Function
    Friend Overridable Function GetSysDate() As String
        Return "select sysdate from dual"
    End Function
    Friend Function QuerySO003CitemBySeq() As String
        Dim result As String = String.Format("Select CitemCode From  SO003 Where CustId= {0}0 " & _
                            " And AccountNo= {0}1" & _
                            " And CompCode= {0}2" & _
                            " And StopFlag=0" & _
                            " And SeqNo= {0}3 " & _
                            " AND BANKCODE={0}4", Sign)
        Return result

    End Function
    Friend Function chkSO003(ByVal seqno As String) As String
        Return String.Format("Select Nvl(Count(*),0) Cnt From SO106  Where CustId= {0}0 " &
                    " And AccountID= {0}1 " &
                    " And CompCode= {0}2" &
                    " And StopFlag=0" & _
                    " And StopDate is Null" & _
                    " And Masterid <> {0}3 " &
                     " And Instr(','||Citemstr||',' ,','||Chr(39)||" & seqno & "||Chr(39)||',')>0", Sign)
                 
    End Function
    Friend Function StopSO003BySeq() As String
        Return String.Format("UPDATE  SO003 SET " & _
                                 "BANKCODE=NULL" & _
                                 ",BANKNAME=NULL" & _
                                 ",ACCOUNTNO=NULL" & _
                                 ",InvSeqNo=NULL" & _
                                 ",CMCODE= {0}0" & _
                                ",CMNAME= {0}1" & _
                                ",PTCODE= {0}2" & _
                                ",PTNAME= {0}3 " & _
                                " WHERE CUSTID= {0}4 " & _
                                " AND ACCOUNTNO= {0}5" & _
                                " AND COMPCODE= {0}6 " & _
                                " AND BANKCODE= {0}7 " & _
                                 " And SEQNO = {0}8 ", Sign)

    End Function
    Friend Function AddAuthorize() As String
        'Dim ret As String
        Return String.Format("INSERT INTO SO106A ( MasterRowID," & _
                          " ACHTNO,CitemCodeStr,CitemNameStr," & _
                          " UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                           " AuthorizeStatus,ACHDesc,MasterId,StopFlag,Stopdate )" & _
                           " Values ( {0}0,{0}1,{0}2,{0}3,{0}4,{0}5,{0}6,{0}7,{0}8,{0}9," & _
                           "{0}10,{0}11,{0}12,{0}13 )", Sign)


    End Function
    
    Friend Function DelWaitAuthorize(ByVal CitemCode As Object, ByVal CitemName As Object) As String
        Return String.Format("delete from SO106A " &
                                    " Where MasterId={0}0 " &
                                    " And ACHTNO = {0}1 " &
                                    " And ACHDesc = {0}2 " &
                                    " And CitemCodeStr  " & IIf(DBNull.Value.Equals(CitemCode), " Is NULL ", String.Format("='{0}'", CitemCode)) &
                                    " And CitemNameStr " & IIf(DBNull.Value.Equals(CitemName), " Is NULL ", String.Format("='{0}'", CitemName)) &
                                    " And AuthorizeStatus=4 ", Sign)
    End Function

    Friend Overridable Function UpdAuthorize() As String

        'Return String.Format("Update SO106A Set " & _
        '                        " CitemCodeStr={0}0, " & _
        '                        " CitemNameStr = {0}1," & _
        '                        " UpdEn = {0}2, " & _
        '                        " UpdTime = {0}3 " & _
        '                         " Where MasterId = {0}4 " & _
        '                            " And ACHTNO = {0}5 " & _
        '                            " And ACHDesc = {0}6 " & _
        '                            " AND CitemCodeStr " & IIf(DBNull.Value.Equals(CitemCode), " IS  NOT NULL ", String.Format("<>'{0}'", CitemCode)) & _
        '                            " And Nvl(StopFlag,0 ) = 0", Sign)
        Return String.Format("Update SO106A Set " &
                               " CitemCodeStr={0}0, " &
                               " CitemNameStr = {0}1," &
                               " UpdEn = {0}2, " &
                               " UpdTime = {0}3 " &
                                " Where ROWID = {0}4 ", Sign)

    End Function
    Friend Function InsSO002A() As String
        Dim aInsValue As String = "{0}0"
        For i As Int32 = 1 To 21
            aInsValue = aInsValue & ",{0}" & i
        Next
        Dim aSQL As String = "INSERT INTO SO002A " & _
                            "(CUSTID,COMPCODE,BANKCODE,BANKNAME,ID,ACCOUNTNO," & _
                            "CARDNAME,CARDEXPDATE,CHARGEADDRNO,CHARGEADDRESS," & _
                            "MAILADDRNO,MAILADDRESS,CVC2,NOTE,CHARGETITLE," & _
                            "INVNO,INVTITLE,INVADDRESS,INVOICETYPE," & _
                            "CITEMSTR,CITEMSTR2,ADDCITEMACCOUNT) VALUES (" & aInsValue & ")"
        aSQL = String.Format(aSQL, Sign)
        Return aSQL
    End Function

    Friend Function ChkSO002ACnt() As String
        Dim aSQL As String = String.Format("select count(1) from so002a " & _
                                         "where accountNo={0}0 and custid={0}1 " & _
                                         "and compcode={0}2", Sign)
        Return aSQL
    End Function
    Friend Function GetSO001() As String
        Dim aSQL As String = String.Format("SELECT * FROM SO001 WHERE COMPCODE={0}0 " & _
                                         " AND CUSTID={0}1", Sign)
        Return aSQL
    End Function
    Friend Overridable Function GetNewSO002() As String

        'Dim custId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                         " where a.seqno ={0}1 And a.memberid = b.memberid  " & _
        '                                         " and b.HomeId = c.HomeId "
        Dim custId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno ={0}1 And a.seqno = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
        Dim aSQL As String = String.Format("SELECT * FROM SO002 WHERE COMPCODE={0}0 " &
                                       " AND CUSTID IN ( " & custId & ") ORDER BY SERVICETYPE", Sign)
        Return aSQL
    End Function
    Friend Function GetSO002() As String
        Dim aSQL As String = String.Format("SELECT * FROM SO002 WHERE COMPCODE={0}0 " & _
                                         " AND CUSTID={0}1 ORDER BY SERVICETYPE", Sign)
        Return aSQL
    End Function
    Friend Function StopChildSO106() As String
        Dim aSQL As String = String.Format("UPDATE SO106 " & _
                                " SET INHERITFLAG=0,INHERITKEY=NULL" & _
                                " WHERE ACCOUNTID={0}0 " & _
                                " AND INHERITKEY={0}1 " & _
                                " AND INHERITFLAG=1" & _
                                " AND COMPCODE={0}2", Sign)
        Return aSQL
    End Function
    'Friend Function StopChildSO002A() As String
    '    Dim aSQL As String = "UPDATE  SO002A" & _
    '                                " SET INHERITFLAG=0,INHERITKEY=NULL" & _
    '                                " WHERE INHERITKEY IN (" & strChild2A & ")" & _
    '                                " AND INHERITFLAG=1" & _
    '                                " AND COMPCODE=" & gCompCode

    'End Function

    Friend Overridable Function StopSO002A(ByVal filterCustId As Boolean, ByVal SEQNO As String) As String
        Dim aSQL As String = String.Empty
        'Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                            " where a.seqno = " & SEQNO & " And a.memberid = b.memberid  " & _
        '                                            " and b.HomeId = c.HomeId "
        Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                   " where a.seqno = " & SEQNO & " And a.seqno = b.memberid  " &
                                                   " and b.HomeId = c.HomeId "
        If filterCustId Then
            aSQL = String.Format("UPDATE SO002A" &
                                    " SET STOPFLAG=1,STOPDATE={0}0" &
                                    " WHERE ACCOUNTNO={0}1 " &
                                    " AND CUSTID={0}2" &
                                    " AND COMPCODE={0}3", Sign)
        Else
            aSQL = String.Format("UPDATE SO002A" &
                                  " SET STOPFLAG=1,STOPDATE={0}0" &
                                  " WHERE ACCOUNTNO={0}1 " &
                                  " AND CUSTID IN (" & aCustId & " )" &
                                  " AND COMPCODE={0}2", Sign)
        End If
        Return aSQL
    End Function
    Friend Function UpdSO003C(ByVal strServiceId As String, ByVal strNewUpdTime As String) As String
        Dim aSQL As String = String.Format("UPDATE SO003C SET CMCode = {0}0,CMName={0}1, " &
                      "PTCode = {0}2,PTName = {0}3,MasterId = {0}4,UpdEn ={0}5,  " &
                      " UpdTime = {0}6 ,NewUpdTime = To_Date('" & strNewUpdTime & "','yyyymmddhh24miss') " &
                      " WHERE ServiceId IN (" & strServiceId & ") ", Sign)
        Return aSQL
    End Function
    Friend Function UpdSO003C() As String
        Dim aSQL As String = "UPDATE SO003C SET CMCode = {0},CMName='{1}', " &
                      "PTCode = {2},PTName = '{3}',MasterId = {4},UpdEn ='{5}',  " & _
                      " UpdTime = '{6}' ,NewUpdTime = To_Date('{7}','yyyymmddhh24miss') " & _
                      " WHERE ServiceId IN ( {8}) "
        Return aSQL
    End Function
    Friend Function UpdSO004() As String
        Dim aSQL As String = "UPDATE SO004 SET AccountNo='{0}', " &
                       "BankCode = {1},BankName = '{2}',ChCMCode = {3},ChCMName='{4}', " &
                       "ChPTCode = {5},ChPTName = '{6}',MasterId = {7}  WHERE SEQNO IN ( {8}) "
        Return aSQL
    End Function
    Friend Overridable Function GetNewSO003C(ByVal strServiceId As String) As String
        Return String.Format("SELECT *  FROM SO003C WHERE ServiceId IN (" & strServiceId & ") AND CUSTID In ( " &
                                                   "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                  " where a.seqno = {0}0 And a.seqno = b.memberid  " &
                                                  " and b.HomeId = c.HomeId )", Sign)

    End Function
    Friend Overridable Function GetNewSO003C() As String
        'Return "SELECT *  FROM SO003C WHERE ServiceId IN ({0}) AND CUSTID In ( " & _
        '                                            "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                           " where a.seqno = {1} And a.memberid = b.memberid  " & _
        '                                           " and b.HomeId = c.HomeId )"
        Return "SELECT *  FROM SO003C WHERE ServiceId IN ({0}) AND CUSTID In ( " &
                                                   "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                  " where a.seqno = {1} And a.seqno = b.memberid  " &
                                                  " and b.HomeId = c.HomeId )"
    End Function

    Friend Function GetSO003C() As String

        Return "SELECT *  FROM SO003C WHERE ServiceId IN ({0}) AND CUSTID = {1}"
    End Function
    Friend Function GetPTCode() As String
        Return "Select CodeNo,Description,RefNo From CD032 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
    End Function
    Friend Function GetCMCode() As String
        Return "Select CodeNo,Description,RefNo From CD031 Where Nvl(StopFlag,0) = 0 ORDER BY CodeNo"
    End Function
    Friend Function ChkAchSN() As String
        Return String.Format("SELECT COUNT(*) FROM SO106 WHERE ACHSN= {0}0 And MasterId <> {0}1", Sign)
    End Function
    Friend Function GetSystemPara() As String
        Return String.Format("Select Nvl(StartPost,0) From SO041 Where SysID = {0}0", Sign)
    End Function
    Friend Function GetProposer() As String
        Dim strQry As String = String.Empty
        strQry = String.Format("Select A.CustName From SO001 A " &
             " Where A.CustId={0}0" &
             " And A.CompCode={0}1" &
             " Union All " &
             "Select B.DeclarantName CustName From SO004 B " &
             " Where B.CustId={0}2" &
             " And B.CompCode={0}3" &
             " And (B.PRDate Is Null OR B.InstDate > B.PRDate)", Sign)
        strQry = "Select distinct * from (" & strQry & ")"
        Return strQry
    End Function
    Friend Function GetNewProposer() As String
        Dim strQry As String = String.Empty
        strQry = String.Format("Select DeclarantName From SO137 " & _
            " Where SEQNO = {0}0 ", Sign)
        Return strQry
    End Function
    Friend Function GetSO137() As String
        Dim strQry = String.Format("Select * From SO137 Where SEQNO = {0}0", Sign)
        Return strQry
    End Function
    Friend Function GetBankCodeByCode(ByVal blnStartPost As Boolean) As String
        Dim aSQL As String = Nothing
        If blnStartPost Then
            aSQL = String.Format("Select CodeNo,Description,RefNo,ActLength,PRGNAME, " & _
                "(case when prgname like'ACH%'  then 1 " & _
                                          " when prgname like'%POST4%' then 2 " & _
                                          " else 0 " & _
                                         " end ) ACHTYPE " & _
            " From CD018 Where Nvl(StopFlag,0) = 0 AND CodeNo = {0}0 ORDER BY CODENO", Sign)
        Else
            aSQL = String.Format("Select CodeNo,Description,RefNo,ActLength,PRGNAME, " & _
                "(case when prgname like'ACH%'  then 1 " & _
                                          " else 0 " & _
                                         " end ) ACHTYPE " & _
            " From CD018 Where Nvl(StopFlag,0) = 0 AND CodeNo = {0}0 ORDER BY CODENO", Sign)
        End If

        Return aSQL
    End Function
    Friend Function GetBankCode(ByVal blnStartPost As Boolean) As String
        Dim aSQL As String = Nothing
        If blnStartPost Then
            aSQL = "Select CodeNo,Description,RefNo,ActLength,PRGNAME, " & _
                "(case when prgname like'ACH%'  then 1 " & _
                                          " when prgname like'%POST4%' then 2 " & _
                                          " else 0 " & _
                                         " end ) ACHTYPE " & _
            " From CD018 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
        Else
            aSQL = "Select CodeNo,Description,RefNo,ActLength,PRGNAME, " & _
                "(case when prgname like'ACH%'  then 1 " & _
                                          " else 0 " & _
                                         " end ) ACHTYPE " & _
            " From CD018 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
        End If

        Return aSQL
        'Return "Select CodeNo,Description,RefNo,ActLength,PRGNAME From CD018 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
    End Function
    Friend Function GetBankCode() As String
        Return "Select CodeNo,Description,RefNo,ActLength,PRGNAME From CD018 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
    End Function
    Friend Function GetCardCode() As String
        Return "Select CodeNo,Description,RefNo From CD037 Where Nvl(StopFlag,0) = 0  ORDER BY CODENO"
    End Function
    Friend Function GetCardCodeByCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD037 Where Nvl(StopFlag,0) = 0  And CodeNo ={0}0 ORDER BY CODENO", Sign)
    End Function
    Friend Function GetMediaCode() As String
        Return "Select CodeNo,Description,RefNo From CD009 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
    End Function
    Friend Function GetAcceptName() As String
        Return String.Format("SELECT EmpNo,EmpName FROM CM003 WHERE COMPCODE={0}0", Sign)
    End Function
    'Friend Function GetIntroId(ByVal MediaRefNo As Integer) As String
    '    Dim aRet As String = String.Empty
    '    Select Case MediaRefNo
    '        Case 1
    '            aRet = String.Format("Select CustId as CodeNo ,CustName as Description From SO001 " &
    '                                 " Where CustId = {0}0", Sign)
    '        Case 2
    '            aRet = "Select EmpNo As CodeNo, EmpName As Description From CM003 " &
    '                " Where Nvl(StopFlag,0) = 0"
    '        Case 3
    '            aRet = "Select NameP As Description, IntroID As CodeNo FROM SO013"
    '    End Select
    '    Return aRet
    'End Function
    'Friend Function GetIntroData(ByVal MediaRefNo As Integer) As String
    '    Dim aRet As String = String.Empty
    '    Select Case MediaRefNo
    '        Case 1
    '            aRet = "Select CustId as CodeNo ,CustName as Description From SO001 "

    '        Case 2
    '            aRet = "Select EmpNo As CodeNo, EmpName As Description From CM003 " &
    '                " Where Nvl(StopFlag,0) = 0"
    '        Case 3
    '            aRet = "Select NameP As Description, IntroID As CodeNo FROM SO013 "
    '        Case Else
    '            aRet = "Select CustId as CodeNo ,CustName as Description From SO001 "
    '    End Select
    '    Return aRet
    'End Function
    Friend Overridable Function GetNewCanChooseCharge() As String
        Dim result As String = String.Format(
       " Select  A.CUSTID,DECODE(A.FACISNO,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " &
           " A.BILLNO,A.ITEM,A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME, " &
           " A.REALSTARTDATE,A.REALSTOPDATE,A.FACISNO,A.billno||A.item PKBILLNO " &
           " From so033 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO)  " &
           " LEFT JOIN SO137 C ON (B.ID=C.ID) " &
           " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " &
           " Where a.custid in (SELECT DISTINCT c.custid " &
            " FROM so137 a, so002c b, so001 c " &
           " WHERE(a.seqno = {0}0) " &
           " AND A.seqno = b.memberid " &
           " AND b.HomeId = c.HomeId) " &
          " AND A.UCCode NOT IN (SELECT CodeNo " &
          " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3,4,7, 8)) " &
           " AND A.UCCode IS NOT NULL ", Sign)

        Return result
    End Function
    Friend Overridable Function GetNewCanChooseCharge(ByVal SEQNO As String) As String
        'Dim result As String =
        '   " Select  A.CUSTID,DECODE(A.FACISNO,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " & _
        '       " A.BILLNO,A.ITEM,A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME, " & _
        '       " A.REALSTARTDATE,A.REALSTOPDATE,A.FACISNO,A.billno||A.item PKBILLNO " & _
        '       " From so033 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO)  " & _
        '       " LEFT JOIN SO137 C ON (B.ID=C.ID) " & _
        '       " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " & _
        '       " Where a.custid in (SELECT DISTINCT c.custid " & _
        '        " FROM so137 a, so002c b, so001 c " & _
        '       " WHERE(a.seqno = " & SEQNO & ") " & _
        '       " AND A.memberid = b.memberid " & _
        '       " AND b.HomeId = c.HomeId) " & _
        '      " AND A.UCCode NOT IN (SELECT CodeNo " & _
        '      " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " & _
        '       " AND A.UCCode IS NOT NULL "
        Dim result As String =
        " Select  A.CUSTID,DECODE(A.FACISNO,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " &
            " A.BILLNO,A.ITEM,A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME, " &
            " A.REALSTARTDATE,A.REALSTOPDATE,A.FACISNO,A.billno||A.item PKBILLNO " &
            " From so033 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO)  " &
            " LEFT JOIN SO137 C ON (B.ID=C.ID) " &
            " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " &
            " Where a.custid in (SELECT DISTINCT c.custid " &
             " FROM so137 a, so002c b, so001 c " &
            " WHERE(a.seqno = " & SEQNO & ") " &
            " AND A.seqno = b.memberid " &
            " AND b.HomeId = c.HomeId) " &
           " AND A.UCCode NOT IN (SELECT CodeNo " &
           " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " &
            " AND A.UCCode IS NOT NULL "

        Return result


    End Function
    Friend Function getCD068A(ByVal SeqNo As String) As String
        Dim result As String = Nothing
        result = String.Format("Select A.CitemCode,B.Description CITEMNAME From CD068A A,CD019 B " & _
            " Where A.CitemCode = B.CodeNo And 1=1 " & _
            " And A.BillHeadFmt = {0}0 And A.CitemCode in " & _
            "(Select CitemCode From SO003 Where CustId = {0}1 And SEQNO IN (" & SeqNo & " ) " & _
               " And CompCode = {0}2 )", Sign)
        Return result
    End Function
    Friend Function GetCanChooseCharge() As String
        Return String.Format("Select A.BILLNO || A.ITEM PKBILLNO,A.CITEMNAME DESCRIPTION, " & _
                             " A.CITEMCODE,A.CITEMNAME, " & _
                             "A.BILLNO,A.ITEM,A.REALPERIOD,A.SHOULDAMT,A.CMNAME, " & _
                             "Nvl(A.AccountNo,' ') AccountNo, " & _
                             " to_char(A.REALSTARTDATE,'yyyy/mm/dd') REALSTARTDATE , " & _
                             " to_char(A.REALSTOPDATE,'yyyy/mm/dd' ) REALSTOPDATE  From SO033 A  " & _
            " Where A.CustId = {0}0 " & _
            " And A.UCCode NOT IN (Select CodeNo " & _
            "           From CD013 Where PayOk = 1 Or RefNo in (3,7,8))  " & _
            " AND A.UCCode IS NOT NULL " & _
            " Order By A.BillNo,A.Item", Sign)

    End Function
    Friend Function GetACHTNo(ByVal blnStartPost As Boolean) As String
        Dim ACHType As String = "1"
        If blnStartPost Then ACHType = "1,2"
        Return "SELECT ACHTNO,ACHTDESC,ACHTYPE,BillHeadFmt FROM CD068 " &
            " Where ACHTNO Is Not NULL " &
            " And ACHTDESC is Not NULL " &
            " And ACHType in (" & ACHType & ") " &
            " Group By ACHTNO,ACHTDESC,ACHType,BillHeadFmt Order By ACHTNO"

    End Function
    Friend Function GetACHTNo() As String
        Return "SELECT ACHTNO,ACHTDESC FROM CD068 " &
            " Where ACHTNO Is Not NULL " &
            " And ACHTDESC is Not NULL " &
            " And ACHType in (1) " &
            " Group By ACHTNO,ACHTDESC"

    End Function
    Friend Function GetDataServiceType() As String
        Return String.Format("SELECT SERVICETYPE FROM SO033 WHERE BILLNO = {0}0 AND ITEM = {0}1", Sign)
    End Function
    
    Friend Function GetVirtualAccountQry() As String
        Dim aRet As String = String.Format("Select Nvl(To_number(Max(SubStr(Accountid,1,8))),0) + 1 As Cnt From SO106 " &
                                           " Where CMCode Not In (Select CodeNo From CD031 Where RefNo in (2,4)) " &
                                           " And CustId = {0}0", Sign)
        Return aRet

    End Function
    Friend Function GetOldVirtualAccountQry() As String
        'Dim aRet As String = String.Format("SELECT COUNT(*) FROM SO002A WHERE CUSTID= {0}0 " & _
        '                                  " AND ID=2", Sign)
        Dim aRet As String = String.Format("Select Nvl(Max(SubStr(AccountNO,1,8)),0) + 1   " & _
                                            " FROM SO002A WHERE CUSTID={0}0 AND ID=2", Sign)
        Return aRet

    End Function
    Friend Overridable Function GetCanChooseBillNo(ByVal SeqNo As String) As String
        'Dim result As String =
        '  " Select  A.CUSTID,DECODE(C.DECLARANTNAME,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " & _
        '      " A.BILLNO,A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME, " & _
        '      " A.REALSTARTDATE,A.REALSTOPDATE,A.FACISNO,A.billno||A.item billPK " & _
        '      " From so033 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO)  " & _
        '      " LEFT JOIN SO137 C ON (B.ID=C.ID) " & _
        '      " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " & _
        '      " Where a.custid in (SELECT DISTINCT c.custid " & _
        '       " FROM so137 a, so002c b, so001 c " & _
        '      " WHERE(a.seqno = " & SeqNo & ") " & _
        '      " AND a.memberid = b.memberid " & _
        '      " AND b.HomeId = c.HomeId) " & _
        '     " AND A.UCCode NOT IN (SELECT CodeNo " & _
        '     " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " & _
        '      " AND A.UCCode IS NOT NULL "

        Dim result As String =
        " Select  A.CUSTID,DECODE(C.DECLARANTNAME,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " &
            " A.BILLNO,A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME, " &
            " A.REALSTARTDATE,A.REALSTOPDATE,A.FACISNO,A.billno||A.item billPK " &
            " From so033 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO)  " &
            " LEFT JOIN SO137 C ON (B.ID=C.ID) " &
            " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " &
            " Where a.custid in (SELECT DISTINCT c.custid " &
             " FROM so137 a, so002c b, so001 c " &
            " WHERE(a.seqno = " & SeqNo & ") " &
            " AND a.seqno = b.memberid " &
            " AND b.HomeId = c.HomeId) " &
           " AND A.UCCode NOT IN (SELECT CodeNo " &
           " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " &
            " AND A.UCCode IS NOT NULL "
        Return result
    End Function
    Friend Overridable Function GetNewCanChooseBillNo(ByVal ACHTNO As String, ByVal ACHTDESC As String) As String
        Dim BillHeadFmt As String = "Select BillHeadFmt From CD068 Where ACHTNO IN(" & ACHTNO & ") " &
                                    " And ACHTDESC IN (" & ACHTDESC & ")"
        Dim result As String = String.Format(
        " Select  A.CUSTID,DECODE(C.DECLARANTNAME,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " &
            " A.BILLNO,A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME, " &
            " A.REALSTARTDATE,A.REALSTOPDATE,A.FACISNO,A.billno||A.item billPK " &
            " From so033 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO)  " &
            " LEFT JOIN SO137 C ON (B.ID=C.ID) " &
            " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " &
            " Where a.custid in (SELECT DISTINCT c.custid " &
             " FROM so137 a, so002c b, so001 c " &
            " WHERE (a.seqno = {0}0) " &
            " AND a.seqno = b.memberid " &
            " AND b.HomeId = c.HomeId) " &
           " AND A.UCCode NOT IN (SELECT CodeNo " &
           " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " &
           " AND A.CitemCode IN (Select CitemCode From CD068A  Where CD068A.BillHeadFmt IN (" & BillHeadFmt & ")) " &
            " AND A.UCCode IS NOT NULL ", Sign)
        Return result
    End Function
    Friend Overridable Function GetNewCanChooseNonePeriodWithACH(ByVal ACHTNO As String, ByVal ACHTDESC As String) As String
        Dim BillHeadFmt As String = "Select BillHeadFmt From CD068 Where ACHTNO IN(" & ACHTNO & ") " &
                                    " And ACHTDESC IN (" & ACHTDESC & ")"
        Dim result As String = String.Format(
          "Select A.CUSTID,DECODE(C.DECLARANTNAME,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " &
              " A.CITEMCODE,A.CITEMNAME,A.STOPFLAG,A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME, " &
              " A.STARTDATE,A.STOPDATE,A.FACISNO,A.SeqNo " &
               " From so003 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO) " &
                " LEFT JOIN SO137 C ON (B.ID=C.ID) " &
                " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " &
                " JOIN CD019 E ON (A.CITEMCODE=E.CODENO) " &
                " Where a.custid in (SELECT DISTINCT c.custid " &
                " FROM so137 a, so002c b, so001 c " &
                " WHERE (a.seqno = {0}0) " &
                 " And a.seqno = b.memberid " &
                  "  AND b.HomeId = c.HomeId) " &
                    " AND A.CITEMCODE=E.CODENO " &
                    " AND A.CitemCode IN (Select CitemCode From CD068A  Where CD068A.BillHeadFmt IN (" & BillHeadFmt & ")) " &
                    " AND E.PRODUCTCODE IS NULL ", Sign)


        Return result
    End Function
    Friend Overridable Function GetCanChooseNonePeriod(ByVal SeqNo As String) As String

        'Dim result As String =
        '    "Select A.CUSTID,DECODE(C.DECLARANTNAME,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " & _
        '        " A.CITEMCODE,A.CITEMNAME,A.STOPFLAG,A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME, " & _
        '        " A.STARTDATE,A.STOPDATE,A.FACISNO,A.SeqNo " & _
        '         " From so003 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO) " & _
        '          " LEFT JOIN SO137 C ON (B.ID=C.ID) " & _
        '          " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " & _
        '          " JOIN CD019 E ON (A.CITEMCODE=E.CODENO) " & _
        '          " Where a.custid in (SELECT DISTINCT c.custid " & _
        '          " FROM so137 a, so002c b, so001 c " & _
        '          " WHERE (a.seqno = " & SeqNo & ") " & _
        '           " And a.memberid = b.memberid " & _
        '            "  AND b.HomeId = c.HomeId) " & _
        '              " AND A.CITEMCODE=E.CODENO " & _
        '              " AND E.PRODUCTCODE IS NULL "


        Dim result As String =
           "Select A.CUSTID,DECODE(C.DECLARANTNAME,NULL,D.CUSTNAME,C.DECLARANTNAME) DECLARANTNAME, " &
               " A.CITEMCODE,A.CITEMNAME,A.STOPFLAG,A.PERIOD,A.AMOUNT,A.ACCOUNTNO,A.CMNAME, " &
               " A.STARTDATE,A.STOPDATE,A.FACISNO,A.SeqNo " &
                " From so003 A LEFT JOIN SO004 B ON (A.CUSTID=B.CUSTID AND FACISEQNO=B.SEQNO) " &
                 " LEFT JOIN SO137 C ON (B.ID=C.ID) " &
                 " JOIN SO001 D ON (A.CUSTID=D.CUSTID) " &
                 " JOIN CD019 E ON (A.CITEMCODE=E.CODENO) " &
                 " Where a.custid in (SELECT DISTINCT c.custid " &
                 " FROM so137 a, so002c b, so001 c " &
                 " WHERE (a.seqno = " & Integer.Parse(SeqNo) & ") " &
                  " And a.seqno = b.memberid " &
                   "  AND b.HomeId = c.HomeId) " &
                     " AND A.CITEMCODE=E.CODENO " &
                     " AND E.PRODUCTCODE IS NULL "


        Return result
    End Function
    Friend Overridable Function GetNewCanChooseProdutWithACH(ByVal ACHTNO As String, ByVal ACHTDESC As String) As String
        Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                 " where a.seqno = {0}0 And a.seqno = b.memberid  " &
                                                 " and b.HomeId = c.HomeId "
        Dim BillHeadFmt As String = "Select BillHeadFmt From CD068 Where ACHTNO IN(" & ACHTNO & ") " &
                                    " And ACHTDESC IN (" & ACHTDESC & ")"

        Dim aRet As String = String.Format("Select A.ServiceId,A.ProductName,C.FaciSNo," &
                                         "B.ACHTNO,D.CUSTID,D.HOMEID,D.InstAddress, " &
                                         " Nvl( (select SO137.DeclarantName from so137 where C.ID = SO137.ID),D.CUSTNAME) DeclarantName" &
                                         " FROM SO003C A left join SO004 C on  A.FACISEQNO = C.SEQNO, CD046 B,SO001 D  " &
                                       " Where A.ServiceType = B.CodeNo " &
                                       " And A.CustId = D.CustId " &
                                       " And A.CustId  In (" & aCustId & ") " &
                                       " And (( A.InstDate is null ) Or ( A.PRdate is null ) Or ( A.InstDate > A.PRDate))" &
                                       " And A.ServiceId Is Not Null " &
                                       " AND A.CitemCode IN (Select CitemCode From CD068A  Where CD068A.BillHeadFmt IN (" & BillHeadFmt & ")) " &
                                       " Order By D.CUSTID,A.ServiceId", Sign)

        Return aRet
    End Function
    Friend Overridable Function GetNewCanChooseProduct() As String


        Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                  " where a.seqno = {0}0 And a.seqno = b.memberid  " &
                                                  " and b.HomeId = c.HomeId "

        Dim aRet As String = String.Format("Select A.ServiceId,A.ProductName,C.FaciSNo," &
                                         "B.ACHTNO,D.CUSTID,D.HOMEID,D.InstAddress, " &
                                         " Nvl( (select SO137.DeclarantName from so137 where C.ID = SO137.ID),D.CUSTNAME) DeclarantName" &
                                         " FROM SO003C A LEFT JOIN  SO004 C ON A.FACISEQNO = C.SEQNO ,CD046 B,SO001 D  " &
                                       " Where A.ServiceType = B.CodeNo " &
                                       " And A.CustId = D.CustId " &
                                       " And A.CustId  In (" & aCustId & ") " &
                                       " And (( A.InstDate is null ) Or ( A.PRdate is null ) Or ( A.InstDate > A.PRDate))" &
                                       " And A.ServiceId Is Not Null " &
                                       " Order By D.CUSTID,A.ServiceId", Sign)

        Return aRet
    End Function
    Friend Overridable Function GetNewCanChooseProduct(ByVal SeqNo As String) As String


        Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                  " where a.seqno = " & SeqNo & " And a.seqno = b.memberid  " &
                                                  " and b.HomeId = c.HomeId "
        'Dim aRet As String = String.Format("Select A.ServiceId,A.ProductName,C.FaciSNo," & _
        '                                 "B.ACHTNO,D.CUSTID,D.HOMEID,D.InstAddress, " & _
        '                                 " Nvl( (select SO137.DeclarantName from so137 where C.ID = SO137.ID),D.CUSTNAME) DeclarantName" & _
        '                                 " FROM SO003C A,CD046 B,SO004 C,SO001 D  " & _
        '                               " Where A.ServiceType = B.CodeNo " & _
        '                               " And A.CustId = D.CustId " & _
        '                               " And A.CustId  In (" & aCustId & ") " & _
        '                               " AND A.FACISEQNO = C.SEQNO(+) " & _
        '                               " And (( A.InstDate is null ) Or ( A.PRdate is null ) Or ( A.InstDate > A.PRDate))" & _
        '                               " And A.ServiceId Is Not Null " & _
        '                               " Order By D.CUSTID,A.ServiceId", Sign)
        Dim aRet As String = String.Format("Select A.ServiceId,A.ProductName,C.FaciSNo," &
                                         "B.ACHTNO,D.CUSTID,D.HOMEID,D.InstAddress, " &
                                         " Nvl( (select SO137.DeclarantName from so137 where C.ID = SO137.ID),D.CUSTNAME) DeclarantName" &
                                         " FROM SO003C A LEFT JOIN SO004 C ON A.FACISEQNO = C.SEQNO ,CD046 B,SO001 D  " &
                                       " Where A.ServiceType = B.CodeNo " &
                                       " And A.CustId = D.CustId " &
                                       " And A.CustId  In (" & aCustId & ") " &
                                       " And (( A.InstDate is null ) Or ( A.PRdate is null ) Or ( A.InstDate > A.PRDate))" &
                                       " And A.ServiceId Is Not Null " &
                                       " Order By D.CUSTID,A.ServiceId", Sign)

        Return aRet
    End Function
    Friend Function GetCanChooseProduct() As String
        Dim aRet As String = String.Format("Select A.ServiceId,A.ProductName,C.FaciSNo," & _
                                           "B.ACHTNO FROM SO003C A,CD046 B,SO004 C " & _
                                         " Where A.ServiceType = B.CodeNo " & _
                                         " And A.CustId = {0}0 " & _
                                         " AND A.FACISEQNO = C.SEQNO " & _
                                         " And (( A.InstDate is null ) Or ( A.PRdate is null ) Or ( A.InstDate > A.PRDate))" & _
                                         " And A.ServiceId Is Not Null " & _
                                         " Order By A.ServiceId ", Sign)
        
        Return aRet
    End Function
    Friend Function GetCanChooseFaci() As String
        'Return String.Format("Select * From SO004 " &
        '                     " Where CustId = {0}0 " &
        '                     " And PRdate is null " &
        '                     " And GetDate is null", Sign)

        Dim aRet As String = String.Empty


        aRet = String.Format("SELECT A.*,B.ACHTNO FROM SO004 A, CD046 B " &
                " WHERE A.CUSTID ={0}0 And A.PRdate is null And A.GetDate is null " &
                " AND A.SERVICETYPE = B.CODENO " &
                " And A.FaciCode In (Select CodeNo From CD022 Where RefNo in (2,3,5,6,7,8,10))", Sign)

        Return aRet

    End Function
    Friend Function GetDefPTCode() As String
        Return "Select CodeNo,Description From CD032 Where CodeNo = 1"
    End Function
    Friend Function GetDefCMCode(ByVal LoginInfo As LoginInfo) As String
        Dim aOBJ As New CableSoft.SO.BLL.Utility.Charge(LoginInfo)
        Try
            Dim aCMCode As String = aOBJ.GetDefaultCMCode(String.Empty).ToString
            Return "SELECT " & aCMCode & " CODENO , Description FROM CD031 WHERE CODENO = " & aCMCode
        Finally
            If aOBJ IsNot Nothing Then
                aOBJ.Dispose()
                aOBJ = Nothing
            End If

        End Try

    End Function
    Friend Overridable Function QuerySO033() As String
        Return String.Format("SELECT BILLNO,ServiceType FROM  SO033 " & _
                            " WHERE CUSTID= {0}0 " &
                            " AND ACCOUNTNO= {0}1 " & _
                            " AND COMPCODE= {0}2" & _
                             " AND BANKCODE= {0}3 " & _
                            " AND UCCODE > 0 AND CANCELFLAG=0" & _
                            " AND CitemCode= {0}4 " & _
                            " AND ROWNUM=1", Sign)
    End Function
    Friend Function stopOldSO033() As String
        Return String.Format("UPDATE SO033 SET " & _
                                   "BANKCODE=NULL" & _
                                   ",BANKNAME=NULL" & _
                                   ",ACCOUNTNO=NULL" & _
                                   ",InvSeqNo=NULL" & _
                                   ",CMCODE= {0}0 " & _
                                   ",CMNAME= {0}1 " & _
                                   ",PTCODE= {0}2 " & _
                                   ",PTNAME= {0}3 " & _
                                   " WHERE CUSTID= {0}4 " & _
                                   " AND ACCOUNTNO= {0}5 " & _
                                   " AND COMPCODE= {0}6 " & _
                                    " AND BANKCODE= {0}7 " & _
                                    " AND UCCODE > 0 AND CANCELFLAG=0" & _
                                    " AND CitemCode= {0}8", Sign)
    End Function
    Friend Function GetDefCMCode(ByVal LoginInfo As LoginInfo, ByVal ServiceType As String) As String
        Dim aOBJ As New CableSoft.SO.BLL.Utility.Charge(LoginInfo)
        Try

            Dim aCMCode As String = aOBJ.GetDefaultCMCode(ServiceType).ToString
            Return "SELECT " & aCMCode & " CODENO , Description FROM CD031 WHERE CODENO = " & aCMCode
        Finally
            If aOBJ IsNot Nothing Then
                aOBJ.Dispose()
                aOBJ = Nothing
            End If

        End Try

    End Function
    Friend Function GetCMRefNo() As String
        Return String.Format(" Select Nvl(RefNo,0) RefNo From CD031 " &
                                                           " Where CodeNo = {0}0", Sign)
    End Function
    Friend Function GetCardNoLen() As String
        Return String.Format("SELECT NVL(CARDNOLEN,0) CARDNOLEN FROM CD037 " &
                                                " WHERE CODENO= {0}0", Sign)
    End Function

    Friend Function GetActLength() As String
        Return String.Format("SELECT NVL(ActLength,0) ActLength FROM CD018 WHERE CODENO= {0}0", Sign)
    End Function
    Friend Function GetCardRefNo() As String
        Return String.Format("Select Nvl(RefNo,0) RefNo From CD037 Where CodeNo = {0}0 ", Sign)
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
