Public Class AccountDALMultiDB
    Inherits AccountDAL
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function TakeSO106SeqNo() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT sf_getsequenceno('S_SO106_MasterId')"
            Case Else
                Return MyBase.TakeSO106SeqNo
        End Select
    End Function
    Friend Overrides Function QueryAccountDetail() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("select a.ctid::text,a.* from so106a a where masterid = {0}0", Sign)
            Case Else
                Return MyBase.QueryAccountDetail
        End Select

    End Function
    Friend Overrides Function QuerySO106Log() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select A.ctid::text,A.* From SO106  A where Masterid = {0}0", Sign)
            Case Else
                Return MyBase.QuerySO106Log
        End Select

    End Function
    Friend Overrides Function GetSO137CustId() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select * from SO001 Where CustId In  (Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno ={0}0 And a.seqno ::text = b.memberid  " &
                                                " and b.HomeId = c.HomeId )", Sign)
            Case Else
                Return MyBase.GetSO137CustId
        End Select


    End Function
    Friend Overrides Function GetNewCitemCode(ByVal FaciSeqNos As String, ByVal ProductCodes As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno = {0}0 And a.seqno ::text = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
                Return String.Format("Select SeqNo,CitemCode,CitemName From SO003 " &
               " Where CustId In (" & aCustId & ") And FaciSeqNo IN (" & FaciSeqNos & ") And NVL(StopFlag,0) = 0 " &
               " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN (" & ProductCodes & "))",
               Sign)
            Case Else
                Return MyBase.GetNewCitemCode(FaciSeqNos, ProductCodes)
        End Select



    End Function
    Friend Overrides Function UpdNewNonePeriod(ByVal SeqNo As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                 " where a.seqno = {10} And a.seqno :: text = b.memberid  " &
                                                 " and b.HomeId = c.HomeId "
                Dim aSQL As String = "UPDATE SO003 SET AccountNo='{0}', " &
                       "BankCode = {1},BankName = '{2}',CMCode = {3},CMName='{4}', " &
                       "PTCode = {5},PTName = '{6}',UpdEn = '{7}',UpdTime = '{8}',NewUpdTime = To_Date('{9}','yyyymmddhh24miss')  " &
                       " WHERE 1=1 AND CUSTID In (" & aCustId & ") " &
                       " And SEQNO IN (" & SeqNo & ") " &
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode is Null)"
                Return aSQL
            Case Else
                Return MyBase.UpdNewNonePeriod(SeqNo)
        End Select



    End Function

    Friend Overrides Function GetNewCitemCode(ByVal SeqNo As String, ByVal FaciSeqNos As String, ByVal ProductCodes As String) As String
        'Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
        '                                          " where a.seqno = " & SeqNo & " And a.memberid = b.memberid  " & _
        '                                          " and b.HomeId = c.HomeId "
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno = " & SeqNo & " And a.seqno :: text = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
                Return String.Format("Select SeqNo,CitemCode,CitemName From SO003 " &
               " Where CustId In ({0}) And FaciSeqNo IN ({1}) And NVL(StopFlag,0) = 0 " &
               " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN ({2}))",
               aCustId, FaciSeqNos, ProductCodes)
            Case Else
                Return MyBase.GetNewCitemCode(SeqNo, FaciSeqNos, ProductCodes)
        End Select


    End Function
    Friend Overrides Function UpdNewSO003() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                   " where a.seqno = {12} And a.seqno :: text = b.memberid  " &
                                                   " and b.HomeId = c.HomeId "
                Dim aSQL As String = "UPDATE SO003 SET AccountNo='{0}', " &
                       "BankCode = {1},BankName = '{2}',CMCode = {3},CMName='{4}', " &
                       "PTCode = {5},PTName = '{6}',UpdEn = '{7}',UpdTime = '{8}',NewUpdTime = To_Date('{9}','yyyymmddhh24miss')  " &
                       " WHERE FaciSeqNo  IN ( {10}) AND CUSTID In (" & aCustId & ") " &
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode IN ({11}))"
                Return aSQL
            Case Else
                Return MyBase.UpdNewSO003
        End Select

    End Function

    Friend Overrides Function ClearNoneSO003(ByVal SEQNO As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno = {7} And a.seqno :: text = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
                Dim aSQL As String = "UPDATE SO003 SET AccountNo=NULL, " &
                       "BankCode = NULL,BankName = NULL,CMCode = {0},CMName='{1}', " &
                       "PTCode = {2},PTName = '{3}',UpdEn = '{4}',UpdTime = '{5}', NewUpdTime = To_Date('{6}','yyyymmddhh24miss') " &
                       " WHERE 1=1 AND CUSTID In (" & aCustId & " )  " &
                       " And SEQNO IN (" & SEQNO & " ) " &
                       " AND CitemCode In (Select CodeNo From CD019 Where ProductCode is Null ) "

                Return aSQL
            Case Else
                Return MyBase.ClearNoneSO003(SEQNO)
        End Select

    End Function

    Friend Overrides Function ClearNewSO003() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                 " where a.seqno = {9} And a.seqno::text = b.memberid  " &
                                                 " and b.HomeId = c.HomeId "
                Dim aSQL As String = "UPDATE SO003 SET AccountNo=NULL, " &
                       "BankCode = NULL,BankName = NULL,CMCode = {0},CMName='{1}', " &
                       "PTCode = {2},PTName = '{3}',UpdEn = '{4}',UpdTime = '{5}', NewUpdTime = To_Date('{6}','yyyymmddhh24miss') " &
                       " WHERE FaciSeqNo  IN ( {7}) AND CUSTID In (" & aCustId & " )  " &
                       " AND CitemCode In (Select CodeNo From CD019 Where ProductCode  IN ({8})) "

                Return aSQL
            Case Else
                Return MyBase.ClearNewSO003()
        End Select




    End Function
    Friend Overrides Function ChkNewSameAcc(ByVal SEQNO As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                             " where a.seqno = " & SEQNO & " And a.seqno ::text = b.memberid  " &
                                             " and b.HomeId = c.HomeId "
                Dim aSQL As String = String.Format("SELECT COUNT(1) CNT FROM SO106 " &
           " WHERE ACCOUNTID={0}0 " &
               " AND COMPCODE ={0}1 " &
               " AND CUSTID IN (" & aCustId & " ) " &
               " AND STOPFLAG = 0 AND STOPDATE IS NULL " &
               " AND MASTERID <> {0}2 ", Sign)
                Return aSQL
            Case Else
                Return MyBase.ChkNewSameAcc(SEQNO)
        End Select



    End Function
    Friend Overrides Function GetNewSO002() As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim custId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                " where a.seqno ={0}1 And a.seqno ::text = b.memberid  " &
                                                " and b.HomeId = c.HomeId "
                Dim aSQL As String = String.Format("SELECT * FROM SO002 WHERE COMPCODE={0}0 " &
                                       " AND CUSTID IN ( " & custId & ") ORDER BY SERVICETYPE", Sign)
                Return aSQL
            Case Else
                Return MyBase.GetNewSO002
        End Select

    End Function
    Friend Overrides Function GetNewCanChooseBillNo(ByVal ACHTNO As String, ByVal ACHTDESC As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
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
            " AND a.seqno::text = b.memberid " &
            " AND b.HomeId = c.HomeId) " &
           " AND A.UCCode NOT IN (SELECT CodeNo " &
           " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " &
           " AND A.CitemCode IN (Select CitemCode From CD068A  Where CD068A.BillHeadFmt IN (" & BillHeadFmt & ")) " &
            " AND A.UCCode IS NOT NULL ", Sign)
                Return result
            Case Else
                Return MyBase.GetNewCanChooseBillNo(ACHTNO, ACHTDESC)
        End Select

    End Function
    Friend Overrides Function GetNewCanChooseNonePeriodWithACH(ByVal ACHTNO As String, ByVal ACHTDESC As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
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
                 " And a.seqno ::text = b.memberid " &
                  "  AND b.HomeId = c.HomeId) " &
                    " AND A.CITEMCODE=E.CODENO " &
                    " AND A.CitemCode IN (Select CitemCode From CD068A  Where CD068A.BillHeadFmt IN (" & BillHeadFmt & ")) " &
                    " AND E.PRODUCTCODE IS NULL ", Sign)


                Return result
            Case Else
                Return MyBase.GetNewCanChooseNonePeriodWithACH(ACHTNO, ACHTDESC)
        End Select

    End Function
    Friend Overrides Function GetNewCanChooseProdutWithACH(ByVal ACHTNO As String, ByVal ACHTDESC As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                 " where a.seqno = {0}0 And a.seqno ::text = b.memberid  " &
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
            Case Else
                Return MyBase.GetNewCanChooseProdutWithACH(ACHTNO, ACHTDESC)
        End Select

    End Function
    Friend Overrides Function GetNewCanChooseProduct() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                  " where a.seqno = {0}0 And a.seqno::text = b.memberid  " &
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
            Case Else
                Return MyBase.GetNewCanChooseProduct
        End Select


    End Function
    Friend Overrides Function GetNewCanChooseProduct(ByVal SeqNo As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                  " where a.seqno = " & SeqNo & " And a.seqno::text = b.memberid  " &
                                                  " and b.HomeId = c.HomeId "

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
            Case Else
                Return MyBase.GetNewCanChooseProduct(SeqNo)
        End Select


    End Function
    Friend Overrides Function GetCanChooseNonePeriod(ByVal SeqNo As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
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
                  " And a.seqno::text = b.memberid " &
                   "  AND b.HomeId = c.HomeId) " &
                     " AND A.CITEMCODE=E.CODENO " &
                     " AND E.PRODUCTCODE IS NULL "


                Return result
            Case Else
                Return MyBase.GetCanChooseNonePeriod(SeqNo)
        End Select



    End Function
    Friend Overrides Function GetCanChooseBillNo(ByVal SeqNo As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
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
            " AND a.seqno::text = b.memberid " &
            " AND b.HomeId = c.HomeId) " &
           " AND A.UCCode NOT IN (SELECT CodeNo " &
           " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " &
            " AND A.UCCode IS NOT NULL "
                Return result
            Case Else
                Return MyBase.GetCanChooseBillNo(SeqNo)
        End Select


    End Function
    Friend Overrides Function GetNewCanChooseCharge(ByVal SEQNO As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
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
            " AND A.seqno::text = b.memberid " &
            " AND b.HomeId = c.HomeId) " &
           " AND A.UCCode NOT IN (SELECT CodeNo " &
           " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3, 7, 8)) " &
            " AND A.UCCode IS NOT NULL "

                Return result
            Case Else
                Return MyBase.GetNewCanChooseCharge(SEQNO)
        End Select



    End Function
    Friend Overrides Function GetNewCanChooseCharge() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
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
           " AND A.seqno::text = b.memberid " &
           " AND b.HomeId = c.HomeId) " &
          " AND A.UCCode NOT IN (SELECT CodeNo " &
          " FROM CD013  WHERE PayOk = 1 OR RefNo IN (3,4,7, 8)) " &
           " AND A.UCCode IS NOT NULL ", Sign)

                Return result

            Case Else
                Return MyBase.GetNewCanChooseCharge
        End Select

    End Function
    Friend Overrides Function GetNewSO003C() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT *  FROM SO003C WHERE ServiceId IN ({0}) AND CUSTID In ( " &
                                                    "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                   " where a.seqno = {1} And a.seqno::text = b.memberid  " &
                                                   " and b.HomeId = c.HomeId )"
            Case Else
                Return MyBase.GetNewSO003C
        End Select

    End Function
    Friend Overrides Function GetNewSO003C(ByVal strServiceId As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("SELECT *  FROM SO003C WHERE ServiceId IN (" & strServiceId & ") AND CUSTID In ( " &
                                                    "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                   " where a.seqno = {0}0 And a.seqno ::text = b.memberid  " &
                                                   " and b.HomeId = c.HomeId )", Sign)

            Case Else
                Return MyBase.GetNewSO003C(strServiceId)
        End Select

    End Function
    Friend Overrides Function StopSO002A(ByVal filterCustId As Boolean, ByVal SEQNO As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aSQL As String = String.Empty
                'Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " & _
                '                                            " where a.seqno = " & SEQNO & " And a.memberid = b.memberid  " & _
                '                                            " and b.HomeId = c.HomeId "
                Dim aCustId As String = "Select distinct c.custid from so137 a ,so002c b,so001 c " &
                                                   " where a.seqno = " & SEQNO & " And a.seqno :: text = b.memberid  " &
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
            Case Else
                Return MyBase.StopSO002A(filterCustId, SEQNO)
        End Select

    End Function
    Friend Overrides Function GetSO106RowId() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("select ctid::text from so106 where masterid={0}0", Sign)
            Case Else
                Return MyBase.GetSO106RowId
        End Select

    End Function
    Friend Overrides Function QuerySO033() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT BILLNO,ServiceType FROM  SO033 " & _
                          " WHERE CUSTID= {0}0 " &
                          " AND ACCOUNTNO= {0}1 " & _
                          " AND COMPCODE= {0}2" & _
                           " AND BANKCODE= {0}3 " & _
                          " AND UCCODE > 0 AND CANCELFLAG=0" & _
                          " AND CitemCode= {0}4 " & _
                          " Limit 1"
            Case Else
                Return MyBase.QuerySO033
        End Select

    End Function

    Friend Overrides Function GetSysDate() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "select now()"
            Case Else
                Return MyBase.GetSysDate
        End Select

    End Function
    Friend Overrides Function UpdAuthorize() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Update SO106A Set " &
                               " CitemCodeStr={0}0, " &
                               " CitemNameStr = {0}1," &
                               " UpdEn = {0}2, " &
                               " UpdTime = {0}3 " &
                                " Where CTID::text = {0}4 ", Sign)
                'Return String.Format("Update SO106A Set " &
                '            " UpdEn = {0}0, " &
                '            " UpdTime = {0}1 " &
                '             " Where CTID::text = {0}2 ", Sign)
            Case Else
                Return MyBase.UpdAuthorize
        End Select


    End Function
End Class
