Imports CableSoft.BLL.Utility
Public Class InvoiceDAL
    Inherits DALBasic
    Implements IDisposable
    Private Const IdentifyId As String = " And IdentifyId1 = '1' And IdentifyId2 = 0 "
    Private Language As New CableSoft.BLL.Language.SO61.Invoice
    Protected Friend misOwner As String = Nothing


    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryINV007() As String
        Return String.Format("Select * From INV007 Where Invid={0}0 " & IdentifyId, Sign)
    End Function
    Friend Function QueryINV008() As String
        Return String.Format("Select A.*,0 NEWADD, " &
                " Decode(Nvl((Select Count(*) From INV008A Where A.INVID = INV008A.INVID AND A.SEQ = INV008A.SEQ),'0'),'0',null,'1') MergeFlag " &
                " From INV008 A Where Invid = {0}0  " & IdentifyId & " Order by Seq ", Sign)

    End Function

    Friend Function QueryInv008Single() As String
        Return String.Format("Select A.*,0 NEWADD From INV008 A Where Invid = {0}0 And Seq = {0}1  " & IdentifyId & " Order by Seq ", Sign)
    End Function
    Friend Function QueryINV008A() As String
        Return String.Format("Select * From INV008A Where InvId = {0}0 Order by Seq", Sign)
    End Function
    Friend Function QueryINV008ASingle() As String
        Return String.Format("Select * From INV008A Where InvId = {0}0 and Seq={0}1 Order by Seq", Sign)
    End Function

    Friend Function QueryINV028() As String
        Dim result As String = Nothing
        result = String.Format("Select ItemId CodeNo,Description,RefNo  From INV028 Where CompId = {0}0 Order By ItemId", Sign)
        Return result
    End Function
    Friend Function DeleteInvDataByMainId() As String
        Return String.Format("Delete From INV007 Where MainInvId = {0}0 And CompId = {0}1", Sign)
    End Function
    Friend Function UpdateSOBill(ByVal tbName As String) As String
        Return String.Format("UPDATE " & Me.misOwner & tbName & " SET GUINO =  {0}0, " &
                  "  PREINVOICE =  {0}1 , INVDATE = {0}2, " &
                  " INVOICETIME = {0}3,  INVPURPOSECODE = {0}4, " &
                 "  INVPURPOSENAME = {0}5         " &
                "   WHERE BILLNO = {0}6  AND ITEM = {0}7 " &
                " And GUINO IS NULL", Sign)
    End Function
    Friend Function ReStoreSOBillByINV(ByVal tbSOName As String, ByVal tbInvName As String) As String
        Return String.Format("UPDATE " & Me.misOwner & tbSOName & " SET GUINO =  null, " &
                  "  PREINVOICE =  null , INVDATE = null, " &
                  " INVOICETIME = null,  INVPURPOSECODE = null, " &
                 "  INVPURPOSENAME = null         " &
                "   WHERE BILLNO = {0}0 " &
                "  And ITEM = {0}1 " &
                "  And GUINO Is Not NULL", Sign)
    End Function

    Friend Function DeleteInv007DataByInvId() As String
        Return String.Format("Delete From INV007 Where INVID = {0}0 And CompId = {0}1 " & IdentifyId, Sign)
    End Function
    Friend Function DeleteINV008AByInvId() As String
        Return String.Format("Delete From INV008A Where INVID={0}0 ", Sign)
    End Function
    Friend Function DeleteINV008ByInvId() As String
        Return String.Format("Delete From INV008 Where INVID={0}0 ", Sign)
    End Function
    Friend Function DeleteINV008ABySeq() As String
        Return String.Format("Delete From INV008A Where INVID={0}0 And SEQ = {0}1 ", Sign)
    End Function
    Friend Function QueryINV041() As String
        Return String.Format("Select  ITEMID CODENO, DESCRIPTION  From INV041 Where IDENTIFYID1 = '1'  " &
                   "And IDENTIFYID2 = 0  And COMPID = {0}0 And (REFNO <> 999  Or  REFNO Is NULL ) " &
                    " Order BY ITEMID", Sign)
    End Function

    Friend Function QueryCurrentInv099() As String
        Return String.Format("select * from inv099 a   " &
                          "  where 1=1 " & IdentifyId & " AND  a.compid = {0}0  and {0}1 between startnum and endnum " &
                          "  And yearmonth = {0}2 And PREFIX ={0}3", Sign)
    End Function
    Friend Function QueryIsLock() As String

        Return String.Format("Select count(1) from inv099 a, inv018 b   " &
                          "  where a.compid = b.compid  And a.yearmonth = b.yearmonth " &
                        "    And b.islocked = 'Y'  and a.compid = {0}0   " &
                        "    and {0}1 between startnum and endnum ", Sign)

    End Function
    Friend Function QueryINV099() As String
        Dim result As String = Nothing
        result = String.Format("Select SUBSTR( A.YEARMONTH, 1, 4 ) ||  " &
         "        SUBSTR( A.YEARMONTH, 5, 2 ) YEARMONTH, A.INVFORMAT, " &
         "DECODE(A.INVFORMAT, '1', '" & Language.INVFORMAT1 & "','2', '" & Language.INVFORMAT2 &
         "','3','" & Language.INVFORMAT3 & "','" & Language.INVFORMAT2 & "' ) INVFORMATDESC," &
         "A.PREFIX, A.ENDNUM,A.CURNUM,TO_CHAR( A.LASTINVDATE,'YYYY/MM/DD') LASTINVDATE," &
         "Nvl(UPLOADFLAG,0) UPLOADFLAG2 ," &
         "A.MEMO,Decode(UPLOADFLAG,1,'" & Language.UPLOADFLAGYes & "','" & Language.UPLOADFLAGNo & "')  UPLOADFLAG " &
         ",A.STARTNUM From INV099 A " &
         " Where 1= 1 And A.COMPID = {0}0 And A.YEARMONTH = {0}1 " &
        "  And A.USEFUL = 'Y'   AND A.LASTINVDATE <= TO_DATE( {0}2, 'YYYYMMDD' )  " &
       "    AND A.CURNUM <= A.ENDNUM  " & IdentifyId, Sign)
        Return result





    End Function
    Function GetCurrentAvailableInvCount() As String
        Return String.Format(" SELECT ( A.ENDNUM - A.CURNUM + 1 ) AS COUNTS  " &
                        "   FROM INV099 A    " &
                    "  WHERE  A.COMPID = {0}0   AND A.YEARMONTH = {0}1 " &
                    "    AND A.PREFIX = {0}2   AND A.STARTNUM = {0}3   AND A.USEFUL = 'Y' ", Sign)
    End Function
    Friend Function QueryINVCustInfo(ByVal QueryWhere As Integer) As String
        Dim aWhere As String = Nothing
        Dim result As String = Nothing
        result = "SELECT A.CUSTID, A.CUSTSNAME,   A.CUSTNAME,  " &
                    " B.MAILADDR,  A.TEL1,  A.TEL2,   A.TEL3,   A.APPCONTACTEE1,  " &
                   "  A.APPCONTACTEE2,   A.FINACONTACTEE1,  A.FINACONTACTEE2,  " &
                  "  B.TITLESNAME, B.TITLENAME,  B.BUSINESSID,  B.MAILADDR AS MAILADDR2, " &
                  " B.INVADDR,  B.MEMO,  B.MZIPCODE " &
                  " FROM INV002 A, INV019 B   WHERE A.IDENTIFYID1 = B.IDENTIFYID1  AND A.IDENTIFYID2 = B.IDENTIFYID2 " &
                 "  AND A.COMPID = B.COMPID   AND A.CUSTID = B.CUSTID   AND A.IDENTIFYID1 = '1'   AND A.IDENTIFYID2 = 0 " &
                 "  AND A.COMPID = {0}0 "
        Select Case QueryWhere
            Case 1
                aWhere = " AND instr(A.CUSTID,{0}1) > 0  ORDER BY A.CUSTID "
            Case 2
                aWhere = " AND  (instr( A.CUSTSNAME ,{0}1) > 0 OR " &
                                            " Instr( A.CUSTNAME ,{0}1) > 0  OR  " &
                                            " Instr( B.TITLESNAME, {0}1)> 0 OR  " &
                                           " Instr( B.TITLENAME ,{0}1) > 0)  ORDER BY A.APPCONTACTEE1 "
            Case 3
                aWhere = " AND ( instr( A.APPCONTACTEE1,{0}1)>0  OR    " &
                                 " instr( A.APPCONTACTEE2 ,{0}1) > 0  OR  " &
                                " instr( A.FINACONTACTEE1 ,{0}1) > 0  OR  " &
                                " instr( A.FINACONTACTEE2,{0}1)> 0   )  ORDER BY A.APPCONTACTEE1 "
            Case 4
                aWhere = "AND instr(B.BUSINESSID,{0}1) > 0 ORDER BY B.BUSINESSID "

            Case 5
                aWhere = "' AND ( instr(A.TEL1 ,{0}1) >0 OR  " &
                                " instr( A.TEL2,{0}1)>0  OR  instr( A.TEL3 ,{0}1) > 0 ) ORDER BY A.TEL1 "

        End Select
        result = String.Format(result & aWhere, Sign)
        Return result
    End Function

    Friend Function QueryCustWhere() As String
        Return "Select 1 CodeNo ,'" & Language.QueryWhere1 & "' Description From Dual " &
                    " Union All " &
                    " Select 2 CodeNo ,'" & Language.QueryWhere2 & "'  Description From Dual " &
                    " Union All " &
                    " Select 3 CodeNo ,'" & Language.QueryWhere3 & "'  Description From Dual " &
                    " Union All " &
                    " Select 4 CodeNo ,'" & Language.QueryWhere4 & "'  Description From Dual " &
                    " Union All " &
                    " Select 5 CodeNo ,'" & Language.QueryWhere5 & "'  Description From Dual "



        '" Union All " &
        '" Select 6 CodeNo ,'發票地址 / 郵寄地址' Description From Dual "
    End Function
    Friend Function QuerySOCustByQuery(ByVal whereCode As Integer) As String
        Dim result As String = Nothing
        Select Case whereCode
            Case 1

                result = String.Format(" Select SO002.CustId From " & Me.misOwner & "SO002 Where CustId = {0}0 And CompCode = {0}1  ", Sign)
            Case 2
                Return String.Format(" Select SO002.CustId From " & Me.misOwner & "SO002  Where  " &
                                     " CustID in (Select CustId From " & Me.misOwner & "SO001 Where instr(CustName,{0}0) > 0 And CompCode = {0}1) " &
                                     " And CompCode = {0}1", Sign)
            Case 3
                result = String.Format("Select SO002.CustId From " & Me.misOwner & "SO002 Where CustId IN ( " &
                    "Select CustId From  " & Me.misOwner & "SO001 Where instr(ContName,{0}0)> 0 And CompCode = {0}1) " &
                      " CompCode = {0}1  ", Sign)

            Case 4
                result = String.Format("Select CustId From " & Me.misOwner & "SO002 Where ID in (Select ID From " & Me.misOwner & "SO138 Where INVNO = {0}0 And CompCode = {0}1 ) " &
                 " And CompCode = {0}1", Sign)
            Case 5
                result = String.Format("Select CustId From " & Me.misOwner & "SO002 Where CustId In( Select CustId From " & Me.misOwner & "SO001 Where instr(Tel1,{0}0)>0 Or instr(Tel2,{0}0)>0 or instr(Tel3,{0}0)> 0 And CompCode = {0}1)", Sign)
            Case 6
                result = String.Format("Select CustId From " & Me.misOwner & "SO002 Where ID IN (Select ID From  " & Me.misOwner & "SO138 Where instr(InvAddress,{0}0) > 0 Or  instr(MailAddress,{0}0)> 0  And CompCode = {0}1 )" &
                            " And CompCode = {0}1", Sign)
            Case Else
                result = String.Format(" Select SO002.CustId From " & Me.misOwner & "SO002 Where CustId = {0}0 And CompCode = {0}1  ", Sign)
        End Select
        Return result
    End Function
    Friend Function QueryOldSOCustInfo(ByVal custid As String, ByVal invServiceTypeStr As String, ByVal A_CarrierType As String) As String
        Dim result As String = Nothing
        result = "SELECT   A.CUSTNAME,  A.CUSTID, A.CUSTNAME , A.MAILADDRESS,  " &
                   " C.SERVICETYPE, C.INVTITLE, C.INVNO BUSINESSID,  " &
                   "  C.INVADDRESS, A.INSTADDRESS,C.ACCOUNTNO , B.ZIPCODE, null CHARGETITLE," &
                  " null INVSEQNO,  C.INVPURPOSECODE,  C.INVPURPOSENAME,INVOICEKIND, " &
                  " Decode(Nvl(C.INVOICEKIND,0),0,'" & Language.InvoiceKind1 & "' ,'" & Language.InvoiceKind2 & "') INVOICEKIND2,  " &
                 "  C.DENRECCODE,C.DENRECNAME,  C.EMAIL,A.TEL3 ,  " &
                 " (SELECT NVL(CARRIERTYPECODE, '" & A_CarrierType & "') From " & Me.misOwner & "SO002 Where SO002.SERVICETYPE = C.SERVICETYPE And SO002.Custid = A.CUSTID ) CARRIERTYPECODE, " &
                 " (SELECT DECODE(CARRIERTYPECODE,NULL,A_CarrierId1,CARRIERID1) FROM " & Me.misOwner & "SO002 Where SO002.SERVICETYPE = C.SERVICETYPE And SO002.Custid = A.CUSTID ) CARRIERID1, " &
                 " (SELECT DECODE(CARRIERTYPECODE,NULL,A_CarrierId2,CARRIERID2) FROM " & Me.misOwner & "SO002 Where SO002.SERVICETYPE = C.SERVICETYPE And SO002.Custid = A.CUSTID ) CARRIERID2, " &
                 " (SELECT A_CarrierId1 FROM " & Me.misOwner & "SO002 Where SO002.SERVICETYPE = C.SERVICETYPE And SO002.Custid = A.CUSTID ) A_CARRIERID1, " &
                 " (SELECT A_CarrierId2 FROM " & Me.misOwner & "SO002 Where SO002.SERVICETYPE = C.SERVICETYPE And SO002.Custid = A.CUSTID ) A_CARRIERID2, " &
                 " (SELECT LOVENUM FROM " & Me.misOwner & "SO002 Where SO002.SERVICETYPE = C.SERVICETYPE And SO002.Custid = A.CUSTID ) LOVENUM " &
                 " FROM " & Me.misOwner & "SO001 A, " & Me.misOwner & "SO014 B," & Me.misOwner & "SO002 C, " & Me.misOwner & "SO014 D " &
                 " WHERE A.MAILADDRNO = B.ADDRNO   And A.INSTADDRNO = D.ADDRNO   " &
                "   And A.CUSTID = C.CUSTID  And A.CUSTID IN (" & custid & " ) " &
               "    And C.SERVICETYPE IN ( " & invServiceTypeStr & ") " &
               " UNION ALL " &
              "   Select  UNIQUE  A.CUSTNAME,  A.CUSTID, A.CUSTNAME, D.MAILADDRESS, " &
              "   null  ServiceType , D.INVTITLE, D.INVNO BUSINESSID, " &
              "    D.INVADDRESS, A.INSTADDRESS,  C.ACCOUNTNO ,  B.ZIPCODE, " &
             "    D.CHARGETITLE,  D.INVSEQNO,  D.INVPURPOSECODE,  D.INVPURPOSENAME,INVOICEKIND, " &
            "  DECODE(Nvl(D.INVOICEKIND,0),0,'" & Language.InvoiceKind1 & "','" & Language.InvoiceKind2 & "') INVOICEKIND2, D.DENRECCODE,D.DENRECNAME, " &
            "      NULL EMAIL,A.TEL3, " &
            " (SELECT NVL(CARRIERTYPECODE, '" & A_CarrierType & "') From " & Me.misOwner & "SO138 Where SO138.INVSEQNO = D.INVSEQNO  ) CARRIERTYPECODE, " &
            " (SELECT DECODE(CARRIERTYPECODE,NULL,A_CarrierId1,CARRIERID1) From " & Me.misOwner & "SO138 Where SO138.INVSEQNO = D.INVSEQNO  ) CARRIERID1, " &
            " (SELECT DECODE(CARRIERTYPECODE,NULL,A_CarrierId2,CARRIERID2) From " & Me.misOwner & "SO138 Where SO138.INVSEQNO = D.INVSEQNO  ) CARRIERID2, " &
            " (SELECT A_CarrierId1 From " & Me.misOwner & "SO138 Where SO138.INVSEQNO = D.INVSEQNO  ) A_CARRIERID1, " &
            " (SELECT A_CarrierId2 From " & Me.misOwner & "SO138 Where SO138.INVSEQNO = D.INVSEQNO  ) A_CARRIERID2, " &
            " (SELECT LOVENUM From " & Me.misOwner & "SO138 Where SO138.INVSEQNO = D.INVSEQNO  ) LOVENUM " &
           " FROM  " & Me.misOwner & "SO001 A," & Me.misOwner & "SO014 B," & Me.misOwner & "SO002AD C," & Me.misOwner & "SO138 D," & Me.misOwner & "SO014 E  " &
           " WHERE D.MAILADDRNO = B.ADDRNO   AND A.INSTADDRNO = E.ADDRNO  " &
           "  AND A.CUSTID = C.CUSTID   AND C.INVSEQNO = D.INVSEQNO  " &
           "  AND A.CUSTID IN ( " & custid & ") " &
           "  AND D.STOPFLAG = 0 "
        Return result
    End Function
    Friend Function QuerySOCustInfo(ByVal existsBillNo As String, ByVal custid As String) As String
        Dim result As String = Nothing

        result = String.Format("Select A.*,SO001.CustName,SO001.TEL3,SO001.INSTADDRESS From ( " & _
                   " Select  SO138.MailAddress,SO138.InvTitle,SO138.InvNo,SO138.InvAddress,SO014.ZIPCODE,SO138.InvSeqNo, " &
                    " SO138.InvoiceKind,DECODE(Nvl(SO138.INVOICEKIND,0),0,'" & Language.InvoiceKind1 & "','" & Language.InvoiceKind2 & "') INVOICEKIND2," &
                    "SO138.CHARGETITLE,SO138.CarrierTypeCode,SO138.CarrierId1,SO138.CarrierId2,SO138.LoveNum," &
                    "A_CarrierId1,A_CarrierId2," &
                    " (Select max(InvPurposeCode) From " & Me.misOwner & "SO002 Where SO002.ID = SO138.ID and compcode ={0}0) InvPurposeCode ," &
                    " (Select max(InvPurposeName) From " & Me.misOwner & "SO002 Where SO002.ID = SO138.ID and compcode ={0}0) InvPurposeName ," &
                    " (Select max(EMAIL) From " & Me.misOwner & "SO002 Where SO002.ID = SO138.ID and compcode ={0}0) EMAIL, " &
                    " (Select max(CUSTID) From " & Me.misOwner & "SO002 Where SO002.ID = SO138.ID and compcode ={0}0) CUSTID, " &
                     " SO138.DENRECCODE,SO138.DENRECNAME " &
                    " from " & Me.misOwner & "so138 left join " & Me.misOwner & "so014 on MailAddrNo = so014.addrno" &
                    " where SO138.id in (select id from " & Me.misOwner & "so002 where custid IN (" & custid & ") and compcode ={0}0) " &
                    " And (SO138.INVSEQNO =  " & existsBillNo & " or -1 =" & existsBillNo & ")) A,SO001  " &
                    " Where A.CustId = SO001.CustId And SO001.CompCode = {0}0", Sign)
        Return result
    End Function
    Friend Function QuerySO001(ByVal custid As String) As String
        Return String.Format("Select CustId,CustName,TEL3,INSTADDRESS From " & Me.misOwner & "SO001 Where Custid in (" & custid & ")  And CompCode = {0}0", Sign)
    End Function
    Friend Function QueryBillInvseqNo() As String
        Return String.Format("Select invseqno from " & Me.misOwner & "so033 where billno = {0}0 And item = {0}1 union " & " Select invseqno from " & Me.misOwner & "so034 where billno = {0}0 And item = {0}1", Sign)
    End Function
    Friend Function QueryInv001() As String
        Return String.Format("Select * From Inv001 Where compid={0}0 " & IdentifyId, Sign)
    End Function
    Friend Function UpdateInv099() As String


        Return String.Format("UPDATE INV099    Set CURNUM = LPAD( TO_CHAR( TO_NUMBER( CURNUM ) + {0}0 ), 8, '0' ), " &
                 "   LASTINVDATE =  {0}1  " &
                 "   WHERE 1=1  " & IdentifyId &
                 "     AND COMPID =  {0}2   AND YEARMONTH = {0}3   AND PREFIX = {0}4 " &
                "     AND STARTNUM = {0}5 ", Sign)


    End Function
    Friend Function UpdateIn099Useful() As String

        Return String.Format("UPDATE INV099   SET USEFUL = 'N'  WHERE 1=1 " & IdentifyId &
                "  And COMPID = {0}0   AND YEARMONTH = {0}1 " &
                "   AND PREFIX = {0}2   AND STARTNUM = {0}3 " &
                "    AND to_number(CURNUM) > to_number( ENDNUM)  ", Sign)
    End Function

    Friend Function QueryItemId() As String
        Return String.Format("Select  a.*,inv005.description  ItemIdRefDesc  From ( " &
                              " Select  ItemId CodeNo,Description,Sign,TaxCode,TaxName,ItemIdRef " &
                              " From Inv005 Where  CompId = {0}0 " & IdentifyId & "  And ItemId ={0}1) a  left join inv005 on " &
                            " a.itemidref = Inv005.itemid Order by a.CodeNo", Sign)
    End Function

    Friend Function QueryINV005() As String
        'Return String.Format("Select ItemId CodeNo,Description,Sign,TaxCode,TaxName,ItemIdRef," &
        '                     "(Select Description From INV005 Where ItemId =  ItemIdRef) ItemIdRefName " &
        '                     " From Inv005 " &
        '    " Where CompId = {0}0 " & IdentifyId & " Order By ItemId", Sign)

        Return String.Format("Select  a.*,inv005.description  ItemIdRefDesc  From ( " &
                              " Select  ItemId CodeNo,Description,Sign,TaxCode,TaxName,ItemIdRef " &
                              " From Inv005 Where CompId = {0}0 " & IdentifyId & " ) a  left join inv005 on " &
                            " a.itemidref = Inv005.itemid Order by a.CodeNo", Sign)

    End Function
    Friend Function QueryOldSOBill(ByVal invServiceTypeStr As String, ByVal existsBillNo As String) As String
        Dim result As String = Nothing
        result = "SELECT  '" & Language.noBillClose & "' SOURCE,'33' SOURCE2, A.COMPCODE,  A.CUSTID,   A.BILLNO,  " &
            " A.ITEM,  A.CITEMCODE,  A.CITEMNAME,  A.SHOULDDATE,  " &
           "  DECODE(NVL(A.COMBAMOUNT,0),0,A.SHOULDAMT,A.COMBAMOUNT) SHOULDAMT, A.REALDATE, " &
          "   DECODE(NVL(A.COMBAMOUNT,0),0,A.REALAMT,A.COMBAMOUNT) REALAMT, " &
          "   DECODE(A.COMBSTARTDATE,NULL,A.REALSTARTDATE,COMBSTARTDATE) REALSTARTDATE, " &
         "    DECODE(A.COMBSTOPDATE,NULL,A.REALSTOPDATE,A.COMBSTOPDATE) REALSTOPDATE, " &
         "    A.REALPERIOD,  A.CLCTNAME,  A.STNAME, A.ACCOUNTNO, A.FACISNO,  " &
        "    A.SERVICETYPE,  C.RATE1, B.TAXCODE, C.DESCRIPTION TAXNAME, " &
        "    A.INVSEQNO ,  A.UCCode,  " &
        "   A.CARRIERID1,  A.CARRIERTYPECODE,  A.LOVENUM,  A.CARDLASTNO,  " &
        " (Select ItemIdRef From INV005 Where A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = {0}2) COMBCITEMCODE," &
        " (Select Description From  INV005 Where  ItemId = (Select ItemIdRef From INV005 Where  " &
                " A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = {0}2) " & IdentifyId & " ) COMBCITEMNAME," &
        " (Select Nvl(ShowFaci,0) From INV001 Where CompId = {0}2)   ShowFaci ," &
        " (Select SMARTCARDNO From  " & Me.misOwner & "SO004 Where SEQNO=A.FACISEQNO And FACISNO = A.FACISNO And SO004.CUSTID = A.CUSTID) SMARTCARDNO, " &
        " (Select REFNO From " & Me.misOwner & "CD031 Where CD031.CODENO = A.CMCODE ) CMREFNO, " &
        "   (Select SIGN From INV005 Where A.CITEMCODE = ITEMID " & IdentifyId & " And COMPID = {0}2 ) SIGN," &
        " (Select Nvl(LinkToMIS,'N') From INV001 Where CompId = {0}2 " & IdentifyId & ") LinkToMIS" &
        "    FROM " & Me.misOwner & "SO033 A, " & Me.misOwner & "CD019 B, " & Me.misOwner & "CD033 C  " &
        "   WHERE A.CITEMCODE = B.CODENO    AND B.TAXCODE = C.CODENO " &
        "     AND C.TAXFLAG = 1    AND A.CANCELFLAG = 0 " &
        "     AND ( A.SHOULDAMT + A.REALAMT <> 0 )   " &
        "     AND (A.GUINO IS NULL  AND A.INVOICETIME IS NULL  Or A.GUINO = {0}1)   " &
        "     AND A.CUSTID =  {0}0     AND A.SERVICETYPE IN (" & invServiceTypeStr & ")" &
        "    And A.BillNO || A.ITEM not In ( " & existsBillNo & ") " &
        "  UNION " &
       " SELECT  '" & Language.BillClose & "' SOURCE,'34' SOURCE2, A.COMPCODE,  A.CUSTID,   A.BILLNO,  " &
            " A.ITEM,  A.CITEMCODE,  A.CITEMNAME,  A.SHOULDDATE,  " &
           "  DECODE(NVL(A.COMBAMOUNT,0),0,A.SHOULDAMT,A.COMBAMOUNT) SHOULDAMT, A.REALDATE, " &
          "   DECODE(NVL(A.COMBAMOUNT,0),0,A.REALAMT,A.COMBAMOUNT) REALAMT, " &
          "   DECODE(A.COMBSTARTDATE,NULL,A.REALSTARTDATE,COMBSTARTDATE) REALSTARTDATE, " &
         "    DECODE(A.COMBSTOPDATE,NULL,A.REALSTOPDATE,A.COMBSTOPDATE) REALSTOPDATE, " &
         "    A.REALPERIOD,  A.CLCTNAME,  A.STNAME, A.ACCOUNTNO, A.FACISNO,  " &
        "    A.SERVICETYPE,  C.RATE1, B.TAXCODE, C.DESCRIPTION TAXNAME, " &
        "    A.INVSEQNO ,  A.UCCode,   " &
        "   A.CARRIERID1,  A.CARRIERTYPECODE,  A.LOVENUM,  A.CARDLASTNO,  " &
        " (Select ItemIdRef From INV005 Where A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = {0}2) COMBCITEMCODE," &
        " (Select Description From  INV005 Where  ItemId = (Select ItemIdRef From INV005 Where A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = {0}2) " & IdentifyId & " ) COMBCITEMNAME," &
        " (Select Nvl(ShowFaci,0) From INV001 Where CompId = {0}2)   ShowFaci ," &
        " (Select SMARTCARDNO From " & Me.misOwner & "SO004 Where SEQNO=A.FACISEQNO AND FACISNO = A.FACISNO And SO004.CUSTID = A.CUSTID) SMARTCARDNO, " &
        " (Select REFNO From  " & Me.misOwner & "CD031 Where CD031.CODENO = A.CMCODE ) CMREFNO, " &
        "   (Select SIGN From INV005 Where A.CITEMCODE = ITEMID " & IdentifyId & " And COMPID = {0}2 ) SIGN," &
        " (Select Nvl(LinkToMIS,'N') From INV001 Where CompId ={0}2 " & IdentifyId & ") LinkToMIS" &
        "    FROM " & Me.misOwner & "SO034 A," & Me.misOwner & "CD019 B," & Me.misOwner & "CD033 C  " &
        "   WHERE A.CITEMCODE = B.CODENO    AND B.TAXCODE = C.CODENO " &
        "     AND C.TAXFLAG = 1    AND A.CANCELFLAG = 0 " &
        "     AND ( A.SHOULDAMT + A.REALAMT <> 0 )   " &
        "     AND ( A.GUINO IS NULL AND A.INVOICETIME IS NULL  Or A.GUINO = {0}1)    " &
        "     AND A.CUSTID =  {0}0     AND A.SERVICETYPE IN (" & invServiceTypeStr & ")" &
        "    And A.BillNO || A.ITEM not In ( " & existsBillNo & ")   "



        result = String.Format(result, Sign)
        Return result
    End Function
    Friend Function QuerySOBill(ByVal invServiceTypeStr As String, ByVal existsBillNo As String) As String
        Dim result As String = Nothing
        result = "SELECT  '" & Language.noBillClose & "' SOURCE,'33' SOURCE2, A.COMPCODE,  A.CUSTID,   A.BILLNO,  " &
            " A.ITEM,  A.CITEMCODE,  A.CITEMNAME,  A.SHOULDDATE,  " &
           "  DECODE(NVL(A.COMBAMOUNT,0),0,A.SHOULDAMT,A.COMBAMOUNT) SHOULDAMT, A.REALDATE, " &
          "   DECODE(NVL(A.COMBAMOUNT,0),0,A.REALAMT,A.COMBAMOUNT) REALAMT, " &
          "   DECODE(A.COMBSTARTDATE,NULL,A.REALSTARTDATE,COMBSTARTDATE) REALSTARTDATE, " &
         "    DECODE(A.COMBSTOPDATE,NULL,A.REALSTOPDATE,A.COMBSTOPDATE) REALSTOPDATE, " &
         "    A.REALPERIOD,  A.CLCTNAME,  A.STNAME, A.ACCOUNTNO, A.FACISNO,  " &
        "    A.SERVICETYPE,  C.RATE1, B.TAXCODE, C.DESCRIPTION TAXNAME, " &
        "    A.INVSEQNO ,  A.UCCode,  " &
        "   A.CARRIERID1,  A.CARRIERTYPECODE,  A.LOVENUM,  A.CARDLASTNO,  " &
        " (Select ItemIdRef From INV005 Where A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = A.COMPCODE || '') COMBCITEMCODE," &
        " (Select Description From  INV005 Where  ItemId = (Select ItemIdRef From INV005 Where A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = {0}3) " & IdentifyId & " ) COMBCITEMNAME," &
        " (Select Nvl(ShowFaci,0) From INV001 Where CompId={0}3)   ShowFaci ," &
        " (Select SMARTCARDNO From " & Me.misOwner & "SO004 Where  SEQNO=A.FACISEQNO AND FACISNO = A.FACISNO And SO004.CUSTID = A.CUSTID) SMARTCARDNO, " &
        " (Select REFNO From " & Me.misOwner & "CD031 Where CD031.CODENO = A.CMCODE ) CMREFNO, " &
        "   (Select SIGN From INV005 Where A.CITEMCODE = ITEMID " & IdentifyId & " And COMPID = {0}3) SIGN," &
        " (Select Nvl(LinkToMIS,'N') From INV001 Where CompId = A.CompCode || '' " & IdentifyId & ") LinkToMIS" &
        "    FROM " & Me.misOwner & "SO033 A," & Me.misOwner & "CD019 B, " & Me.misOwner & "CD033 C  " &
        "   WHERE A.CITEMCODE = B.CODENO    AND B.TAXCODE = C.CODENO " &
        "     AND C.TAXFLAG = 1    AND A.CANCELFLAG = 0 " &
        "     AND ( A.SHOULDAMT + A.REALAMT <> 0 )   " &
        "     AND (A.GUINO IS NULL  AND A.INVOICETIME IS NULL  Or A.GUINO = {0}2)   " &
        "     AND A.CUSTID =  {0}0     AND A.SERVICETYPE IN (" & invServiceTypeStr & ")" &
        "    And A.BillNO || A.ITEM not In ( " & existsBillNo & ") " &
        "   And A.InvSeqNo = {0}1  UNION " &
       " SELECT  '" & Language.BillClose & "' SOURCE,'34' SOURCE2, A.COMPCODE,  A.CUSTID,   A.BILLNO,  " &
            " A.ITEM,  A.CITEMCODE,  A.CITEMNAME,  A.SHOULDDATE,  " &
           "  DECODE(NVL(A.COMBAMOUNT,0),0,A.SHOULDAMT,A.COMBAMOUNT) SHOULDAMT, A.REALDATE, " &
          "   DECODE(NVL(A.COMBAMOUNT,0),0,A.REALAMT,A.COMBAMOUNT) REALAMT, " &
          "   DECODE(A.COMBSTARTDATE,NULL,A.REALSTARTDATE,COMBSTARTDATE) REALSTARTDATE, " &
         "    DECODE(A.COMBSTOPDATE,NULL,A.REALSTOPDATE,A.COMBSTOPDATE) REALSTOPDATE, " &
         "    A.REALPERIOD,  A.CLCTNAME,  A.STNAME, A.ACCOUNTNO, A.FACISNO,  " &
        "    A.SERVICETYPE,  C.RATE1, B.TAXCODE, C.DESCRIPTION TAXNAME, " &
        "    A.INVSEQNO ,  A.UCCode,   " &
        "   A.CARRIERID1,  A.CARRIERTYPECODE,  A.LOVENUM,  A.CARDLASTNO,  " &
        " (Select ItemIdRef From INV005 Where A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = A.COMPCODE || '') COMBCITEMCODE," &
        " (Select Description From  INV005 Where  ItemId = (Select ItemIdRef From INV005 Where A.CitemCode = INV005.ItemId " & IdentifyId & " And COMPID = {0}3) " & IdentifyId & ") COMBCITEMNAME," &
        " (Select Nvl(ShowFaci,0) From INV001 Where CompId={0}3)   ShowFaci ," &
        " (Select SMARTCARDNO From " & Me.misOwner & "SO004 Where SEQNO=A.FACISEQNO AND FACISNO = A.FACISNO And SO004.CUSTID = A.CUSTID) SMARTCARDNO, " &
        " (Select REFNO From  " & Me.misOwner & "CD031 Where CD031.CODENO = A.CMCODE ) CMREFNO, " &
        "   (Select SIGN From INV005 Where A.CITEMCODE = ITEMID " & IdentifyId & " And COMPID ={0}3 ) SIGN," &
        " (Select Nvl(LinkToMIS,'N') From INV001 Where CompId = A.CompCode || '' " & IdentifyId & ") LinkToMIS" &
        "    FROM " & Me.misOwner & "SO034 A," & Me.misOwner & "CD019 B, " & Me.misOwner & "CD033 C  " &
        "   WHERE A.CITEMCODE = B.CODENO    AND B.TAXCODE = C.CODENO " &
        "     AND C.TAXFLAG = 1    AND A.CANCELFLAG = 0 " &
        "     AND ( A.SHOULDAMT + A.REALAMT <> 0 )   " &
        "     AND ( A.GUINO IS NULL AND A.INVOICETIME IS NULL  Or A.GUINO = {0}2)    " &
        "     AND A.CUSTID =  {0}0     AND A.SERVICETYPE IN (" & invServiceTypeStr & ")" &
        "    And A.BillNO || A.ITEM not In ( " & existsBillNo & ")  And A.InvSeqNo = {0}1 "



        result = String.Format(result, Sign)
        Return result
    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下方的 Finalize()。
            ' TODO: 將大型欄位設為 null。
        End If
        disposedValue = True
    End Sub

    ' TODO: 只有當上方的 Dispose(disposing As Boolean) 具有要釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Visual Basic 加入這個程式碼的目的，在於能正確地實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
        Dispose(True)
        ' TODO: 覆寫上列 Finalize() 時，取消下行的註解狀態。
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
