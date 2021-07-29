Imports CableSoft.BLL.Utility
Public Class BatchCreateDAL
    Inherits DALBasic
    Implements IDisposable
    Private Language As New CableSoft.BLL.Language.SO61.Invoice
    Private Const IdentifyId As String = " And IdentifyId1 = '1' And IdentifyId2 = 0 "
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryAutoCreateNum() As String
        Return String.Format("Select Nvl(AutoCreateNum,9999) From INV001 " &
                " Where CompId={0}0 " & IdentifyId, Sign)
    End Function
    Friend Function QueryAllCreateData() As String
        Return String.Format("SELECT SUM( A.SALEAMOUNT ) SALEAMOUNT, " &
        "        SUM( A.TAXAMOUNT ) TAXAMOUNT,   " &
        "        SUM( A.INVAMOUNT ) INVAMOUNT,    " &
        "        COUNT(*) TOTAL          " &
        "  FROM INV031 A  WHERE 1=1 AND COMPID = {0}0 " & IdentifyId, Sign)
    End Function
    Friend Function QueryElectricInvData() As String
        Return String.Format("SELECT SUM( A.SALEAMOUNT ) SALEAMOUNT, " &
        "        SUM( A.TAXAMOUNT ) TAXAMOUNT,   " &
        "        SUM( A.INVAMOUNT ) INVAMOUNT,    " &
        "        COUNT(*) TOTAL          " &
        "  FROM INV031 A  WHERE 1=1 AND INVOICEKIND = 1 And COMPID = {0}0 " & IdentifyId, Sign)
    End Function
    Friend Function QueryNoneElectricInvData() As String
        Return String.Format("SELECT SUM( A.SALEAMOUNT ) SALEAMOUNT, " &
        "        SUM( A.TAXAMOUNT ) TAXAMOUNT,   " &
        "        SUM( A.INVAMOUNT ) INVAMOUNT,    " &
        "        COUNT(*) TOTAL          " &
        "  FROM INV031 A  WHERE 1=1 AND Nvl(INVOICEKIND,0) = 0 And COMPID = {0}0 " & IdentifyId, Sign)
    End Function
    Friend Function QueryStarttoEndInv() As String
        Return String.Format(" SELECT MIN( INVID ) MININV,   " &
                "  MAX( INVID ) MAXINV     " &
                "   FROM  INV031   WHERE COMPID = {0}0 ", Sign)
    End Function
    Friend Function exesf() As String
        Return String.Format("Select * From sf_assigninvid({0}0,{0}1,{0}2,{0}3,{0}4,{0}5,{0}6,{0}7,{0}8,{0}9," &
                        "0{10},0{11},0{12},0{13},0{14},0{15},0{16},0{17},0{18},0{19})", Sign)
    End Function
    Friend Overridable Function QueryUnusualInv() As String
        Return String.Format("Select max(LogTime) LogTime From INV033 Where CompId= {0}0 And LOGTIME = {0}1", Sign)

        'Return String.Format("select max(LogTime) LogTime from inv033 where CompId= {0}0  And logtime = to_date('{0}1', 'yyyy/mm/dd hh24:mi:ss') ", Sign)
        'Return "Select max(LogTime) LogTime from inv033 where compid = '3' and rownum<=10"
    End Function
    Friend Overridable Function QueryCanCreateInv(ByVal invoicekind As Integer) As String
        Dim result As String = Nothing
        Dim invoicekindStr As String = " - 1"
        If invoicekind = 2 Then invoicekindStr = "InvoiceKind"
        result = "Select SUM( COUNTS ) Count FROM  " &
                           " ( Select B.SEQ, CEIL( COUNT( B.SEQ ) / {0}0 ) As COUNTS  " &
                           "   FROM INV016 A, INV017 B   WHERE A.SEQ = B.SEQ  " &
                           "   And A.COMPID =  {0}1  And A.CHARGEDATE BETWEEN  " &
                           " {0}2  And {0}3  And A.BEASSIGNEDINVID = 'N'  " &
                           "   AND A.ISVALID = 'Y'   AND A.HOWTOCREATE = {0}4  " &
                           "   AND A.SHOULDBEASSIGNED = 'Y'  AND A.INVAMOUNT > 0  " &
                          "   AND A.TAXTYPE <> '0'  AND A.STOPFLAG = 0 " &
                          "   AND B.SHOULDBEASSIGNED = 'Y' And   (Nvl(InvoiceKind,0)  = {0}5 Or InvoiceKind = " & invoicekindStr & " )" &
                          "  GROUP BY B.SEQ  ) "
        Return String.Format(result, Sign)
    End Function
    Friend Function QueryInvoiceKind() As String
        Return "Select 1 CodeNo,'" & Language.InvoiceKind3 & "' Description From Dual Union All " &
                        "Select 2 CodeNo,'" & Language.InvoiceKind1 & "' Description From Dual  Union All " &
                     "Select 3 CodeNo,'" & Language.InvoiceKind2 & "' Description From Dual "
    End Function
    Friend Function QueryINV063() As String
        Return String.Format("Select * From INV063 Where CODENO = {0}0", Sign)
    End Function
    Friend Function QueryINV001() As String
        Return String.Format("Select * From INV001 Where CompId = {0}0", Sign)
    End Function
    Friend Function GetCompCode(ByVal GroupId As String, ByVal strCD039 As String, ByVal strSO026 As String) As String
        If GroupId = "0" Then
            Return "Select A.CodeNo ,A.Description From " & strCD039 & " A Order By CodeNo"
        End If
        Return String.Format("Select distinct A.CodeNo ,A.Description " &
                             " From " & strCD039 & " A," & strSO026 & " B  " &
                             " Where Instr(',' ||B.CompStr|| ',' , ',' ||A.CodeNo|| ',') > 0 " &
                             " And UserId = {0}0 Order By CodeNO", Sign)
    End Function
    Friend Function QueryExceptInvDetail(ByVal DataType As Integer) As String
        Dim result As String = Nothing
        If DataType = 1 Then
            result = "SELECT   A.SEQ,  A.BILLID, A.BILLIDITEMNO,   " &
                 " DECODE( A.TAXTYPE, '1', '" & Language.TAXTYPE1 & "','2', '" & Language.TAXTYPE2 & "', '3', '" & Language.TAXTYPE3 & "', A.TAXTYPE ) AS TAXDESCRIPTION,  " &
                "  A.CHARGEDATE, A.DESCRIPTION AS ITEMDESCRIPTION,  A.QUANTITY,  A.UNITPRICE, " &
                "  A.TAXRATE, A.TAXAMOUNT, A.TOTALAMOUNT, A.STARTDATE,  A.ENDDATE, " &
               "  A.CHARGEEN,  A.SERVICETYPE   FROM INV017 A   WHERE A.SEQ = {0}0 " &
               "  AND A.SHOULDBEASSIGNED = 'Y'   AND A.TOTALAMOUNT <> 0   AND A.TAXTYPE <>'0' " &
               "  ORDER BY A.BILLIDITEMNO "
        Else
            'result = "SELECT   A.SEQ,  A.BILLID, A.BILLIDITEMNO,   " &
            '    " DECODE( A.TAXTYPE, '1', '" & Language.TAXTYPE1 & "','2', '" & Language.TAXTYPE2 & "', '3', '" & Language.TAXTYPE3 & "', A.TAXTYPE ) AS TAXDESCRIPTION,  " &
            '   "  A.CHARGEDATE, A.DESCRIPTION AS ITEMDESCRIPTION,  A.QUANTITY,  A.UNITPRICE, " &
            '   "  A.TAXRATE, A.TAXAMOUNT, A.TOTALAMOUNT, A.STARTDATE,  A.ENDDATE, " &
            '  "  A.CHARGEEN,  A.SERVICETYPE   FROM INV017 A   WHERE A.SEQ = {0}0 " &
            '  "    AND ( A.SHOULDBEASSIGNED = 'N'  OR   A.TOTALAMOUNT = 0 OR   A.TAXTYPE = '0'  )  " &
            '  "  ORDER BY A.BILLIDITEMNO "
            result = "SELECT   A.SEQ,  A.BILLID, A.BILLIDITEMNO,   " &
               " DECODE( A.TAXTYPE, '1', '" & Language.TAXTYPE1 & "','2', '" & Language.TAXTYPE2 & "', '3', '" & Language.TAXTYPE3 & "', A.TAXTYPE ) AS TAXDESCRIPTION,  " &
              "  A.CHARGEDATE, A.DESCRIPTION AS ITEMDESCRIPTION,  A.QUANTITY,  A.UNITPRICE, " &
              "  A.TAXRATE, A.TAXAMOUNT, A.TOTALAMOUNT, A.STARTDATE,  A.ENDDATE, " &
             "  A.CHARGEEN,  A.SERVICETYPE   FROM INV017 A   WHERE A.SEQ = {0}0 " &
             "  ORDER BY A.BILLIDITEMNO "
        End If
        Return String.Format(result, Sign)
    End Function

    Friend Function QueryExceptInvInfo(ByVal orderByNum As Integer, ByVal invoicekind As Integer, ByVal DataType As Integer) As String
        Dim result As String = Nothing
        Dim orderby As String = Nothing
        Dim invoicekindStr As String = "-1"
        Dim dataTypeWhere As String = " And A.SHOULDBEASSIGNED = 'Y'  And A.INVAMOUNT > 0  And A.TAXTYPE <> '0'   And A.STOPFLAG = 0 "
        If invoicekind = 2 Then invoicekindStr = "InvoiceKind"
        If DataType <> 1 Then
            dataTypeWhere = "  And (A.SHOULDBEASSIGNED = 'N' OR  " &
                   "     A.INVAMOUNT <= 0  OR   A.TAXTYPE = '0'   )  AND A.STOPFLAG = 0 "

        End If
        result = "Select A.SEQ,   A.CUSTID,   A.TITLE,   A.TEL,  A.BUSINESSID,   A.ZIPCODE,  A.INVADDR,  " &
            " A.MAILADDR, A.CHARGEDATE,  " &
            "DECODE( A.TAXTYPE, '1', '" & Language.TAXTYPE1 & "',   '2', '" & Language.TAXTYPE2 & "',   '3', '" & Language.TAXTYPE3 & "',  A.TAXTYPE ) AS DESCRIPTION,   " &
            " A.TAXRATE,  A.SALEAMOUNT,    A.TAXAMOUNT,   A.INVAMOUNT,  " &
            " DECODE( A.HOWTOCREATE, '1', '" & Language.HowtoCreate1 & "', '2', '" & Language.HowtoCreate2 & "', A.HOWTOCREATE ) HOWTOCREATE, " &
            " A.CHARGETITLE,     A.UPTTIME,   A.UPTEN " &
            " FROM INV016 A  WHERE A.COMPID = {0}0   And A.CHARGEDATE BETWEEN  {0}1 " &
            " And  {0}2   And A.BEASSIGNEDINVID = 'N'  And A.ISVALID = 'Y'  And A.HOWTOCREATE = {0}3 " &
            " And (Nvl(InvoiceKind,0)  = {0}4 Or InvoiceKind = " & invoicekindStr & " )" & dataTypeWhere
        '" And A.SHOULDBEASSIGNED = 'Y'  And A.INVAMOUNT > 0  And A.TAXTYPE <> '0'   And A.STOPFLAG = 0   "
        Select Case orderByNum
            Case 1
                orderby = " Order by CHARGEDATE,CustID"
            Case 2
                orderby = " Order by ZIPCODE,CustID"
            Case 3
                orderby = " Order by A.CHARGEDATE, A.ZIPCODE "
            Case 4
                orderby = " Order by A.CHARGEDATE, A.ZIPCODE,A.CUSTID "
        End Select
        result = result & orderby
        Return String.Format(result, Sign)

    End Function
    Friend Function QueryINV099() As String
        Return String.Format(" SELECT A.IDENTIFYID1,A.IDENTIFYID2,A.COMPID,A.STARTNUM,A.PREFIX, " &
                     " to_Char(A.LASTINVDATE,'yyyy/mm/dd') LASTINVDATE ,A.YEARMONTH,A.CURNUM,A.MEMO, A.EndNum," &
                    " DECODE(NVL(A.UPLOADFLAG,0),0,'" & Language.UPLOADFLAGNo & "','" & Language.UPLOADFLAGYes & "') UPLOADFLAG  " &
                    " FROM INV099 A " &
                   "  WHERE A.COMPID ={0}0  AND INVFORMAT=1 AND USEFUL='Y'  ORDER BY A.YEARMONTH DESC ", Sign)
    End Function
    Friend Function QueryGridOrder() As String
        Return "Select 1 CodeNo,'" & Language.GridOrder1 & "' Description From Dual Union All " &
                        "Select 2 CodeNo,'" & Language.GridOrder2 & "' Description From Dual  Union All " &
                     "Select 3 CodeNo,'" & Language.GridOrder3 & "' Description From Dual Union All " &
                     "Select 4 CodeNo,'" & Language.GridOrder4 & "' Description From Dual "
    End Function
    Friend Function QueryNoCreateAmount(ByVal invoicekind As Integer) As String
        Dim result As String = Nothing
        Dim invoicekindStr As String = "-1"
        If invoicekind = 2 Then invoicekindStr = "InvoiceKind"
        result = " SELECT    SUM( B.QUANTITY * B.UNITPRICE ) SALEAMOUNT,  " &
                    "           SUM( B.TAXAMOUNT ) TAXAMOUNT,  " &
                    "        SUM( B.TOTALAMOUNT ) INVAMOUNT,    " &
                    "      Count(distinct A.SEQ) NOCOUNT " &
                    "   FROM INV016 A, INV017 B         " &
                    "  WHERE A.COMPID =  {0}0   AND A.SEQ = B.SEQ   " &
                    "    AND A.CHARGEDATE BETWEEN  {0}1 AND  {0}2   " &
                    "    AND A.BEASSIGNEDINVID = 'N'   AND A.ISVALID = 'Y'  " &
                    "    AND A.HOWTOCREATE = {0}3   " &
                    " And (Nvl(InvoiceKind,0)  = {0}4 Or InvoiceKind = " & invoicekindStr & " )" &
                    " AND ( A.SHOULDBEASSIGNED = 'N' OR  " &
                   "     A.INVAMOUNT <= 0  OR   A.TAXTYPE = '0' OR  B.SHOULDBEASSIGNED = 'N' OR  " &
                   "     B.TOTALAMOUNT = 0  OR  B.TAXTYPE = '0' )  AND A.STOPFLAG = 0  "
        Return String.Format(result, Sign)
    End Function
    Friend Function QueryMustCreateAmount(ByVal invoicekind As Integer) As String
        Dim result As String = Nothing
        Dim invoicekindStr As String = "-1"
        If invoicekind = 2 Then invoicekindStr = "InvoiceKind"
        result = "SELECT SUM( B.QUANTITY  * B.UNITPRICE ) AS SALEAMOUNT,  " &
                    "       SUM( B.TAXAMOUNT ) AS TAXAMOUNT,       " &
                   "       SUM( B.TOTALAMOUNT ) AS INVAMOUNT     " &
                  "   FROM INV016 A, INV017 B         " &
                  " WHERE A.SEQ = B.SEQ     AND A.COMPID =  {0}0    " &
                  "   AND A.CHARGEDATE BETWEEN  {0}1 AND {0}2  " &
                  "   AND A.BEASSIGNEDINVID = 'N'  AND A.ISVALID = 'Y'  AND A.HOWTOCREATE = {0}3 " &
                  "  And (Nvl(InvoiceKind,0)  = {0}4 Or InvoiceKind = " & invoicekindStr & " )" &
                 "   AND A.SHOULDBEASSIGNED = 'Y'   AND A.INVAMOUNT > 0 " &
                 "   AND A.TAXTYPE <> '0'   AND A.STOPFLAG = 0  AND B.SHOULDBEASSIGNED = 'Y' "
        Return String.Format(result, Sign)
    End Function
    Friend Function QueryHowtoCreate() As String
        Return "Select 1 CodeNo,'" & Language.HowtoCreate1 & "' Description From Dual Union All " &
                       "Select 2 CodeNo,'" & Language.HowtoCreate2 & "' Description From Dual "

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
