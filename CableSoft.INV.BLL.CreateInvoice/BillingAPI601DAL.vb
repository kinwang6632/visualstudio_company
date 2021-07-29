Imports CableSoft.BLL.Utility
Public Class BillingAPI601DAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryInv008() As String
        Dim result As String = Nothing
        result = String.Format("SELECT A.CUSTID, B.SEQ, B.BILLID, B.BILLIDITEMNO," & _
                                           " A.INVUSEID, A.INVUSEDESC FROM INV007 A, INV008  B  " & _
                                           "  WHERE A.INVID = B.INVID  AND A.IDENTIFYID1 = {0}0 " & _
                                           "  AND LinkToMIS = 'Y' " & _
                                           "    AND A.IDENTIFYID2 = {0}1  AND A.COMPID = {0}2   AND A.INVID = {0}3", Sign)
        Return result
    End Function
    Function QueryInv008A() As String
        Dim result As String = Nothing
        result = String.Format("SELECT B.SEQ, B.BILLID, B.BILLIDITEMNO " & _
                                          " From INV008A B  " & _
                                           "  WHERE 1=1  AND INVID = {0}0 AND SEQ = {0}1", Sign)
        Return result

    End Function
    Friend Function QueryBillNo(ByVal index As Integer) As String
        Dim aTable As String = "INV008A"
        Dim result As String = Nothing
        If index <> 0 Then aTable = "INV008"

        result = String.Format("SELECT A.CUSTID, B.SEQ, B.BILLID, B.BILLIDITEMNO," & _
                                            " A.INVUSEID, A.INVUSEDESC FROM INV007 A, " & aTable & " B  " & _
                                            "  WHERE A.INVID = B.INVID  AND A.IDENTIFYID1 = {0}0 " & _
                                            "    AND A.IDENTIFYID2 = {0}1  AND A.COMPID = {0}2   AND A.INVID = {0}3", Sign)


        Return result
    End Function
    Friend Function QueryInv007() As String
        Return String.Format("SELECT MAININVID, INVID,INVDATE, ISOBSOLETE,HowToCreate,InvUseId,InvUseDesc " & _
                                            "   FROM INV007  WHERE IDENTIFYID1 = '1'  AND IDENTIFYID2 = 0 " & _
                                            "    AND COMPID = {0}0   AND INVID = {0}1 ORDER BY INVID ", Sign)
    End Function
    Function QueryCountINV014() As String
        Return String.Format("select count(1) from inv014  where identifyid1 = '1' " & _
                                                " and identifyid2 = 0  and compid = {0}0   and IsObsolete = 'N' " & _
                                                "  and invid = {0}1", Sign)
    End Function
    Friend Function QueryInv018() As String
        Return String.Format("SELECT Count(*) From INV018  Where " & _
                             " ISLOCKED = 'Y'  AND IDENTIFYID1 = '1'    AND IDENTIFYID2 = 0 " & _
                             "  AND COMPID = {0}0  AND YEARMONTH = {0}1", Sign)

    End Function
    Friend Overridable Function DropInv007() As String
        Dim result As String = Nothing
        result = String.Format("UPDATE INV007  " &
                                        "  SET ISOBSOLETE = 'Y',OBSOLETEID = {0}0, " &
                                        "  OBSOLETEREASON = (SELECT Description FROM INV006 WHERE IdentifyId1 = '1' AND IdentifyId1 = 0 AND ItemId={0}1), " &
                                       "   CANMODIFY = 'N', UPTTIME = SYSDATE, UPTEN = {0}2   WHERE IDENTIFYID1 = '1' " &
                                       "    AND IDENTIFYID2 = 0    AND COMPID = {0}3   AND INVID = {0}4", Sign)

        Return result

    End Function
    Friend Function InsInv024(ByVal index As Integer) As String
        Dim result As String = Nothing
        If index = 0 Then
            result = String.Format("INSERT INTO INV024 ( " &
               "    COMPID, INVID, BILLID, BILLIDITEMNO, CUSTID )" &
              "   SELECT DISTINCT A.COMPID, A.INVID, " &
              "  DECODE(B.BILLID,NULL,'',B.BILLID) BILLID, " &
              " DECODE(B.BILLIDITEMNO,NULL,0,B.BILLIDITEMNO), " &
             "  A.CUSTID     FROM INV007 A, INV008A  B " &
             "  WHERE A.INVID = B.INVID   AND A.ISOBSOLETE = 'Y'  AND B.BILLID IS NOT NULL " &
            "    AND B.BILLIDITEMNO IS NOT NULL   AND A.IDENTIFYID1 = '1' AND A.IDENTIFYID2 =0 " &
            "  AND A.COMPID = {0}0     AND A.INVID = {0}1 " &
            " AND NVL(B.BILLID,'X') || NVL(BILLIDITEMNO,-1) NOT IN  " &
            "        (SELECT NVL(BILLID,'X') || NVL(BILLIDITEMNO,-1) FROM INV008 WHERE INVID = B.INVID) ", Sign)
        Else
            result = String.Format(" INSERT INTO INV024 (  " &
                     " COMPID, INVID, BILLID, BILLIDITEMNO, CUSTID ) " &
                " SELECT DISTINCT A.COMPID, A.INVID, " &
             " DECODE(B.BILLID,NULL,'',B.BILLID) BILLID, " &
             " DECODE(B.BILLIDITEMNO,NULL,0,B.BILLIDITEMNO), " &
              " A.CUSTID " &
            "    FROM INV007 A, INV008 B left join  INV008A C " &
            " on B.Invid = C.Invid and B.BILLID = C.BILLID And B.BILLIDITEMNO = C.BILLIDITEMNO  AND C.INVID IS NULL" &
            "   WHERE A.INVID = B.INVID  AND A.ISOBSOLETE = 'Y' " &
            " AND A.IDENTIFYID1 = '1'  AND A.IDENTIFYID2 = 0 " &
            "   AND A.COMPID = {0}0  AND A.INVID = {0}1 ", Sign)
        End If
        Return result
    End Function
    Friend Function GetEmpName(ByVal aOwner As String) As String
        Dim strSQL As String
        strSQL = String.Format("Select EmpName From " & aOwner & ".CM003 Where EmpNo = {0}0", Sign)
        Return strSQL
    End Function
    Friend Function updSO033SO034(ByVal aOwner As String, ByVal index As Integer) As String
        Dim result As String = Nothing
        If index = 0 Then
            result = String.Format("UPDATE " & aOwner & ".SO033 " &
                  "      SET PREINVOICE = NULL,   GUINO = NULL, " &
                  "          INVDATE = NULL,  " &
                  "          INVOICETIME = NULL, " &
                  "          INVPURPOSECODE = NULL, " &
                  "          INVPURPOSENAME = NULL , " &
                  "          updtime = {0}0,newupdtime={0}1,upden={0}2 " &
                  "    WHERE GUINO = {0}3 AND COMPCODE = {0}4", Sign)
        Else
            result = String.Format("UPDATE " & aOwner & ".SO034  " &
                        "    SET PREINVOICE = NULL, " &
                        "    GUINO = NULL,  " &
                        "     INVDATE = NULL, " &
                        "     INVOICETIME = NULL, " &
                        "     INVPURPOSECODE = NULL, " &
                        "     INVPURPOSENAME = NULL,  " &
                          "    updtime = {0}0,newupdtime={0}1,upden={0}2 " &
                        "     WHERE GUINO = {0}3 " &
                        "     AND COMPCODE = {0}4 " &
                        "     AND ( PREINVOICE IS NULL OR  " &
                        "    PREINVOICE IN ( 0, 1, 2, 3 ))  ", Sign)
        End If
      
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
