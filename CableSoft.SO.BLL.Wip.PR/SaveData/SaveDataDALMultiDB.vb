Public Class SaveDataDALMultiDB
    Inherits SaveDataDAL
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function QuerySO015(ByVal strWhere As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select  CTID:: text,SO015.* From SO015 Where {0} Order by  InDate Desc", strWhere)
            Case Else
                Return MyBase.QuerySO015(strWhere)
        End Select



    End Function
    
    Friend Overrides Function GetRtnWip(ByVal SEQNO As String, ByVal RefNo As Integer) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            'so009.sno=so004d.sno and so004d.seqno=so004.seqno and so004.facicode=cd022.codeno 
            '#8790 第一點 DVR取回單不應該連動退單
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aSql As String
                Dim ref10Where As String = String.Empty
                If RefNo = 10 Then
                    ref10Where = " And SO004D.SeqNo in (Select Seqno From SO004 where SO004D.SeqNo = SO004.SeqNo " & _
                                            " And FaciCode in (Select CodeNo From CD022 Where SO004.FaciCode = CD022.CodeNo And CD022.RefNo =  9))"

                    aSql = String.Format("select CTID::text,A.* from so009 A where A.sno in ( " & _
                               " select SO004d.sno from so004d join so004 on SO004D.SEQNO = SO004.SEQNO  join cd022 " & _
                               "  on so004.facicode = CD022.CODENO and cd022.refno = 9 " & _
                               "  where so004d.seqno in (" & SEQNO & ") " & _
                               " and so004d.sno in (select sno from so009 where mainsno =( select  distinct mainsno from so009 where custid = {0}0 and sno = {0}1) " & _
                               " AND PRCode In (Select CodeNo  FROM CD007 Where RefNo = 9) ) " & _
                               " and (kind ={0}2 or kind = {0}3)) ", Sign)

                Else
                    aSql = String.Format("Select CTID::text,A.* From SO009 A Where A.CustId = {0}0 " & _
                                    " And A.MainSNo = {0}1 " & _
                                    " And A.SNO IN (Select SNO From SO004D Where SO004D.SNO = A.SNO  AND SEQNO IN (" & SEQNO & ") " & ref10Where & _
                                    " AND ( KIND = {0}2 or Kind = {0}3)" & _
                                    " And A.SignDate is null And A.PRCode In (Select CodeNo From CD007 Where RefNo = 9) Order By SNo", Sign)
                End If
               
                Return aSql
                'Return String.Format("Select  CTID::text,A.* From SO009 A Where A.CustId = {0}0 And A.MainSNo = {0}1 " & _
                '                     " AND A.SNO IN (SELECT SNO FROM SO004D WHERE A.SNO = SO004D.SNO " & _
                '                     " AND SO004D.SEQNO IN (SELECT SEQNO FROM SO004 WHERE SO004D.SEQNO = SO004.SEQNO " & _
                '                     " AND SO004.FACICODE IN (SELECT CODENO FROM CD022 WHERE CD022.CODENO = SO004.FACICODE AND CD022.REFNO <> 4)))" &
                '                     " And A.SignDate is null And A.PRCode In (Select CodeNo From CD007 Where RefNo = 9) Order By SNo", Sign)
            Case Else
                Return MyBase.GetRtnWip(SEQNO, RefNo)
        End Select

    End Function
    Friend Overrides Function GetRtnWip(ByVal refno As Integer) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            'so009.sno=so004d.sno and so004d.seqno=so004.seqno and so004.facicode=cd022.codeno 
            '#8790 第一點 DVR取回單不應該連動退單
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select  CTID::text,A.* From SO009 A Where A.CustId = {0}0 And A.MainSNo = {0}1 " & _
                                     " AND A.SNO IN (SELECT SNO FROM SO004D WHERE A.SNO = SO004D.SNO " & _
                                     " AND SO004D.SEQNO IN (SELECT SEQNO FROM SO004 WHERE SO004D.SEQNO = SO004.SEQNO " & _
                                     " AND SO004.FACICODE IN (SELECT CODENO FROM CD022 WHERE CD022.CODENO = SO004.FACICODE AND CD022.REFNO = 9)))" &
                                     " And A.SignDate is null And A.PRCode In (Select CodeNo From CD007 Where RefNo = 9) Order By SNo", Sign)
            Case Else
                Return MyBase.GetRtnWip(refno)
        End Select

    End Function
    Friend Overrides Function GetOtherFailityUtil() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return " Select  A.ctid::text,A.*,B.Description InitPlaceName,C.Description PgName,D.DVRSize " &
                      "From SO004 A left Join CD056 B on A.FaciCode = B.CodeNo   " &
                      " Left Join CD029 C On A.PgNo=C.CodeNo, CD102 D Where " &
                      " 1=1 And A.DVRAuthSizeCode=D.CODENO "
            Case Else
                Return MyBase.GetOtherFailityUtil
        End Select


    End Function
    Friend Overrides Function GetOtherFacility() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql

                Dim strSQL As String = " Select  A.ctid::text,A.*,B.Description InitPlaceName,C.Description PgName,D.DVRSize " &
                      "From SO004 A left Join CD056 B on A.FaciCode = B.CodeNo   " &
                      " Left Join CD029 C On A.PgNo=C.CodeNo, CD102 D Where " &
                      " 1=1 And A.DVRAuthSizeCode=D.CODENO "
                strSQL = String.Format("{1} And SNo = {0}0 ", Sign, strSQL)
                Return strSQL
            Case Else
                Return MyBase.GetOtherFacility
        End Select


    End Function
    Friend Overrides Function GetMoveFaciData(InterDependRefNo As String, strCalcFaciRefNo As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql

                Dim strSQL As String

                strSQL = String.Format("Select A.CTID::text,A.*,B.RefNo FaciRefNo From SO004 A,CD022 B Where A.FaciCode = B.CodeNo  And A.CustId = {0}0 And A.ServiceType = {0}1" &
                               " And A.PRDate is null And A.GetDate is null And (A.PRSNo is null Or Exists (Select 1 From SO004 X Where A.SeqNo = X.ReSeqNo And A.CustId = X.CustId)) And A.InstDate is not null" &
                               " And B.RefNo in (" & strCalcFaciRefNo & ") And B.RefNo in (" & InterDependRefNo & ") Order By (Case B.RefNo when 9 then 0 when 7 then 2 when 8 then 3 else 4 end ),A.SeqNo", Sign)
                Return strSQL
            Case Else
                Return MyBase.GetMoveFaciData(InterDependRefNo, strCalcFaciRefNo)
        End Select
    End Function
End Class
