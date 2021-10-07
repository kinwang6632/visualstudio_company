Imports CableSoft.BLL.Utility

Public Class ChangeFaciDAL
    Inherits DALBasic
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function GetDelete003Citem(ByVal ServiceIds As String) As String
        Return String.Format("SELECT CitemCode FROM SO003C WHERE  ServiceId IN ({0})",
                                                                ServiceIds)
    End Function
    Friend Function QryCanChooseFaciRefNo(ByVal WipType As Integer, ByVal WipData As DataSet) As String
        Dim result As String = Nothing
        Dim CodeNo As String = Nothing

        Select Case WipType
            Case 1
                If WipData.Tables("WIP").Rows.Count = 0 Then
                    CodeNo = -1
                Else
                    If DBNull.Value.Equals(WipData.Tables("WIP").Rows(0).Item("INSTCODE")) Then
                        CodeNo = -1
                    Else
                        CodeNo = WipData.Tables("WIP").Rows(0).Item("INSTCODE")
                    End If
                End If

                
                result = "Select CanChooseFaciRefNo From CD005 Where CodeNo = " & CodeNo
            Case 2
                If WipData.Tables("WIP").Rows.Count = 0 Then
                    CodeNo = -1
                Else
                    If DBNull.Value.Equals(WipData.Tables("WIP").Rows(0).Item("ServiceCode")) Then
                        CodeNo = -1
                    Else
                        CodeNo = WipData.Tables("WIP").Rows(0).Item("ServiceCode")
                    End If
                End If
                result = "Select CanChooseFaciRefNo From CD006 Where CodeNo = " & CodeNo
            Case 3
                If WipData.Tables("WIP").Rows.Count = 0 Then
                    CodeNo = -1
                Else
                    If DBNull.Value.Equals(WipData.Tables("WIP").Rows(0).Item("PRCode")) Then
                        CodeNo = -1
                    Else
                        CodeNo = WipData.Tables("WIP").Rows(0).Item("PRCode")
                    End If
                End If
                result = "Select CanChooseFaciRefNo From CD007 Where CodeNo = " & CodeNo
        End Select
        Return result
    End Function
    Friend Function GetCanChangeFaci(ByVal IncludePR As Boolean, ByVal IncludeDVR As Boolean) As String
        Dim aRefNos As String = "2,3,5,6,7,8,10"
        Dim aPRSQL As String = " 1=1 "
        Dim aRet As String = String.Empty
        If IncludeDVR Then
            aRefNos = aRefNos & ",9"
        End If
        If Not IncludePR Then
            aPRSQL = "PRDATE IS NULL "
        End If
        aRet = String.Format("SELECT A.*,NVL(B.REFNO,0) FACIREFNO " & _
                           " FROM SO004 A,CD022 B " & _
                           " WHERE A.FACICODE = B.CODENO " & _
                           " AND A.CUSTID ={0}0 " & _
                           " AND A.SERVICETYPE = {0}1 AND GETDATE IS NULL " & _
                           " AND B.RefNo In (" & aRefNos & ") " & _
                           " AND " & aPRSQL, Sign)
        Return aRet
    End Function
    Friend Function ChkDataOK() As String
        Dim aRet As String = String.Format("Select Count(1) From SO004D " & _
                                          " Where SeqNo = {0}0 And (Kind = '拆除' or Kind = '移拆' or Kind = '更換') " & _
                                          " And SNo <> {0}1" & _
                                          " And FinTime is null and ReturnCode is null", Sign)
        Return aRet
    End Function
    Friend Function GetCMRateCode(ByVal Type As Integer) As String
        Dim aSQL As String = "SELECT CODENO,DESCRIPTION,BAUDRATEORD FROM CD064 " & _
            " WHERE NVL( STOPFLAG,0) =0 "


        If Type = 0 Then
            aSQL = String.Format(aSQL & " AND BAUDRATEORD < ( " & _
                                 " SELECT BAUDRATEORD FROM CD064 WHERE CODENO ={0}0)", Sign)
        Else
            aSQL = String.Format(aSQL & " AND BAUDRATEORD > ( " & _
                                 " SELECT BAUDRATEORD FROM CD064 WHERE CODENO ={0}0)", Sign)
        End If
        aSQL = aSQL & " ORDER BY CODENO "
        Return aSQL
    End Function
    Friend Function GetDVRSizeCode(ByVal Type As Integer) As String
        Dim aSQL As String = "SELECT CODENO,DESCRIPTION,DVRSIZE " & _
            " FROM CD102 WHERE NVL(STOPFLAG,0) = 0 AND TYPE = 1 "
        If Type = 0 Then
            aSQL = String.Format(aSQL & " AND DVRSIZE <(  " & _
                "SELECT DVRSIZE FROM CD102 WHERE CODENO = {0}0)", Sign)
        Else
            aSQL = String.Format(aSQL & " AND DVRSIZE >(  " & _
                "SELECT DVRSIZE FROM CD102 WHERE CODENO = {0}0)", Sign)
        End If
        aSQL = aSQL & " ORDER BY CODENO "
        Return aSQL
    End Function
    Friend Function GetIPCount(ByVal Type As Integer, ByVal ZeroIPCount As Boolean) As String
        Dim aSQL As String = "SELECT CODENO,DESCRIPTION,IPQUANTITY " & _
            " FROM CD074 WHERE NVL(STOPFLAG,0) = 0 AND IPQUANTITY <> {0}0 "
        If Type = 0 Then
            aSQL = String.Format(aSQL & " AND ( IPAPPLY = 1 OR IPAPPLY = 3 ) ", Sign)
        Else
            aSQL = String.Format(aSQL & " AND ( IPAPPLY = 2 OR IPAPPLY = 3 ) ", Sign)
        End If
        If ZeroIPCount Then
            aSQL = aSQL & " UNION  ALL SELECT CODENO,DESCRIPTION,IPQUANTITY " & _
                        " FROM CD074 WHERE  IPQUANTITY =0 AND NVL(STOPFLAG , 0) =0 "
            aSQL = "SELECT DISTINCT * FROM (" & aSQL & ")"
        Else
            aSQL = "SELECT * FROM (" & aSQL & ") WHERE IPQUANTITY <> 0 "
        End If
        aSQL = aSQL & " ORDER BY CODENO "
        Return aSQL
    End Function

    Friend Function GetFaciCode(ByVal SeqNos As String) As String
        Dim aRet As String = String.Format("SELECT SEQNO,FACICODE,FACINAME, " &
                                            " SMARTCARDNO,STBSNO,RESEQNO,PRDATE," &
                                            "(SELECT REFNO FROM CD022 WHERE FACICODE = CD022.CODENO ) REFNO, " &
                                            "(SELECT NVL(FaciRecoupSNO,0) FROM SO042 WHERE COMPCODE = {0}0  " &
                                            "  AND SO004.SERVICETYPE = SO042.SERVICETYPE) FACIRECOUPSNO" &
                                            " FROM SO004" &
                                          " WHERE 1=1 " &
                                          " And CUSTID = {0}1 And SEQNO IN ( " & SeqNos & " )", Sign)
        Return aRet
    End Function
    Friend Function GetPeriodCharge() As String
        'Dim aRet As String = String.Format("SELECT A.* FROM SO003 A " & _
        '                                 " WHERE A.CUSTID = {0}0 And A.FACISEQNO = {0}1 " & _
        '                                 " And A.SERVICETYPE = {0}2", Sign)
        Dim aRet As String = String.Format("Select A.CustId,A.FaciSeqNo,A.OrderNo,A.InstSNo,A.InstDate," &
            " A.ReInstSNo,A.OpenDate,A.StopSNo,A.CloseDate," &
            " A.PRSNo,A.PRDate,A.ServiceId,A.ProductCode,A.ProductName,B.Amount,B.Period,B.StopFlag," &
            " B.CitemCode,B.CitemName,B.StartDate,B.StopDate,B.FaciSNo,B.ContStartDate,B.ContStopDate," &
            " B.ClctDate,B.BPCode,B.BPName From SO003C A Left Join (Select A.CustId,A.FaciSeqNo," &
            " A.CitemCode,A.CitemName,A.StartDate,A.StopDate,A.ClctDate,A.FaciSNO,A.Period,A.Amount," &
            " B.ContNo,B.ContStartDate,B.ContStopDate,B.BPCode,B.BPName,Nvl(A.StopFlag,0) StopFlag " &
            " From SO003 A Left Join SO003A B " &
            " On ( A.CustId = B.CustId And A.STOPFLAG <>1 And A.FaciSeqNo = B.FaciSeqNo And A.CitemCode = B.CitemCode " &
            " And A.ClctDate Between B.DiscountDate1 And B.DiscountDate2)) B " &
            " On (A.CustId =B.CustId And A.FaciSeqNo = B.FaciSeqNo And A.CITEMCODE = B.CITEMCODE " &
            " And Exists (Select 1 From CD019 C " &
            " Where PeriodFlag=1 And Sign='+' And NoReplaceFlag=0 And " &
            " NoShowCitem=0 And A.ProductCode = C.ProductCode And B.CitemCode = C.CodeNo)) " &
            " Where B.StopFlag= 0 AND A.CustId = {0}0 And A.FaciSeqNo = {0}1 And A.ServiceType = {0}2", Sign)

        Return aRet
    End Function
    Friend Function GetChooseServiceID() As String
        Dim aRet As String = String.Format("Select ServiceId From SO003C " & _
                             " Where CustId = {0}0 And FaciSeqNo = {0}1 " & _
                             " And ( PrDate is null Or InstDate > PrDate ) ", Sign)
        Return aRet
    End Function
    Friend Overridable Function GetAllChangeData(ByVal aSeqNos As String) As String
        Return String.Format("SELECT A.RowId AS  CTID,A.*,NVL(B.REFNO,0) FACIREFNO " &
                             " FROM SO004 A,CD022 B " &
                             " WHERE A.FACICODE = B.CODENO AND A.SEQNO IN ( {0}) ", aSeqNos)
    End Function
    Friend Overridable Function GetChangeData() As String
        Return String.Format("SELECT A.RowId AS CTID,A.*,NVL(B.REFNO,0) FACIREFNO " &
                             " FROM SO004 A,CD022 B " &
                             " WHERE A.FACICODE = B.CODENO AND A.PRDATE IS NULL And A.GetDate IS NULL And  A.SEQNO = {0}0", Sign)
    End Function
    Friend Overridable Overloads Function GetChildFaci() As String
        Return String.Format("SELECT A.RowId AS CTID,A.* FROM SO004 A " &
                             " WHERE CUSTID = {0}0 " &
                             " AND PRDATE IS NULL AND GETDATE IS NULL " &
                             " AND (FACISNO = {0}1 " &
                             " AND FaciCode In (Select CodeNo From CD022 Where RefNo = 4) OR STBSNO = {0}2)", Sign)
    End Function
    Friend Overridable Overloads Function GetChildFaci(ByVal FilterDVR As Boolean) As String
        If FilterDVR Then
            Return String.Format("SELECT A.RowId AS CTID,A.* FROM SO004 A " &
                             " WHERE CUSTID = {0}0 " &
                             " AND PRDATE IS NULL AND GETDATE IS NULL " &
                             " AND (FACISNO = {0}1 AND FaciCode In (Select CodeNo From CD022 Where RefNo = 4))", Sign)
        Else
            Return String.Format("SELECT A.RowId AS CTID,A.* FROM SO004 A " &
                             " WHERE CUSTID = {0}0 " &
                             " AND PRDATE IS NULL AND GETDATE IS NULL " &
                             " AND (FACISNO = {0}1  " &
                             " AND FaciCode In (Select CodeNo From CD022 Where RefNo = 4) OR STBSNO = {0}2)", Sign)
        End If
    End Function
End Class
