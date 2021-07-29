Imports CableSoft.BLL.Utility
Public Class PostACHAuthDAL
    Inherits DALBasic
    Implements IDisposable
    Private Const ReturnOK As String = "RETURNOK"
    Private Const ReturnFail As String = "RETURNFAIL"
    Private Const CancelAuth As String = "CANCELAUTH"
    Private Const PostTerminal As String = "POSTTERMINAL"
    Private Const ResumeData As String = "RESUMEDATA"
    Private Const FetchOK As String = "FETCHOK"
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function GetINCD008Where(ByVal BillHeadFmt As String) As String
        Dim aResult As String = Nothing
        aResult = "Exists(Select CitemCode From SO003 B Where " &
                " 1=1 " &
                " And B.CompCode=SO106.CompCode " &
                " And instr(','||SO106.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " &
                " And B.CitemCode In (Select CitemCode From CD068A " &
                " Where Exists (Select * From CD068 Where CD068.BillHeadFmt=CD068A.BillHeadFmt " &
                                    " And CD068.BillHeadFmt In (" & BillHeadFmt & ") And  CD068.ACHType = 2 )))"

        Return aResult


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
    Friend Overridable Function FetchData(ByVal strWhere As String) As String
        Return "SELECT RowID AS CTID,SO106.Custid,SO106.BankCode,AccountID,SO106.AccountNameId,SO106.CitemStr,SO106.ACHTNo,SO106.ACHSN,SO106.ACHCUSTID,SO106.MasterId," &
                        " (Select BankId From CD018 Where CD018.CodeNo = SO106.BankCode And RowNum = 1 ) BankID " &
                        " From SO106 " & strWhere
    End Function
    Friend Overridable Function getApplySO106SQL(ByVal ACHCUSTID As String, ByVal AccountID As String, ByVal AccountNameID As String, ByVal strWhere As String) As String
        Return "Select RowId as CTID,SO106.* From SO106  Where ACHCUSTID='" & ACHCUSTID &
                                 "' And LPAD(AccountID,14,'0')='" & AccountID & "'" &
                                 " And RPAD(nvl(AccountNameID,'0'),10,'0')='" & AccountNameID & "'" & strWhere
    End Function
    Friend Function ChkCancelAuthStopAll() As String
        'Dim result As String = String.Format("Select Count(Decode(AuthorizeStatus,Null,Null)) A," & _
        '                            "Count(Decode(AuthorizeStatus,1,1)) B," & _
        '                            "Count(Decode(AuthorizeStatus,2,2)) C," & _
        '                            "Count(Decode(AuthorizeStatus,3,3)) D" & _
        '                    " From SO106A" & _
        '                    " Where MasterId={0}0", Sign)
        Dim result As String = String.Format("Select Count(CASE AuthorizeStatus WHEN NULL THEN Null ELSE NULL END) A, " &
                                     "Count(CASE AuthorizeStatus WHEN 1 THEN 1 ELSE NULL END) B," &
                                    " Count(CASE AuthorizeStatus WHEN 2 THEN 2 ELSE NULL END) C," &
                                    " Count(CASE AuthorizeStatus WHEN 3 THEN 3 ELSE NULL END) D " &
                             " From SO106A " &
                             "Where MasterId = {0}0", Sign)
        Return result
    End Function
    Friend Function GetApplyTypeWhere(ByVal strINCD008Where As String, ByVal InACHTNo As String) As String
        Return " And  SO106.SnactionDate Is Null And SO106.SendDate Is Null And nvl(SO106.StopFlag,0) = 0 And " & strINCD008Where &
                           " And Exists(Select * From   SO106A " &
                           " Where SO106.MasterId=SO106A.MasterId And SO106A.StopFlag<>1 And SO106A.StopDate is Null And ACHTNO IN(" & InACHTNo & "))"
    End Function
    Friend Function GetCancelWhere(ByVal StopDate1 As Object, ByVal StopDate2 As Object, ByVal INACHTNo As String) As String
        Dim result As String = String.Empty

        If DBNull.Value.Equals(StopDate1) AndAlso DBNull.Value.Equals(StopDate1) Then
            result = IIf(result = String.Empty, "Nvl(StopFlag,0)=0", result & " And Nvl(StopFlag,0)=0")
        Else
            result = IIf(result = String.Empty, "Nvl(StopFlag,0)=1", result & " And Nvl(StopFlag,0)=1")
        End If


        If result <> String.Empty Then result = " And Exists(Select * From SO106A  " &
                                                        " Where SO106.MasterId=SO106A.MasterId And ACHTNO IN(" & INACHTNo & ") And " & result & ")"
        Return result
    End Function

    Friend Function StopNonePeriod(ByVal CitemStrSeqNo As String) As String
        Dim aCustId = "Select distinct custid from so001  " & _
                                                 " where ID = {0}8 "
        Dim aSQL As String = String.Format("UPDATE SO003 SET BankCode= Null, " &
                       "BANKNAME = Null ,ACCOUNTNO = Null ,CMCode = {0}0,CMName = {0}1 " & _
                       " ,PTCode = {0}2,PTName={0}3 " &
                       ",UpdEn = {0}4,UpdTime = {0}5,NewUpdTime = To_Date({0}6,'yyyymmddhh24miss')  " &
                       " WHERE 1=1  And ACCOUNTNO = {0}7 AND CUSTID In (" & aCustId & ") " & _
                       " And SEQNO IN (" & CitemStrSeqNo & ") " & _
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode is Null)", Sign)
        Return aSQL
    End Function
    Friend Function QueryNonePeriodSEQNo(ByVal CitemStr As String) As String
        Dim aCustId = "Select distinct custid from so001  " & _
                                                 " where ID = {0}0 "
        Dim aSQL As String = String.Format("Select * From  SO003 " & _
                       " WHERE 1=1   AND CUSTID In (" & aCustId & ") " & _
                       " And CitemCode IN (" & CitemStr & ") " & _
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode is Null)", Sign)
        Return aSQL
    End Function
    Friend Overridable Function StopSO003() As String
        Dim Result As String = Nothing
        Result = String.Format("Update SO003 Set " &
                 "BANKCODE=NULL" &
                 ",BANKNAME=NULL" &
                 ",ACCOUNTNO=NULL" &
                 ",CMCode = {0}0 " &
                 ",CMName = {0}1 " &
                 ",PTCode = {0}2 " &
                 ",PTName = {0}3 " &
                 ",UpdEn = {0}4,UpdTime = {0}5,NewUpdTime = To_Date({0}6,'yyyymmddhh24miss') " &
                 " Where CompCode = {0}7 " &
                 " And RowId In ( " &
                "Select so003.rowid  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}8 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) " &
                    " And ACCOUNTNO = {0}9", Sign)
        Return Result
    End Function
    Friend Function ResumeACH() As String
        Return String.Format("Update SO106 Set ACHTNo = {0}0,ACHTDESC = {0}1,CitemStr = {0}2 Where MasterId = {0}3", Sign)
    End Function
    Friend Function UpdResumeData() As String
        Dim result As String = String.Format("Update SO106  Set AuthorizeStatus=1" &
                         ",SnactionDate=To_Date({0}0,'YYYYMMDD'),ReAuthorizeStatus={0}1" &
                          ",UpdEn={0}2,NewUpdTime = To_Date({0}3,'yyyymmddhh24miss') ,UpdTime={0}4," &
                         " SendDate=To_Date({0}5,'YYYYMMDD'), " &
                         "StopFlag = 0,StopDate = null " &
                         " Where" &
                         " ACHCustId={0}6" &
                         " And LPAD(AccountID,14,'0')={0}7 " &
                         " And SnactionDate is Null  " &
                         " And StopFlag = 1 And Masterid ={0}8", Sign)

        Return result
    End Function
    Friend Function GetDefCMCode(ByVal LoginInfo As LoginInfo, ByVal ServiceType As String) As String
        Dim aOBJ As New CableSoft.SO.BLL.Utility.Charge(LoginInfo)
        Try
            Dim aCMCode As String = aOBJ.GetDefaultCMCode(ServiceType).ToString
            Return "SELECT " & aCMCode & " CODENO , Description FROM CD031 WHERE CODENO = " & aCMCode
        Finally
            aOBJ.Dispose()
            aOBJ = Nothing
        End Try

    End Function
    Friend Function GetPTCode() As String
        Return "Select CodeNo,Description,RefNo From CD032 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
    End Function
    Friend Function UpdateCancelAuth() As String
        Return String.Format("Update SO106  Set AuthorizeStatus=2, SENDDATE = null," &
                                    " ReAuthorizeStatus={0}0,UpdEn={0}1,UpdTime={0}2, " &
                                   " NewUpdTime = To_Date({0}3,'yyyymmddhh24miss'), " &
                                   " StopFlag = 1,StopDate = To_Date({0}4,'yyyymmdd') " &
                                    " Where ACHCustId={0}5 And LPAD(AccountId,14,'0')= {0}6 ", Sign)
    End Function
    Friend Function UpdNote() As String
        Return String.Format("Update SO106 Set Note = {0}0 Where MasterId = {0}1", Sign)
    End Function

    Friend Overridable Function UpdateSO106AStatus(ByVal ReturnString As String) As String
        Select Case ReturnString
            Case FetchOK
                Return String.Format("Update  SO106A Set AuthorizeStatus = null,UpdTime = {0}0,UpdEn = {0}1 Where RowId = {0}2", Sign)
            Case ReturnOK
                Return String.Format("Update  SO106A Set AuthorizeStatus = 1,UpdTime = {0}0,UpdEn = {0}1 Where RowId = {0}2", Sign)
            Case CancelAuth
                Return String.Format("Update SO106A Set AuthorizeStatus = 2,UpdTime = {0}0,UpdEn = {0}1 Where RowId = {0}2", Sign)
            Case ReturnFail
                Return String.Format("Update SO106A Set AuthorizeStatus = 3, " &
                                                "UpdTime = {0}0," &
                                                "UpdEn = {0}1, " &
                                                "Notes = {0}2 " &
                                  " Where RowId  = {0}3", Sign)
        End Select

    End Function
    Friend Function StopSO106() As String
        Dim Result As String = Nothing

        Result = String.Format("Update SO106  Set SendDate=Null,ACHCustId=ACHCustId,UpdEn={0}0,UpdTime={0}1 ," &
                             " ReAuthorizeStatus={0}2, " &
                             "Note= Note || decode(note,null,'',  chr(13)  || chr(10) ) || {0}3, " &
                             "NewUpdTime = To_Date({0}4,'yyyymmddhh24miss'), " &
                             "StopFlag = 1,StopDate = To_Date({0}5,'yyyymmdd') " &
                             " Where  LPAD(AccountId,14,'0')={0}6  " &
                             "  And ACHCustId={0}7 " &
                             " And SnactionDate is Null And nvl(StopFlag,0) = 0 ", Sign)

        Return Result
    End Function

    Friend Overridable Function QuerySO003() As String
        Dim result As String = Nothing
        result = String.Format("select * from  SO003 " &
            " Where CompCode = {0}0  And Rowid In ( " &
                "Select so003.rowid  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}1 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
        Return result
    End Function
    Friend Function ClearSO106() As String
        Return String.Format("Update SO106 Set ACHTNO = {0}0,ACHTDesc = {0}1,CitemStr = {0}2 Where MasterId = {0}3 ", Sign)
    End Function
    Friend Function UpdateSO106Note() As String
        'Return String.Format("Update SO106 Set Note = Decode(Note,Null,Null,Note || chr(13) || chr(10)) || {0}0  " & _
        '                                     "  Where MasterId = {0}1", Sign)
        Return String.Format("Update SO106 Set Note = (CASE Note WHEN Null THEN Null ELSE Note || chr(13) || chr(10) || {0}0  END) " &
                                            "  Where MasterId = {0}1", Sign)
    End Function
    Friend Function QuerySO106AErrNote() As String
        Return String.Format("Select * From SO106A " & _
                               " Where MasterId={0}0 " & _
                               " And AuthorizeStatus=3", Sign)
    End Function
    Friend Function ChkSO106AAllFail() As String
        Return String.Format("Select Count(*) From SO106A" & _
                            " Where AchtNO<>{0}0 " & _
                            " And ACHDesc<>{0}1 " & _
                            " And MasterId={0}2 " & _
                            " And (AuthorizeStatus=1 or AuthorizeStatus is Null)", Sign)
    End Function

    Friend Function UpdateFetchOK() As String
        Return String.Format("Update SO106  Set " &
                       "SendDate = {0}0," &
                       "NewUpdTime = To_Date({0}1,'yyyymmddhh24miss'),UpdEn={0}2,UpdTime={0}3 Where" &
                       "  MasterId={0}4 ", Sign)
    End Function
    Friend Function UpdateReturnOK(ByVal strINCD008Where As String) As String
        Return String.Format("Update SO106  Set AuthorizeStatus=1" &
               ",SnactionDate=To_Date({0}0,'YYYYMMDD'),ReAuthorizeStatus={0}1" &
                ",UpdEn={0}2, NewUpdTime = To_Date({0}3,'yyyymmddhh24miss'),UpdTime={0}4 " &
           " Where" &
               " ACHCustId={0}5" &
               " And SendDate=To_Date({0}6,'YYYYMMDD')" &
               " And LPAD(AccountID,14,'0')={0}7" &
               " And nvl(StopFlag,0) = 0 And SnactionDate is Null And " & strINCD008Where, Sign)
    End Function
    Friend Overridable Function QueryUniqueSO106() As String
        Return String.Format("Select rowId as ctid,SO106.* From SO106  Where Masterid = {0}0", Sign)
    End Function
    Friend Function UpdateACHSO003C(ByVal prdServiceId As String) As String
        Dim result As String = Nothing
        result = String.Format("Update SO003C Set CMCode={0}0," & _
           " CMName={0}1," & _
           " PTCode={0}2," & _
           " PTName={0}3," & _
           " UpdTime = {0}4, " & _
           " UpdEn = {0}5, " & _
           " NewUpdTime = To_Date({0}6,'yyyymmddhh24miss'), " & _
           " MasterId = {0}7 " & _
           " Where ServiceId In (" & prdServiceId & ")", Sign)
        Return result
    End Function
    Friend Function UpdateSO003C() As String
        Dim result As String = Nothing
        result = String.Format("Update SO003C Set CMCode={0}0," & _
            " CMName={0}1," & _
            " PTCode={0}2," & _
            " PTName={0}3," & _
            " UpdTime = {0}4, " & _
            " UpdEn = {0}5, " & _
            " NewUpdTime = To_Date({0}6,'yyyymmddhh24miss') " & _
            " Where Masterid = {0}7 ", Sign)
        Return result
    End Function
    Friend Function QueryCanEdit() As String
        Return String.Format("Select Nvl(StartPost,0) From SO041 Where SysID= {0}0", Sign)
      
    End Function

    Friend Function chkAuthority(ByVal GroupField As String) As String
        Return String.Format("Select count(*) From SO029 Where Mid = {0}0 And  Group" & GroupField & "= 1", Sign)
    End Function
    Friend Function GetCompCode(ByVal GroupId As String, ByVal strCD039 As String, ByVal strSO026 As String) As String
        If GroupId = "0" AndAlso 1 = 0 Then
            Return "Select A.CodeNo ,A.Description From " & strCD039 & " A Order By CodeNo"
        End If
        Return String.Format("Select distinct A.CodeNo ,A.Description " & _
                             " From " & strCD039 & " A," & strSO026 & " B  " & _
                             " Where Instr(',' ||B.CompStr|| ',' , ',' ||A.CodeNo|| ',') > 0 " & _
                             " And UserId = {0}0 Order By CodeNO", Sign)
    End Function
    Friend Function UpdNonePeriod(ByVal CitemStrSeqNo As String) As String


        Dim aCustId = "Select distinct custid from so001  " & _
                                                   " where ID = {0}10 "
        Dim aSQL As String = String.Format("UPDATE SO003 SET BankCode={0}0, " &
                       "BANKNAME = {0}1,ACCOUNTNO = {0}2,PTCode = {0}3,PTName={0}4, " &
                       "CMCode = {0}5,CMName = {0}6,UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = To_Date({0}9,'yyyymmddhh24miss')  " &
                       " WHERE 1=1 AND CUSTID In (" & aCustId & ") " & _
                       " And SEQNO IN (" & CitemStrSeqNo & ") " & _
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode is Null)", Sign)
        Return aSQL
    End Function
    Friend Overridable Overloads Function UpdateSO003() As String
        Dim result As String = Nothing
        result = String.Format("Update SO003 Set BankCode={0}0," &
            " BANKNAME={0}1," &
            " ACCOUNTNO={0}2," &
            " PTCode={0}3," &
            " PTName={0}4," &
            " CMCode={0}5," &
            " CMName={0}6" &
            ",UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = To_Date({0}9,'yyyymmddhh24miss')" &
            " Where Rowid In ( " &
                "Select so003.rowid  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}10 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
        Return result
    End Function
    Friend Overridable Function QuerySO003SEQNO(ByVal CitemStr As String) As String
        Dim result As String = Nothing
        result = String.Format("select * from  SO003 " &
            " Where CompCode = {0}0  And Rowid In ( " &
                "Select so003.rowid  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}1 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno )  " &
                    " And SO003.CitemCode In (" & CitemStr & ")", Sign)
        Return result
    End Function
    Friend Overridable Function QuerySO106A(ByVal strInACHTNO As String, ByVal strType As String, ByVal rwCondition As DataRow) As String

        Select Case strType
            Case "1"

                Return String.Format("Select RowId AS CTID,SO106A.* From SO106A  " &
                            "Where MasterId={0}0 " &
                            " And ACHTNO In(" & strInACHTNO & ")" &
                            " And RecordType=0" &
                            " And AuthorizeStatus is null" &
                            " And StopFlag<>1", Sign)
            Case "2"
                Dim aWhere As String = " And 1 =1 "
                If Not DBNull.Value.Equals(rwCondition("STOPDATE1")) AndAlso Not String.IsNullOrEmpty(rwCondition("STOPDATE1")) Then
                    aWhere = aWhere & " And StopDate>=To_Date('" & rwCondition("STOPDATE1") & "','YYYYMMDD')"
                End If
                If Not DBNull.Value.Equals(rwCondition("STOPDATE2")) AndAlso Not String.IsNullOrEmpty(rwCondition("STOPDATE2")) Then
                    aWhere = aWhere & " And StopDate<To_Date('" & rwCondition("STOPDATE2") & "','YYYYMMDD')+1"
                End If
                Return String.Format("Select RowId AS CTID,SO106A.* From SO106A  " &
                            "Where MasterId= {0}0  " &
                            " And ACHTNO In(" & strInACHTNO & ")" &
                            " And RecordType=1" &
                            " And AuthorizeStatus is null" &
                            " And StopFlag=1" &
                            aWhere, Sign)
            Case "3"
                Return String.Format("Select RowId AS CTID,SO106A.* From SO106A  " &
                            "Where MasterId= {0}0" &
                            " And ACHTNO In(" & strInACHTNO & ")" &
                            " And AuthorizeStatus = 1 " &
                            " And StopFlag <> 1 ", Sign)

            Case "4"
                Return String.Format("Select RowId AS CTID,SO106A.* From SO106A " &
                            " Where MasterId={0}0 ", Sign)

        End Select

    End Function
    Friend Function QueryCD068Citem(ByVal strBillHeadFmt As String) As String
        Return "Select distinct CitemCode From CD068A " & _
                        " Where Exists (Select * From CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                            " And CD068.ACHType = 2 And BillHeadFmt In (" & strBillHeadFmt & "))"
    End Function
    Friend Overridable Function FetchStopSO106A(ByVal strInACHTNO As String, ByVal strInACHDesc As String) As String
        Return String.Format("Select ROWID AS CTID,SO106A.* From SO106A  " &
                               "Where MasterId= {0}0 " &
                               " And ACHTNO In(" & strInACHTNO & ")" &
                               " And ACHDesc In(" & strInACHDesc & ")" &
                               " And RecordType=1 And AuthorizeStatus=4", Sign)
    End Function
    Friend Overridable Function FetchApplySO106A(ByVal strInACHTNO As String, ByVal strInACHDesc As String) As String
        Return String.Format("Select RowID AS CTID,SO106A.* From SO106A  " &
                               "Where MasterId= {0}0 " &
                               " And ACHTNO In(" & strInACHTNO & ")" &
                               " And ACHDesc In(" & strInACHDesc & ")" &
                               " And RecordType=0 And AuthorizeStatus=4 " &
                               " And StopFlag<>1 And StopDate is Null Order By ACHTNO", Sign)

    End Function
    Friend Function QueryAsignBankID() As String
        Return String.Format("Select BankID From CD018 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QueryBankId() As String
        Return String.Format("Select CodeNo,Description From CD018 " & _
                             " Where CD018.PRGNAME like '%POST%'  AND  COMPCODE ={0}0 " & _
                             " AND STOPFLAG <> 1", Sign)
    End Function
    Friend Function QueryBillHeadFmt() As String
        Return "Select ACHTDESC Description,ACHTNO CodeNo,BillHeadFmt From CD068 " & _
                    " Where ACHTNO Is Not Null And ACHTDesc Is not Null And ACHType=2 "

    End Function
    Friend Function QueryCD068A() As String
        Return "Select * From CD068A Where BillHeadFmt in (" & _
            "Select BillHeadFmt From CD068 " & _
                    " Where ACHTNO Is Not Null And ACHTDesc Is not Null And ACHType=2 )"
    End Function
    Friend Function QueryCitemCode() As String
        Return "Select CodeNo,Description From CD019 Where StopFlag <> 1 order by codeno"
    End Function
    Friend Function QueryCD068() As String
        Return "Select BillHeadFmt || ACHTNO PKName, BillHeadFmt,ACHTNO,ACHTDesc From CD068 " & _
            " Where ACHTNO Is Not Null And ACHTDesc Is not Null And ACHType=2"
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
