Imports CableSoft.BLL.Utility
Imports System.Web
Imports System.Xml
Imports System.Data.Common
Public Class PostACHAuth
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New PostACHAuthDALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.PostACHAuth
    Private FNowDate As Date = Date.Now
    Private Const tbCompCodeName As String = "CompCode"    
    Private Const tbBankIdName As String = "BankId"
    Private Const tbBillHeadFmtName As String = "BillHeadFmt"
    Private Const tbCitemCodeName As String = "CitemCode"
    Private Const tbCD068Name As String = "CD068"
    Private Const tbCD068AName As String = "CD068A"
    Private Const tbStopAllName As String = "StopAll"
    Private Const tbInputACTHNOName As String = "InputACTHNO"
    Private Const errorFileName As String = "SO3297Err.txt"
    Private Const ApplyFileName As String = "SO3297Apply.txt"    
    Private Const StopFileName As String = "SO3297Stop.txt"
    Private FNow As DateTime = Date.Now
    Private Const TxtDirName As String = "TXT"
    Private fileReturnDate As String = Nothing
    Private errText As New System.Text.StringBuilder()
    Private intErrorCount As Integer = 0
    Private intSeq As Integer = 0
    Private Const ReturnOK As String = "RETURNOK"
    Private Const ReturnFail As String = "RETURNFAIL"
    Private Const CancelAuth As String = "CANCELAUTH"
    Private Const PostTerminal As String = "POSTTERMINAL"
    Private Const ResumeData As String = "RESUMEDATA"
    Private Const FetchOK As String = "FETCHOK"
    Private StopAll As Boolean = False
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
    Private Function getWhereCondition(ByVal dsCondition As DataSet) As String
        Dim result As String = " Where 1=1 "
        With dsCondition.Tables(0).Rows(0)
            If Not DBNull.Value.Equals(.Item("PROPDATE1")) Then
                result = result & " And PropDate >= To_date('" & .Item("PROPDATE1") & "','YYYYMMDD')"
            End If
            If Not DBNull.Value.Equals(dsCondition.Tables(0).Rows(0).Item("PROPDATE2")) Then
                result = result & " And PropDate <  To_date('" & .Item("PROPDATE2") & "','YYYYMMDD') + interval '1' day"
            End If
            If Integer.Parse(.Item("APPLYTYPE") & "") = 1 Then
                If Not DBNull.Value.Equals(.Item("STOPDATE1")) Then
                    result = result & " And StopDate >= To_date('" & .Item("STOPDATE1") & "','YYYYMMDD')"
                End If
                If Not DBNull.Value.Equals(.Item("STOPDATE2")) Then
                    result = result & " And  StopDate < To_date('" & .Item("STOPDATE2") & "','YYYYMMDD') + interval '1' day "
                End If
            End If
            If Not DBNull.Value.Equals(.Item("BANKCODE")) Then
                result = result & " And BankCode = " & .Item("BANKCODE")
            End If
            If Not DBNull.Value.Equals(.Item("ACHTNO")) Then
                Dim aryACHID = Split(.Item("ACHTNO").ToString, ",")
                Dim strWhereACHID As String = String.Empty
                For i As Integer = LBound(aryACHID) To UBound(aryACHID)
                    If aryACHID(i) <> String.Empty Then strWhereACHID = IIf(strWhereACHID <> String.Empty, strWhereACHID & " or Instr(ACHTNo,chr(39)||'" & aryACHID(i) & "'||chr(39)) > 0", "Instr(ACHTNo,chr(39)||'" & aryACHID(i) & "'||chr(39)) > 0")
                Next i
                If .Item("APPLYTYPE") = 0 OrElse .Item("APPLYTYPE") = 2 Then
                    If strWhereACHID <> String.Empty Then result = result & " And (" & strWhereACHID & " ) "
                End If
            End If
            Dim strINCD008Where As String = _DAL.GetINCD008Where(.Item("BILLHEADFMT"))
            'Dim strINCD008Where As String = "Exists(Select CitemCode From SO003 B Where " & _
            '    " 1=1 " & _
            '    " And B.CompCode=A.CompCode " & _
            '    " And instr(','||A.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " & _
            '    " And B.CitemCode In (Select CitemCode From CD068A " & _
            '    " Where Exists (Select * From CD068 Where CD068.BillHeadFmt=CD068A.BillHeadFmt " & _
            '                        " And CD068.BillHeadFmt In (" & .Item("BILLHEADFMT") & ") And  CD068.ACHType = 2 )))"
            Dim strWhere As String = String.Empty
            Dim strCancelWhere As String = String.Empty
            Select Case Integer.Parse(.Item("APPLYTYPE"))
                Case 0
                    'strWhere = " And  A.SnactionDate Is Null And A.SendDate Is Null And nvl(A.StopFlag,0) = 0 And " & strINCD008Where & _
                    '       " And Exists(Select * From   SO106A B" & _
                    '       " Where A.MasterId=B.MasterId And B.StopFlag<>1 And B.StopDate is Null And ACHTNO IN(" & .Item("INACHTNO") & "))"
                    strWhere = _DAL.GetApplyTypeWhere(strINCD008Where, .Item("INACHTNO"))
                    result = result & strWhere
                Case 1
                    'If DBNull.Value.Equals(.Item("STOPDATE1")) AndAlso DBNull.Value.Equals(.Item("STOPDATE2")) Then
                    '    strCancelWhere = IIf(strCancelWhere = String.Empty, "Nvl(A.StopFlag,0)=0", strCancelWhere & " And Nvl(A.StopFlag,0)=0")
                    'Else
                    '    strCancelWhere = IIf(strCancelWhere = String.Empty, "Nvl(A.StopFlag,0)=1", strCancelWhere & " And Nvl(A.StopFlag,0)=1")
                    'End If


                    'If strCancelWhere <> String.Empty Then strCancelWhere = " And Exists(Select * From SO106A B " & _
                    '                                                " Where A.MasterId=B.MasterId And ACHTNO IN(" & .Item("INACHTNO") & ") And " & strCancelWhere & ")"
                    result = result & _DAL.GetCancelWhere(.Item("STOPDATE1"), .Item("STOPDATE2"), .Item("INACHTNO"))
            End Select
        End With
       
        Return result
    End Function
    Private Function ChkInSO106ACitem(ByVal strCitemStr As String, ByVal strBillHeadFmt As String) As Boolean
        Dim QueryCitem As String = Nothing
        Dim arySO106ACitem As List(Of String) = strCitemStr.Split(",").ToList
        Using reader As DbDataReader = DAO.ExecDtRdr(_DAL.QueryCD068Citem(strBillHeadFmt))
            If reader.HasRows Then
                Do While reader.Read()
                    For Each item As String In arySO106ACitem
                        If item = reader.Item("CitemCode").ToString Then
                            Return True
                            Exit Do
                        End If
                    Next
                Loop
            End If
        End Using

        Return False
    End Function
    Private Function GetReStatus(ByVal strData As String) As String
        On Error Resume Next
        Select Case strData.Substring(71, 2).Trim(" ", "")
            Case "P"
                Return Language.GetReStatusP
            Case "R"
                Return Language.GetReStatusR
            Case "Y"
                Return Language.GetReStatusY
            Case "M"
                Return Language.GetReStatusM
            Case "S"
                Return Language.GetReStatusS
            Case "C"
                Return Language.GetReStatusC
            Case "D"
                Return Language.GetReStatusD
            Case ""
                Return String.Format(Language.GetReStatusOther, FNow.ToString("yyyy/MM/dd"))
        End Select
    End Function
    Private Function getCancelStatus(ByVal replyType As String)
        If replyType = CancelAuth Then
            Return Language.getCancelStatus1
        Else
            Return Language.getCancelStatus2
        End If
    End Function
    Private Function AlterSO106(ByVal alterType As String, ByVal strReadLine As String, _
                                ByVal rwSO106 As DataRow, ByVal rwCondition As DataRow, ByVal ErrorMsg As String) As Boolean
        Dim note As String = Nothing
        'Dim strINCD008Where As String = "Exists(Select CitemCode From SO003 B Where " & _
        '       " 1=1 " & _
        '       " And B.CompCode=A.CompCode " & _
        '       " And instr(','||A.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " & _
        '       " And B.CitemCode In (Select CitemCode From CD068A " & _
        '       " Where Exists (Select * From CD068 Where CD068.BillHeadFmt=CD068A.BillHeadFmt " & _
        '                           " And CD068.BillHeadFmt In (" & rwCondition.Item("BILLHEADFMT") & ") And  CD068.ACHType = 2 )))"
        Dim strINCD008Where As String = _DAL.GetINCD008Where(rwCondition.Item("BILLHEADFMT"))

        Try
            Select Case alterType.ToUpper
                Case FetchOK

                    DAO.ExecNqry(_DAL.UpdateFetchOK, New Object() {Date.ParseExact(rwCondition.Item("SENDDATE"), "yyyyMMdd", Nothing), _
                                                                    FNow.ToString("yyyyMMddHHmmss"), LoginInfo.EntryName, _
                                                                   CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow), rwSO106.Item("MasterId")})
                Case "RETURNOK".ToUpper
                    DAO.ExecNqry(_DAL.UpdateReturnOK(strINCD008Where), New Object() {fileReturnDate, _
                     GetReStatus(strReadLine), LoginInfo.EntryName, _
                    FNow.ToString("yyyyMMddHHmmss"), _
                     CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow), _
                    GetACHCustID(strReadLine), _
                    getApplyDate(strReadLine), _
                    Right("00000000000000" & strReadLine.Substring(27, 14).Trim(" ", ""), 14)})
                Case CancelAuth, PostTerminal
                    If StopAll Then
                        Dim UpdateCount As Integer = Integer.Parse(DAO.ExecNqry(_DAL.UpdateCancelAuth(), New Object() {
                                                                                        getCancelStatus(alterType),
                                                                                          Me.LoginInfo.EntryName,
                                                                                         CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                                                                                        FNow.ToString("yyyyMMddHHmmss"),
                                                                                        FNow.ToString("yyyyMMdd"),
                                                                                          GetACHCustID(strReadLine),
                                                                                          GetAccountId(strReadLine)}))


                        If Not DBNull.Value.Equals(rwSO106("Note")) Then
                            note = rwSO106("Note")
                        End If
                        If Not String.IsNullOrEmpty(note) Then
                            note = note & Environment.NewLine
                        End If
                        note = note & String.Format(Language.UpdCancelAuthNote, FNow.ToString("yyyy/MM/dd"), getCancelStatus(alterType))
                        DAO.ExecNqry(_DAL.UpdNote, New Object() {note, rwSO106("MasterId")})
                    End If
                Case ReturnFail
                    DAO.ExecNqry(_DAL.StopSO106, New Object() {Me.LoginInfo.EntryName,
                                                                           CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                                                                            ErrorMsg,
                                                                             ErrorMsg & Language.UpdNote & FNow.ToString("yyyy/MM/dd"),
                                                                            FNow.ToString("yyyyMMddHHmmss"),
                                                                            FNow.ToString("yyyyMMdd"),
                                                                             GetAccountId(strReadLine), GetACHCustID(strReadLine)})
                Case ResumeData
                    Dim strReStatus As String = Language.ResumeDataStatus                  
                    note = note & String.Format(Language.UpdCancelAuthNote, FNow.ToString("yyyy/MM/dd"), strReStatus)
                    DAO.ExecNqry(_DAL.UpdNote, New Object() {note, rwSO106("MasterId")})
                    DAO.ExecNqry(_DAL.UpdResumeData, New Object() {fileReturnDate, strReStatus, _
                                                                   LoginInfo.EntryName, _
                                                                    FNow.ToString("yyyyMMddHHmmss"), _
                                                                   CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                                                                   getApplyDate(strReadLine), GetACHCustID(strReadLine), _
                                                                   GetAccountId(strReadLine), rwSO106.Item("MasterId")})


            End Select
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function GetAccountId(ByVal ConText As String) As String
        Return ConText.Substring(27, 14).Replace(" ", "").PadLeft(14, "0")
    End Function
    Private Function WriteZip(ByVal FetchText As String, ByVal ErrorText As String, ByVal ApplyType As Integer) As String
        Dim Path As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & TxtDirName
        Dim retFileName As String = Me.LoginInfo.EntryId & "-SO3297A-" & Now.ToString("yyyyMMddHHmmssff") & ".zip"
        
        Dim AddFileName As String = Nothing
       
        If ApplyType = 0 Then
            AddFileName = ApplyFileName
        Else
            AddFileName = StopFileName
        End If
        Using zip As New Ionic.Zip.ZipFile(Path & "\" & retFileName, System.Text.Encoding.GetEncoding(950))
            If Not String.IsNullOrEmpty(FetchText) Then
                zip.AddEntry(AddFileName, FetchText)
            End If
            If Not String.IsNullOrEmpty(ErrorText) Then
                zip.AddEntry(errorFileName, ErrorText)
            End If
            zip.Save()
        End Using
        Return String.Format("{0}\{1}", TxtDirName, retFileName)
    End Function
    Private Function getApplyErrMsg(ByVal strReadLine As String) As String
        Dim retResult As String = ""
        Select Case UCase(strReadLine.Substring(71, 2).Trim(" ", ""))
            Case "03"
                retResult = Language.ApplyErrMsg03
            Case "06"
                retResult = Language.ApplyErrMsg06
            Case "07"
                retResult = Language.ApplyErrMsg07
            Case "08"
                retResult = Language.ApplyErrMsg08
            Case "09"
                retResult = Language.ApplyErrMsg09
            Case "10"
                retResult = Language.ApplyErrMsg10
            Case "11"
                retResult = Language.ApplyErrMsg11
            Case "12"
                retResult = Language.ApplyErrMsg12
            Case "13"
                retResult = Language.ApplyErrMsg13
            Case "14"
                retResult = Language.ApplyErrMsg14
            Case "16"
                retResult = Language.ApplyErrMsg16
            Case "17"
                retResult = Language.ApplyErrMsg17
            Case "18"
                retResult = Language.ApplyErrMsg18
            Case "19"
                retResult = Language.ApplyErrMsg19
            Case "91"
                retResult = Language.ApplyErrMsg91
            Case "98"
                retResult = Language.ApplyErrMsg98
        End Select
        If retResult.Length = 0 Then
            Select Case UCase(strReadLine.Substring(73, 1).Trim(" ", ""))
                Case "1"
                    retResult = Language.otherErrMsg1
                Case "2"
                    retResult = Language.otherErrMsg2
                Case "3"
                    retResult = Language.otherErrMsg3
                Case "4"
                    retResult = Language.otherErrMsg4
                Case "9"
                    retResult = Language.otherErrMsg9
            End Select
        End If
        If retResult.Length = 0 Then
            retResult = Language.noErrCode
        End If
        Return retResult
    End Function
    Private Function applyAuthFail(ByVal strData As String, ByVal tbSO106 As DataTable, ByVal rwCondition As DataRow) As Boolean
        If tbSO106.Rows.Count = 0 Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, "", strData.Substring(27, 14).Trim(" ", ""), Language.NoExistsACHCustId))
            Return False
        End If
        Dim tbSO106A As DataTable = DAO.ExecQry(_DAL.QuerySO106A(rwCondition.Item("INACHTNO").ToString, "1", rwCondition), _
                                                New Object() {tbSO106.Rows(0).Item("MasterId")})
        If tbSO106A.Rows.Count = 0 Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, tbSO106.Rows(0).Item("ID"), strData.Substring(27, 14).Trim(" ", ""), Language.NoFoundSO106A))
            Return False
        End If
        Return True
    End Function
    Private Function UpdACHSO003C(ByVal MasterId As String) As Boolean
        Try
            Using tbSO106 = DAO.ExecQry(_DAL.QueryUniqueSO106(), New Object() {MasterId})
                Dim aServiceId As String = Nothing
                If DBNull.Value.Equals(tbSO106.Rows(0).Item("ProServiceID")) OrElse String.IsNullOrEmpty(tbSO106.Rows(0).Item("ProServiceID").ToString) Then
                    aServiceId = "-99"
                Else
                    aServiceId = tbSO106.Rows(0).Item("ProServiceID")
                End If
                DAO.ExecNqry(_DAL.UpdateACHSO003C(aServiceId), New Object() {
                             tbSO106.Rows(0).Item("CMCode"),
                             tbSO106.Rows(0).Item("CMName"),
                             tbSO106.Rows(0).Item("PTCode"),
                             tbSO106.Rows(0).Item("PTName"),
                             CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                             Me.LoginInfo.EntryName,
                             FNowDate.ToString("yyyyMMddHHmmss"),
                           Integer.Parse(MasterId)})
                tbSO106.Dispose()
            End Using
        Catch ex As Exception
            Return False

        End Try
        Return True
    End Function
    Private Function UpdSO003C(ByVal MasterId As String) As Boolean
        Try
            Using tbSO106 = DAO.ExecQry(_DAL.QueryUniqueSO106(), New Object() {Integer.Parse(MasterId)})
                DAO.ExecNqry(_DAL.UpdateSO003C(), New Object() {
                             tbSO106.Rows(0).Item("CMCode"),
                             tbSO106.Rows(0).Item("CMName"),
                             tbSO106.Rows(0).Item("PTCode"),
                             tbSO106.Rows(0).Item("PTName"),
                             CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                             Me.LoginInfo.EntryName,
                             FNowDate.ToString("yyyyMMddHHmmss"),
                            Integer.Parse(MasterId)})
                tbSO106.Dispose()
            End Using


        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Public Function CanEdit() As RIAResult
        Dim retRia As New RIAResult With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        If Integer.Parse(DAO.ExecSclr(_DAL.QueryCanEdit, New Object() {LoginInfo.CompCode})) = 0 Then
            retRia.ResultBoolean = False
            retRia.ErrorMessage = Language.NoCanedit
            retRia.ErrorCode = -1
        End If
        Return retRia
    End Function
    Public Function ChkAuthority(ByVal Mid As String) As RIAResult
        Dim result As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        Try
            Using obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
                result = obj.ChkPriv(LoginInfo.EntryId, Mid)
                obj.Dispose()
            End Using
            If Not result.ResultBoolean Then
                result.ErrorCode = -1
                result.ErrorMessage = Language.NoPermission
            End If
            'If Integer.Parse(DAO.ExecSclr(_DAL.chkAuthority(Me.LoginInfo.GroupId), New Object() {Mid})) = 0 Then
            '    result.ResultBoolean = False
            '    result.ErrorCode = -1
            '    result.ErrorMessage = Language.NoPermission
            '    Return result
            'End If

        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -2
        Finally

        End Try
        Return result

    End Function
    Public Function GetCompCode() As DataTable
        Try
            If Me.LoginInfo.GroupId = "0" AndAlso 1 = 0 Then
                Return DAO.ExecQry(_DAL.GetCompCode("0",
                                                    CableSoft.BLL.Utility.Utility.GetCompanyTableName(Me.LoginInfo, Me.DAO),
                                                       CableSoft.BLL.Utility.Utility.GetLoginTableName))
            Else
                Return DAO.ExecQry(_DAL.GetCompCode("1",
                                                    CableSoft.BLL.Utility.Utility.GetCompanyTableName(Me.LoginInfo, Me.DAO),
                                                       CableSoft.BLL.Utility.Utility.GetLoginTableName),
                                   New Object() {Me.LoginInfo.EntryId})
            End If
        Catch ex As Exception
            Throw
        End Try

    End Function
    Private Function UpdNonePeriod(ByVal MasterId As String) As Boolean
        Dim tbSO106 As DataTable = Nothing
        Try
            tbSO106 = DAO.ExecQry(_DAL.QueryUniqueSO106(), New Object() {
                                    Integer.Parse(MasterId)})

            If tbSO106.Rows.Count > 0 Then
                DAO.ExecNqry(_DAL.UpdNonePeriod(tbSO106.Rows(0).Item("CitemStr")), New Object() {
                             tbSO106.Rows(0).Item("BankCode"),
                             tbSO106.Rows(0).Item("BANKNAME"),
                             tbSO106.Rows(0).Item("ACCOUNTID"),
                             tbSO106.Rows(0).Item("PTCode"),
                             tbSO106.Rows(0).Item("PTName"),
                             tbSO106.Rows(0).Item("CMCode"),
                             tbSO106.Rows(0).Item("CMName"),
                             Me.LoginInfo.EntryName,
                            CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                             FNowDate.ToString("yyyyMMddHHmmss"),
                             tbSO106.Rows(0).Item("ID")})

            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        Finally
            If tbSO106 IsNot Nothing Then
                tbSO106.Dispose()
                tbSO106 = Nothing
            End If
        End Try
        Return True
    End Function
    Private Overloads Function UpdateSO003(ByVal masterId As String) As Boolean
        Dim tbSO106 As DataTable = Nothing
        Try
            tbSO106 = DAO.ExecQry(_DAL.QueryUniqueSO106(), New Object() {
                                   Integer.Parse(masterId)})

            If tbSO106.Rows.Count > 0 Then
                DAO.ExecNqry(_DAL.UpdateSO003, New Object() {
                             tbSO106.Rows(0).Item("BankCode"),
                             tbSO106.Rows(0).Item("BANKNAME"),
                             tbSO106.Rows(0).Item("ACCOUNTID"),
                             tbSO106.Rows(0).Item("PTCode"),
                             tbSO106.Rows(0).Item("PTName"),
                             tbSO106.Rows(0).Item("CMCode"),
                             tbSO106.Rows(0).Item("CMName"),
                             Me.LoginInfo.EntryName,
                            CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                             FNowDate.ToString("yyyyMMddHHmmss"),
                             Integer.Parse(masterId)})

            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        Finally
            If tbSO106 IsNot Nothing Then
                tbSO106.Dispose()
                tbSO106 = Nothing
            End If
        End Try
        Return True
    End Function
    Private Function StopACHSO003C(ByVal MasterId As String, ByVal CMCode As Integer, ByVal CMName As String,
                                         ByVal PTCode As Integer, ByVal PTName As String) As Boolean
        Try
            Using tbSO106 = DAO.ExecQry(_DAL.QueryUniqueSO106(), New Object() {Integer.Parse(MasterId)})
                Dim aServiceId As String = Nothing
                If DBNull.Value.Equals(tbSO106.Rows(0).Item("ProServiceID")) OrElse String.IsNullOrEmpty(tbSO106.Rows(0).Item("ProServiceID").ToString) Then
                    aServiceId = "-99"
                Else
                    aServiceId = tbSO106.Rows(0).Item("ProServiceID")
                End If
                DAO.ExecNqry(_DAL.UpdateACHSO003C(aServiceId), New Object() {
                         CMCode,
                         CMName,
                         PTCode,
                         PTName,
                         CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                         Me.LoginInfo.EntryName,
                         FNowDate.ToString("yyyyMMddHHmmss"),
                        DBNull.Value})
                tbSO106.Dispose()
            End Using


        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Function FileReturn(ByVal TakenFileName As String, ByVal dsCondition As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim applyNumber As Integer = -1
        Dim blnAutoClose As Boolean = False
        Dim result As New RIAResult() With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = Language.chkFileFail, .ResultXML = Language.chkFileFail}
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim CurrentDir As String = String.Format("{0}\Txt\", CableSoft.BLL.Utility.Utility.GetCurrentDirectory())
        Dim aryConText As List(Of String) = System.IO.File.ReadAllLines(String.Format("{0}\{1}", CurrentDir, TakenFileName)).ToList
        Dim blnChkFileOK As Boolean = False
      
        'Dim aryConText As List(Of String) = oContext.Split(Environment.NewLine).ToList
        If trans IsNot Nothing Then
            trans.Dispose()
            trans = Nothing
        End If
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction

        Else
            If cn IsNot Nothing AndAlso cn.State <> ConnectionState.Open Then
                cn.ConnectionString = Me.LoginInfo.ConnectionString
                cn.Open()
            End If
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        If blnAutoClose Then
            CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, Language.ClientInfoString)
        End If
        Try
            For Each conText As String In aryConText.AsEnumerable().Reverse()
                If Not String.IsNullOrEmpty(conText) Then
                    If conText.Substring(0, 1) = "2" Then
                        fileReturnDate = (conText.Substring(26, 8) & "").Trim(" ", "")
                        If String.IsNullOrEmpty(fileReturnDate) OrElse fileReturnDate.Length <> 8 OrElse (Not IsNumeric(fileReturnDate)) Then
                            fileReturnDate = Date.Now.ToString("yyyyMMdd")
                        End If
                        If Integer.Parse("0" & conText.Substring(34, 6).Trim(" ", "")) > 0 Then
                            blnChkFileOK = True
                            Exit For
                        End If
                        If Integer.Parse("0" & conText.Substring(40, 6).Trim(" ", "")) > 0 Then
                            blnChkFileOK = True
                            Exit For
                        End If
                    End If
                End If
            Next
            If Not blnChkFileOK Then
                Return result
            End If
            result.ErrorMessage = Nothing
            result.ErrorCode = 0
            For Each conText As String In aryConText
                If Integer.Parse("0" & conText.Substring(0, 1)) <> 1 Then
                    Continue For
                End If
                applyNumber = Integer.Parse(conText.Substring(25, 1))
                Dim strSQL As String = getApplySO106SQL(conText)
                Dim tbSO106A As DataTable = Nothing


                Using tbSO106 As DataTable = DAO.ExecQry(strSQL)
                    If tbSO106.Rows.Count > 0 Then
                        tbSO106A = DAO.ExecQry(_DAL.QuerySO106A(dsCondition.Tables(0).Rows(0).Item("INACHTNO").ToString, conText.Substring(25, 1), dsCondition.Tables(0).Rows(0)), _
                                              New Object() {tbSO106.Rows(0).Item("MasterId")})
                        If tbSO106A.Rows.Count = 0 Then
                            intErrorCount += 1
                            errText.AppendLine(String.Format(Language.CommonErr, tbSO106.Rows(0).Item("ID"), conText.Substring(27, 14).Trim(" ", ""), Language.NoFoundSO106A))
                            Continue For
                        End If
                    Else
                        intErrorCount += 1
                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.NoExistsACHCustId))
                        Continue For
                    End If

                    Select Case applyNumber
                        Case 1
                            Try
                                'Return OK

                                If Integer.Parse("0" & conText.Substring(71, 2).Trim(" ", "")) = 0 AndAlso Integer.Parse("0" & conText.Substring(73, 1).Trim(" ", "")) = 0 Then
                                    If Not applyAuthOK(tbSO106, tbSO106A, conText, dsCondition.Tables(0).Rows(0)) Then
                                        'trans.Rollback()
                                        Continue For
                                    End If

                                    If Not AlterSO106("RETURNOK", conText, tbSO106.Rows(0), dsCondition.Tables(0).Rows(0), Nothing) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO106))
                                        Continue For
                                    End If

                                    If Not UpdACHSO003C(tbSO106.Rows(0).Item("MasterId")) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO003CError))
                                        Continue For
                                    End If
                                    If Not UpdateSO003(tbSO106.Rows(0).Item("Masterid")) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO003Error))
                                        Continue For
                                    End If
                                    If Not UpdNonePeriod(tbSO106.Rows(0).Item("MasterId")) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdNoneSO003Error))
                                        Continue For
                                    End If
                                    If Not UpdateSO106A(ReturnOK, tbSO106A, Nothing) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO106AError))
                                        Continue For
                                    End If
                                    intSeq += 1
                                Else
                                    'Return Failure
                                    Dim retErrMsg As String = getApplyErrMsg(conText)
                                    If Not applyAuthFail(conText, tbSO106, dsCondition.Tables(0).Rows(0)) Then Continue For
                                    If Not UpdateSO106A(ReturnFail, tbSO106A.Copy, retErrMsg) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO106AError))
                                        Continue For
                                    End If
                                    If Not AlterSO106(ReturnFail, conText, tbSO106.Rows(0), dsCondition.Tables(0).Rows(0), retErrMsg) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO106))
                                        Continue For
                                    End If
                                    If Not ClearSO106(tbSO106.Rows(0)) Then
                                        'trans.Rollback()
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.ClearSO106Fail))
                                        Continue For
                                    End If
                                    intSeq += 1
                                End If
                                tbSO106.Dispose()

                            Catch ex As Exception
                                trans.Rollback()
                            Finally
                                If tbSO106A IsNot Nothing Then
                                    tbSO106A.Dispose()
                                    tbSO106A = Nothing
                                End If

                            End Try
                        Case 2, 3
                            If Not StopAuth(conText, tbSO106, tbSO106A, dsCondition.Tables(0).Rows(0)) Then
                                'trans.Rollback()
                                Continue For
                            End If
                        Case 4
                            If Not applyAuthOK(tbSO106, tbSO106A, conText, dsCondition.Tables(0).Rows(0)) Then
                                'trans.Rollback()
                                Continue For
                            End If
                            If Not AlterSO106(ResumeData, conText, tbSO106.Rows(0), dsCondition.Tables(0).Rows(0), Nothing) Then
                                'trans.Rollback()
                                intErrorCount += 1
                                errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO106))
                                Continue For
                            End If
                            If Not ReCoverSO106(tbSO106, tbSO106A) Then
                                'trans.Rollback()
                                intErrorCount += 1
                                errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.ResumeCitemCode))
                                Continue For
                            End If
                            If Not UpdACHSO003C(tbSO106.Rows(0).Item("MasterId")) Then
                                'trans.Rollback()
                                intErrorCount += 1
                                errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO003CError))
                                Continue For
                            End If
                            If Not UpdateSO003(tbSO106.Rows(0).Item("Masterid")) Then
                                'trans.Rollback()
                                intErrorCount += 1
                                errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO003Error))
                                Continue For
                            End If
                            If Not UpdNonePeriod(tbSO106.Rows(0).Item("MasterId")) Then
                                'trans.Rollback()
                                intErrorCount += 1
                                errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdNoneSO003Error))
                                Continue For
                            End If

                            If Not UpdateSO106A(ReturnOK, tbSO106A, Nothing) Then
                                'trans.Rollback()
                                intErrorCount += 1
                                errText.AppendLine(String.Format(Language.CommonErr, "", conText.Substring(27, 14).Trim(" ", ""), Language.UpdSO106AError))
                                Continue For
                            End If
                            intSeq += 1

                    End Select
                End Using


            Next
            trans.Commit()           
            result.ErrorCode = 0
            result.ErrorMessage = Nothing
            result.ResultBoolean = True
            result.ResultXML = String.Format(Language.resultCount, intSeq, intErrorCount)
            If intErrorCount > 0 Then
                result.DownloadFileName = WriteZip("", errText.ToString, -1)
            End If
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
            End If
            
        End Try
        
        Return result
    End Function
    Private Function SearchSameChar(ByVal SourceWord As String, SearchWord As String) As Boolean
        For Each Str As String In SourceWord.Split(",")
            If Str = SearchWord Then Return True
        Next
        Return False
    End Function
    Private Function ReCoverSO106(ByVal tbSO106 As DataTable, ByVal tbSO106A As DataTable) As Boolean

        Try
            Dim OringalACHTNo As String = String.Empty
            Dim UpdACHTNo As String = String.Empty
            Dim OringalACHTDesc As String = String.Empty
            Dim UpdACHTDesc As String = String.Empty
            Dim OringalCitemStr As String = String.Empty
            Dim UpdCitemStr As String = String.Empty
            If Not DBNull.Value.Equals(tbSO106.Rows(0).Item("ACHTNo")) Then
                OringalACHTNo = tbSO106.Rows(0).Item("ACHTNo")
                UpdACHTNo = OringalACHTNo
            End If
            If Not DBNull.Value.Equals(tbSO106.Rows(0).Item("ACHTDESC")) Then
                OringalACHTDesc = tbSO106.Rows(0).Item("ACHTDESC")
                UpdACHTDesc = OringalACHTDesc
            End If
            If Not DBNull.Value.Equals(tbSO106.Rows(0).Item("CitemStr")) Then
                OringalCitemStr = tbSO106.Rows(0).Item("CitemStr")
                UpdCitemStr = OringalCitemStr
            End If
            For Each rwSO106A As DataRow In tbSO106A.Rows
                If Not SearchSameChar(OringalACHTNo, "'" & rwSO106A.Item("ACHTNO") & "'") Then
                    If String.IsNullOrEmpty(UpdACHTNo) Then
                        UpdACHTNo = "'" & rwSO106A.Item("ACHTNO") & "'"
                    Else
                        UpdACHTNo = UpdACHTNo & ",'" & rwSO106A.Item("ACHTNO") & "'"
                    End If
                End If
                If Not SearchSameChar(OringalACHTDesc, "'" & rwSO106A.Item("ACHDesc") & "'") Then
                    If String.IsNullOrEmpty(UpdACHTDesc) Then
                        UpdACHTDesc = "'" & rwSO106A.Item("ACHDesc") & "'"
                    Else
                        UpdACHTDesc = UpdACHTDesc & ",'" & rwSO106A.Item("ACHDesc") & "'"
                    End If
                End If


                Using rd As DbDataReader = DAO.ExecDtRdr(_DAL.QuerySO003SEQNO(rwSO106A.Item("CitemCodeStr")),
                                                                              New Object() {LoginInfo.CompCode, tbSO106.Rows(0).Item("MasterId")})
                    If rd.HasRows Then
                        Do While rd.Read()
                            If Not SearchSameChar(UpdCitemStr, "'" & rd.Item("SEQNO") & "'") Then
                                If String.IsNullOrEmpty(UpdCitemStr) Then
                                    UpdCitemStr = "'" & rd.Item("SEQNO") & "'"
                                Else
                                    UpdCitemStr = UpdCitemStr & ",'" & rd.Item("SEQNO") & "'"
                                End If
                            End If
                        Loop

                    End If
                End Using


                Using rd2 As DbDataReader = DAO.ExecDtRdr(_DAL.QueryNonePeriodSEQNo(rwSO106A.Item("CitemCodeStr")),
                                                        New Object() {tbSO106.Rows(0).Item("ID")})
                    If rd2.HasRows Then
                        Do While rd2.Read()
                            If Not SearchSameChar(UpdCitemStr, "'" & rd2.Item("SEQNO") & "'") Then
                                If String.IsNullOrEmpty(UpdCitemStr) Then
                                    UpdCitemStr = "'" & rd2.Item("SEQNO") & "'"
                                Else
                                    UpdCitemStr = UpdCitemStr & ",'" & rd2.Item("SEQNO") & "'"
                                End If
                            End If
                        Loop
                    End If
                End Using

            Next

            DAO.ExecNqry(_DAL.ResumeACH, New Object() {UpdACHTNo, UpdACHTDesc, UpdCitemStr, tbSO106.Rows(0).Item("MasterId")})
        Catch ex As Exception
            Return False
        End Try

        Return True
        'Dim sql As String
        'Dim rsSO106Upd As New Recordset
        'Dim rsSO003 As New ADODB.Recordset
        'Dim citemCode As String
        'sql = "Select * From " & GetOwner & "SO106 A Where Masterid = " & rsSO106A("MasterId") & " And CompCode = " & gilCompCode.GetCodeNo
        'If Not GetRS(rsSO106Upd, sql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        'If InStr(1, rsSO106Upd("ACHTNo"), "'" & rsSO106A("ACHTNo") & "'") <= 0 Then
        '    rsSO106Upd("ACHTNo") = IIf(Len(rsSO106Upd("ACHTNo") & "") = 0, "'" & rsSO106A("ACHTNo") & "'", rsSO106Upd("ACHTNo") & ",'" & rsSO106A("ACHTNo") & "'")
        '    rsSO106Upd("ACHTDESC") = IIf(Len(rsSO106Upd("ACHTDESC") & "") = 0, "'" & rsSO106A("ACHDesc") & "'", rsSO106Upd("ACHDesc") & ",'" & rsSO106A("ACHDesc") & "'")
        '    rsSO106Upd.Update()
        'End If
        'citemCode = rsSO106A("CitemCodeStr") & ""
        'If Len(citemCode) > 0 Then
        '    sql = "Select SeqNo From " & GetOwner & "SO003 " & _
        '        " Where Custid = " & strCustId & " And CitemCode In ( " & citemCode & " ) " & _
        '        " And CompCode = " & gilCompCode.GetCodeNo
        '    If Not GetRS(rsSO003, sql, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        '    If rsSO003.RecordCount > 0 Then
        '        rsSO003.MoveFirst()
        '        rsSO106Upd.MoveFirst()
        '        Do While Not rsSO003.EOF
        '            If InStr(1, rsSO106Upd("CitemStr"), "'" & rsSO003("SeqNo") & "'") <= 0 Then
        '                rsSO106Upd("CitemStr") = IIf(Len(rsSO106Upd("CitemStr") & "") = 0, "'" & rsSO003("SeqNo") & "'", rsSO106Upd("CitemStr") & ",'" & rsSO003("SeqNo") & "'")
        '                rsSO106Upd.Update()
        '            End If
        '            rsSO003.MoveNext()
        '        Loop
        '    End If
        'End If
        'ReCoverSO106 = True
    End Function
    Private Function StopAuth(ByVal strReadLine As String, ByVal tbSO106 As DataTable, ByVal tbSO106A As DataTable, ByVal rwCondition As DataRow) As Boolean
        Dim RepyType As String = Nothing

        If Integer.Parse("0" & strReadLine.Substring(25, 1)) = 2 Then
            RepyType = CancelAuth
        Else
            RepyType = PostTerminal
        End If
        If Not UpdateSO106A(RepyType, tbSO106A, Nothing) Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, "", strReadLine.Substring(27, 14).Trim(" ", ""), Language.UpdSO106AError))
            Return False
        End If
        If Not AlterSO106(RepyType, strReadLine, tbSO106.Rows(0), rwCondition, Nothing) Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, "", strReadLine.Substring(27, 14).Trim(" ", ""), Language.UpdSO106))
            Return False
        End If
        Dim aPTCode As Integer = tbSO106.Rows(0).Item("PTCode")
        Dim aPTName As String = tbSO106.Rows(0).Item("PTName")
        Dim aCMCode As Integer = tbSO106.Rows(0).Item("CMCode")
        Dim aCMName As String = tbSO106.Rows(0).Item("CMName")

        Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetPTCode)
            dr.Read()
            aPTCode = dr.Item("CODENO")
            aPTName = dr.Item("Description")

        End Using


        Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefCMCode(LoginInfo, String.Empty))
            dr.Read()
            aCMCode = dr.Item("CODENO")
            aCMName = dr.Item("Description")
        End Using
        If Not StopSO003(aCMCode, aCMName, aPTCode, aPTName, tbSO106.Rows(0)) Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, "", strReadLine.Substring(27, 14).Trim(" ", ""), Language.StopSO003Error))
            Return False
        End If
        If Not StopNonePeriod(aCMCode, aCMName, aPTCode, aPTName, tbSO106.Rows(0)) Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, "", strReadLine.Substring(27, 14).Trim(" ", ""), Language.StopNonePeriodError))
            Return False
        End If
        If Not StopACHSO003C(tbSO106.Rows(0).Item("MasterId"), aCMCode, aCMName, aPTCode, aPTName) Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, "", strReadLine.Substring(27, 14).Trim(" ", ""), Language.StopSO003CError))
            Return False
        End If
        
       
        intSeq += 1
        Return True
    End Function
    Private Function StopSO003C(ByVal MasterId As String, ByVal CMCode As Integer, ByVal CMName As String,
                                         ByVal PTCode As Integer, ByVal PTName As String) As Boolean
        Try

            DAO.ExecNqry(_DAL.UpdateSO003C(), New Object() {
                         CMCode,
                         CMName,
                         PTCode,
                         PTName,
                         CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                         Me.LoginInfo.EntryName,
                         FNowDate.ToString("yyyyMMddHHmmss"),
                       Integer.Parse(MasterId)})

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Private Function StopNonePeriod(ByVal CMCode As Integer, ByVal CMName As String,
                                         ByVal PTCode As Integer, ByVal PTName As String,
                                         ByVal rwSO106 As DataRow) As Boolean
        Try
            DAO.ExecNqry(_DAL.StopNonePeriod(rwSO106.Item("CitemStr")), New Object() {
                         CMCode, CMName, PTCode, PTName,
                        Me.LoginInfo.EntryName,
                          CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                        FNowDate.ToString("yyyyMMddHHmmss"),
                         rwSO106.Item("ACCOUNTID"), rwSO106.Item("ID")})
        Catch ex As Exception
            Throw
        Finally

        End Try
        Return True
    End Function
    Private Overloads Function StopSO003(ByVal CMCode As Integer, ByVal CMName As String,
                                         ByVal PTCode As Integer, ByVal PTName As String,
                                         ByVal rwSO106 As DataRow) As Boolean
        Try
            DAO.ExecNqry(_DAL.StopSO003, New Object() {
                         CMCode, CMName, PTCode, PTName,
                        Me.LoginInfo.EntryName,
                          CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                        FNowDate.ToString("yyyyMMddHHmmss"),
                         Me.LoginInfo.CompCode,
                         rwSO106.Item("MasterId"),
                         rwSO106.Item("ACCOUNTID")})
        Catch ex As Exception
            Throw
        Finally

        End Try
        Return True
    End Function
    Private Function ClearSO106(ByVal rwSO106 As DataRow) As Boolean
        Dim OriginalACHTNo As String = Nothing
        Dim OriginalACHTDesc As String = Nothing
        Dim OriginalCitemStr As String = Nothing
        Dim tbSO003 As DataTable = Nothing
        If Not DBNull.Value.Equals(rwSO106("ACHTNo")) Then
            OriginalACHTNo = rwSO106("ACHTNo").ToString
        End If
        If Not DBNull.Value.Equals(rwSO106("ACHTDesc")) Then
            OriginalACHTDesc = rwSO106("ACHTDesc").ToString
        End If
        If Not DBNull.Value.Equals(rwSO106("CitemStr")) Then
            OriginalCitemStr = rwSO106("CitemStr").ToString
        End If
        Try
            If Not String.IsNullOrEmpty(OriginalCitemStr) Then
                'If SourceType = FormatType.NewType Then
                '    tbSO003 = DAO.ExecQry(_DAL.QuerySO003(rwSO106A("CitemCodeStr").ToString),
                '                                    New Object() {rwSO106.Item("CustId"),
                '                                                  Me.LoginInfo.CompCode})

                'Else
                '    tbSO003 = DAO.ExecQry(_DAL.QuerySO003,
                '                              New Object() {Me.LoginInfo.CompCode,
                '                                            rwSO106.Item("masterid")})
                'End If
                tbSO003 = DAO.ExecQry(_DAL.QuerySO003,
                                              New Object() {Me.LoginInfo.CompCode,
                                                            rwSO106.Item("masterid")})
                For Each rwSO003 As DataRow In tbSO003.Rows
                    OriginalCitemStr = OriginalCitemStr.Replace(",'" & rwSO003("SEQNO").ToString & "'", "").Replace("'" & rwSO003("SEQNO").ToString & "'", "")
                Next
            End If
            If Not String.IsNullOrEmpty(OriginalCitemStr) Then
                OriginalCitemStr = OriginalCitemStr.TrimStart(",").TrimEnd(",")
            End If
            'disccussing debby with the issue to decide that it does't necessary to clear achtno by kin 
            'If SourceType = FormatType.NewType Then
            '    If Not String.IsNullOrEmpty(OriginalACHTNo) Then
            '        OriginalACHTNo = OriginalACHTNo.Replace(",'" & rwSO106A("ACHTNO") & "'", "").Replace(rwSO106A("ACHTNO"), "")
            '    End If
            '    If Not String.IsNullOrEmpty(OriginalACHTNo) Then
            '        OriginalACHTNo = OriginalCitemStr.TrimStart(",").TrimEnd(",")
            '    End If
            'Else
            '    OriginalACHTNo = String.Empty
            '    OriginalACHTDesc = String.Empty
            'End If

            DAO.ExecNqry(_DAL.ClearSO106, New Object() {
                         OriginalACHTNo, OriginalACHTDesc, OriginalCitemStr, rwSO106("MasterId")})

        Catch ex As Exception
            Throw
        Finally
            If tbSO003 IsNot Nothing Then
                tbSO003.Dispose()
                tbSO003 = Nothing
            End If
        End Try
        Return True
    End Function
    Private Function UpdateSO106A(ByVal ReplyType As String,
                                  ByVal tbSO106A As DataTable, ByVal ErrMsg As String) As Boolean
        Dim ErrNote As String = Nothing

        Try
            Select Case ReplyType
                Case ReturnOK
                    For Each rwSO106A As DataRow In tbSO106A.Rows
                        'DAO.ExecNqry(String.Format("Update SO106A Set AuthorizeStatus = 1 Where RowId = {0}0", _DAL.Sign), New Object() {rwSO106A("RowId")})
                        'DAO.ExecNqry(_DAL.UpdateSO106AStatus(ReturnOK), _
                        '             New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow), Me.LoginInfo.EntryName, _
                        '                                                              rwSO106A("RowId")})
                        DAO.ExecNqry(_DAL.UpdateSO106AStatus(ReturnOK),
                                     New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow), Me.LoginInfo.EntryName,
                                                                                      rwSO106A("ctid")})
                    Next

                Case CancelAuth
                    For Each rwSO106A In tbSO106A.Rows
                        'DAO.ExecNqry(_DAL.UpdateSO106AStatus(CancelAuth), _
                        '             New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow), Me.LoginInfo.EntryName, _
                        '                                                                rwSO106A.Item("RowId")})
                        DAO.ExecNqry(_DAL.UpdateSO106AStatus(CancelAuth),
                                     New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow), Me.LoginInfo.EntryName,
                                                                                        rwSO106A.Item("ctid")})
                        Using tbStopAll As DataTable = DAO.ExecQry(_DAL.ChkCancelAuthStopAll, _
                                                                   New Object() {rwSO106A.Item("MasterId")})
                            If tbStopAll.Rows.Count > 0 Then
                                With tbStopAll.Rows(0)
                                    If (.Item("A") = 0) AndAlso (.Item("C") = .Item("B")) Then
                                        StopAll = True
                                    Else
                                        StopAll = False
                                    End If
                                End With
                            Else
                                StopAll = False
                            End If
                        End Using
                    Next
                Case PostTerminal
                    For Each rwSO106A In tbSO106A.Rows
                        'DAO.ExecNqry(_DAL.UpdateSO106AStatus(CancelAuth), _
                        '             New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow), _
                        '                           Me.LoginInfo.EntryName, _
                        '                           rwSO106A.Item("RowId")})
                        DAO.ExecNqry(_DAL.UpdateSO106AStatus(CancelAuth),
                                     New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                                                   Me.LoginInfo.EntryName,
                                                   rwSO106A.Item("ctid")})
                        Using tbStopAll As DataTable = DAO.ExecQry(_DAL.ChkCancelAuthStopAll, _
                                                                   New Object() {rwSO106A.Item("MasterId")})
                            If tbStopAll.Rows.Count > 0 Then
                                With tbStopAll.Rows(0)
                                    If (.Item("A") = 0) AndAlso (.Item("D") = 0) Then
                                        StopAll = True
                                    Else
                                        StopAll = False
                                    End If
                                End With
                            Else
                                StopAll = False
                            End If
                        End Using
                    Next

                Case ReturnFail
                    For Each rwSO106A In tbSO106A.Rows

                        'DAO.ExecNqry(_DAL.UpdateSO106AStatus(ReturnFail), New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                        '                             Me.LoginInfo.EntryName,
                        '                             ErrMsg & Language.UpdNote & FNowDate.ToShortDateString, rwSO106A.Item("RowId")})
                        DAO.ExecNqry(_DAL.UpdateSO106AStatus(ReturnFail), New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                                                     Me.LoginInfo.EntryName,
                                                     ErrMsg & Language.UpdNote & FNowDate.ToShortDateString, rwSO106A.Item("ctid")})
                        If Integer.Parse(DAO.ExecSclr(_DAL.ChkSO106AAllFail,
                                   New Object() {rwSO106A.Item("AchtNO"),
                                                rwSO106A.Item("ACHDesc"),
                                                rwSO106A.Item("MasterId")})) = 0 Then

                            'StopAll = True
                            Using tbSO106A2 As DataTable = DAO.ExecQry(_DAL.QuerySO106AErrNote,
                                                                      New Object() {rwSO106A.Item("MasterId")})
                                For Each rw As DataRow In tbSO106A2.Rows
                                    If String.IsNullOrEmpty(ErrNote) Then
                                        ErrNote = String.Format("{0}:{1}", rw("ACHDesc"), rw("Notes"))
                                    Else
                                        ErrNote = ErrNote & Environment.NewLine & String.Format("{0}:{1}", rw("ACHDesc"), rw("Notes"))
                                    End If
                                Next
                            End Using
                            If Not String.IsNullOrEmpty(ErrNote) Then
                                DAO.ExecNqry(_DAL.UpdateSO106Note,
                                             New Object() {ErrNote, rwSO106A.Item("MasterId")})

                            End If
                        Else
                            'StopAll = False
                        End If
                    Next






                    'Case AuthType.OldAuth
                    '    For Each rwSO106A In tblSO106A.Rows
                    '        If Not String.IsNullOrEmpty(OldAuthNewRowId) Then
                    '            With rwSO106A
                    '                DAO.ExecNqry(_DAL.InsertSO106A, New Object() {
                    '                       OldAuthNewRowId,
                    '                       .Item("ACHTNO"),
                    '                       .Item("Notes"),
                    '                       .Item("CitemCodeStr"),
                    '                       .Item("CitemNameStr"),
                    '                       .Item("StopFlag"),
                    '                       .Item("StopDate"),
                    '                       Me.LoginInfo.EntryName,
                    '                       CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                    '                       Me.LoginInfo.EntryName,
                    '                       .Item("RecordType"),
                    '                       .Item("AuthorizeStatus"),
                    '                       .Item("AchDesc"),
                    '                       NewMasterIdSeq})
                    '            End With

                    '        End If
                    '    Next

            End Select


        Catch ex As Exception            
            Return False
        Finally

        End Try

        Return True
    End Function
    Private Function applyAuthOK(ByVal tbSO106 As DataTable, ByVal tbSO106A As DataTable, ByVal strData As String, ByVal rwCondition As DataRow) As Boolean
       
        If tbSO106.Rows.Count = 0 Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, "", strData.Substring(27, 14).Trim(" ", ""), Language.NoExistsACHCustId))
            Return False
        End If
        'Dim tbSO106A As DataTable = DAO.ExecQry(_DAL.QuerySO106A(rwCondition.Item("INACHTNO").ToString, strData.Substring(25, 1)), _
        '                                        New Object() {tbSO106.Rows(0).Item("MasterId")})
        If tbSO106A Is Nothing OrElse tbSO106A.Rows.Count = 0 Then
            intErrorCount += 1
            errText.AppendLine(String.Format(Language.CommonErr, tbSO106.Rows(0).Item("ID"), strData.Substring(27, 14).Trim(" ", ""), Language.NoFoundSO106A))
            Return False
        End If
        
        Return True
    End Function
    Private Function getApplyDate(strReadLine As String) As String
        On Error Resume Next
        Dim Ret As String = strReadLine.Substring(8, 8).Trim(" ", "")
        If String.IsNullOrEmpty(Ret) OrElse (Ret.Length <> 8) OrElse (Not IsNumeric(Ret)) Then
            Ret = Date.Now.ToString("yyyyMMdd")
        End If
        Return Ret
    
    End Function
    Private Function GetACHCustID(strReadLine As String) As String
        On Error Resume Next

        Return strReadLine.Substring(41, 20).Trim(" ", "")
    End Function
    Private Function getApplySO106SQL(ByVal strData As String) As String
        Dim applyNumber As Integer
        Dim aWhere As String
        Dim result As String = Nothing
        aWhere = ""
        applyNumber = Val(Mid(strData, 26, 1) & "")
        Select Case applyNumber
            Case 1
                aWhere = " And StopFlag <> 1 "
            Case 4
                aWhere = " And StopFlag = 1 "
        End Select
        result = _DAL.getApplySO106SQL(GetACHCustID(strData), Right("00000000000000" & strData.Substring(27, 14), 14).Trim(" ", ""),
                                Left(strData.Substring(61, 10).Trim(" ", "") & "0000000000", 10), aWhere)
        'result = "Select RowId,A.* From SO106 A Where ACHCUSTID='" & GetACHCustID(strData) & _
        '                         "' And LPAD(AccountID,14,'0')='" & Right("00000000000000" & strData.Substring(27, 14), 14).Trim(" ", "") & "'" & _
        '                         " And RPAD(nvl(AccountNameID,'0'),10,'0')='" & Left(strData.Substring(61, 10).Trim(" ", "") & "0000000000", 10) & "'" & aWhere
        Return result
    End Function
    Public Function FetchData(ByVal dsCondition As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim tbLogSO106 As DataTable = Nothing
        Dim strWhere As String = getWhereCondition(dsCondition)
        Dim fetchText As New System.Text.StringBuilder()

        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = Language.NoData}
        Dim returnFile As String
        Dim RunTime As New Stopwatch()
        RunTime.Start()
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            If cn.State = ConnectionState.Closed Then
                cn.ConnectionString = Me.LoginInfo.ConnectionString
                cn.Open()
            End If
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId)
        'Dim strSQL = "SELECT A.RowID,A.Custid,A.BankCode,AccountID,A.AccountNameId,A.CitemStr,A.ACHTNo,A.ACHSN,A.ACHCUSTID,A.MasterId," & _
        '                " (Select BankId From CD018 Where CD018.CodeNo = A.BankCode And RowNum = 1 ) BankID " & _
        '                " From SO106 A" & strWhere

        Dim strSQL As String = _DAL.FetchData(strWhere)
        Try
            Using tbSO106 = DAO.ExecQry(strSQL)
                With dsCondition.Tables(0).Rows(0)
                    For Each rwSO106 As DataRow In tbSO106.Rows
                        Dim tbSO106A As DataTable = Nothing
                        Select Case Integer.Parse(.Item("APPLYTYPE") & "")
                            Case 0
                                tbSO106A = DAO.ExecQry(_DAL.FetchApplySO106A(.Item("INACHTNO").ToString, .Item("INACHDESC").ToString), _
                                                    New Object() {rwSO106("MasterId")})
                            Case 1
                                tbSO106A = DAO.ExecQry(_DAL.FetchStopSO106A(.Item("INACHTNO").ToString, .Item("INACHDESC").ToString), _
                                                    New Object() {rwSO106("MasterId")})
                        End Select
                        Try

                            For Each rwSO106A As DataRow In tbSO106A.Rows
                                If Not DBNull.Value.Equals(rwSO106A("CitemCodeStr")) Then
                                    If ChkInSO106ACitem(rwSO106A("CitemCodeStr"), .Item("BILLHEADFMT")) Then
                                        If Not String.IsNullOrEmpty(fetchText.ToString) Then
                                            fetchText.Append(Environment.NewLine)
                                        End If
                                        fetchText.Append("1")
                                        fetchText.Append(Left(rwSO106.Item("BankID") & Space(3), 3))
                                        fetchText.Append(Space(4))
                                        fetchText.Append(.Item("SENDDATE"))
                                        fetchText.Append("001")
                                        intSeq += 1
                                        fetchText.Append(Right("000000" & intSeq.ToString, 6))
                                        fetchText.Append(Integer.Parse(.Item("APPLYTYPE")) + 1)
                                        If rwSO106.Item("AccountID").ToString.Length = 8 Then
                                            fetchText.Append("G")
                                        Else
                                            fetchText.Append("P")
                                        End If
                                        fetchText.Append(Right("00000000000000" & rwSO106("AccountID").ToString, 14))
                                        If DBNull.Value.Equals(rwSO106("ACHCUSTID")) Then
                                            intSeq = intSeq - 1
                                            intErrorCount += 1
                                            errText.AppendLine(String.Format(Language.NoACHCustId, rwSO106("ID"), rwSO106("AccountID")))
                                        End If
                                        fetchText.Append(Right(Space(20) & rwSO106("ACHCustId"), 20))
                                        fetchText.Append(Left(rwSO106("AccountNameID").ToString & Space(10), 10))
                                        fetchText.Append(Space(2))
                                        fetchText.Append(Space(1))
                                        fetchText.Append(Space(26))
                                        Select Case Integer.Parse(.Item("APPLYTYPE") & "")
                                            Case 0
                                                If Not AlterSO106(FetchOK, "", rwSO106, dsCondition.Tables(0).Rows(0), Nothing) Then
                                                    intSeq = intSeq - 1
                                                    intErrorCount += 1
                                                    errText.AppendLine(String.Format(Language.UpdSO106, rwSO106("ID"), rwSO106("AccountID")))
                                                    Continue For
                                                End If
                                        End Select
                                        If Not DBNull.Value.Equals(rwSO106A("AuthorizeStatus")) AndAlso Integer.Parse(rwSO106A("AuthorizeStatus")) = 4 Then
                                            'DAO.ExecNqry(_DAL.UpdateSO106AStatus(FetchOK),
                                            '             New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                                            '                           Me.LoginInfo.EntryName, rwSO106A("RowId")})
                                            Dim oo As Integer = DAO.ExecNqry(_DAL.UpdateSO106AStatus(FetchOK),
                                                        New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNow),
                                                                      Me.LoginInfo.EntryName, rwSO106A("ctid")})
                                            Debug.Print(oo)
                                        End If
                                    Else
                                        intErrorCount += 1
                                        errText.AppendLine(String.Format(Language.NoFoundCitem, rwSO106("ID"), rwSO106("AccountID")))
                                    End If
                                End If
                            Next
                        Catch ex As Exception
                            result.ErrorCode = -1
                            result.ErrorMessage = ex.ToString

                        Finally
                            If tbSO106A IsNot Nothing Then
                                tbSO106A.Dispose()
                                tbSO106A = Nothing
                            End If
                        End Try

                    Next
                    If intSeq > 0 OrElse intErrorCount > 0 Then
                        If intSeq > 0 Then
                            Dim strBankID As String = DAO.ExecSclr(_DAL.QueryAsignBankID, New Object() {.Item("BANKCODE")})
                            fetchText.Append(Environment.NewLine)
                            fetchText.Append("2")
                            fetchText.Append(Left(strBankID & Space(3), 3))
                            fetchText.Append(Space(4))
                            fetchText.Append(Right(Space(8) & .Item("SENDDATE"), 8))
                            fetchText.Append("001")
                            fetchText.Append("B")
                            fetchText.Append(Right("000000" & intSeq.ToString, 6))
                            fetchText.Append(Space(8))
                            fetchText.Append("000000")
                            fetchText.Append("000000")
                            fetchText.Append(Space(54))
                        End If
                        returnFile = WriteZip(fetchText.ToString, errText.ToString, Integer.Parse(.Item("APPLYTYPE")))
                        result.ErrorCode = 0
                        result.ErrorMessage = Nothing
                        result.ResultBoolean = True
                        If Not String.IsNullOrEmpty(returnFile) Then
                            result.DownloadFileName = returnFile
                        End If

                    End If

                End With
                tbSO106.Dispose()
            End Using
            RunTime.Stop()
            If intSeq > 0 OrElse intErrorCount > 0 Then
                result.ResultXML = String.Format(Language.ProcResult, _
                                                                          intSeq + intErrorCount, intSeq, intErrorCount, _
                                                                          Math.Round(RunTime.Elapsed.TotalSeconds, 1))
            End If
            If blnAutoClose Then
                trans.Commit()
            End If
        Catch ex As Exception
            result.ErrorCode = -1
            result.ErrorMessage = ex.ToString
            trans.Rollback()
        Finally
            If tbLogSO106 IsNot Nothing Then
                tbLogSO106.Dispose()
                tbLogSO106 = Nothing
            End If
            If RunTime IsNot Nothing Then
                RunTime = Nothing
            End If
            If CSLog IsNot Nothing Then
                CSLog.Dispose()
                CSLog = Nothing
            End If
        End Try


        Return result

    End Function
    Public Function QueryAllData() As DataSet
        Dim dsReturn As New DataSet
        Dim tbCompCode As DataTable = Nothing
        Dim tbCD068A As DataTable = Nothing
        Dim tbBankId As DataTable = Nothing
        Dim tbBillHeadFmt As DataTable = Nothing
        Dim tbCitemCode As DataTable = Nothing
        Dim tbCD068 As DataTable = Nothing
        Try
            tbCompCode = QueryCompCode.Copy
            tbCompCode.TableName = tbCompCodeName
            tbBankId = QueryBankId.Copy
            tbBankId.TableName = tbBankIdName
            tbBillHeadFmt = QueryBillHeadFmt.Copy
            tbBillHeadFmt.TableName = tbBillHeadFmtName
            tbCitemCode = QueryCitemCode.Copy
            tbCitemCode.TableName = tbCitemCodeName
            tbCD068 = QueryCD068.Copy
            tbCD068.TableName = tbCD068Name
            tbCD068A = QueryCD068A.Copy
            tbCD068A.TableName = tbCD068AName
            With dsReturn.Tables
                .Add(tbBankId)
                .Add(tbBillHeadFmt)
                .Add(tbCD068)
                .Add(tbCitemCode)
                .Add(tbCompCode)
                .Add(tbCD068A)
            End With


        Catch ex As Exception
            Throw
        Finally
            If tbCompCode IsNot Nothing Then
                tbCompCode.Dispose()
                tbCompCode = Nothing
            End If
        End Try
        Return dsReturn
    End Function
    Public Function QueryCompCode() As DataTable
        If Me.LoginInfo.GroupId = "0" AndAlso 1 = 0 Then
            Return DAO.ExecQry(_DAL.QueryCompCode("0"))
        Else
            Return DAO.ExecQry(_DAL.QueryCompCode(Me.LoginInfo.GroupId),
                               New Object() {Me.LoginInfo.EntryId})
        End If
    End Function
    Public Function QueryBankId() As DataTable
        Return DAO.ExecQry(_DAL.QueryBankId, New Object() {Me.LoginInfo.CompCode})
    End Function
    Public Function QueryBillHeadFmt() As DataTable
        Return DAO.ExecQry(_DAL.QueryBillHeadFmt)
    End Function
    Public Function QueryCitemCode() As DataTable
        Return DAO.ExecQry(_DAL.QueryCitemCode)
    End Function
    Public Function QueryCD068() As DataTable
        Return DAO.ExecQry(_DAL.QueryCD068)
    End Function
    Public Function QueryCD068A() As DataTable
        Return DAO.ExecQry(_DAL.QueryCD068A)
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If _DAL IsNot Nothing Then
                    _DAL.Dispose()
                    _DAL = Nothing
                End If
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                    DAO = Nothing
                End If
                If Language IsNot Nothing Then
                    Language.Dispose()
                    Language = Nothing
                End If
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
