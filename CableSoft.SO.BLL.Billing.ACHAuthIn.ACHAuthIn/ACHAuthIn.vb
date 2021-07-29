Imports CableSoft.BLL.Utility
Imports System.Web
Imports System.Xml
Imports System.Data.Common
Public Class ACHAuthIn
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New ACHAuthInDALMultiDB(Me.LoginInfo.Provider)
    Private Const tbCompCodeName As String = "CompCode"
    Private Const tbFormatTypeName As String = "FormatType"
    Private Const tbBankIdName As String = "BankId"
    Private Const tbBillHeadFmtName As String = "BillHeadFmt"
    Private Const tbCitemCodeName As String = "CitemCode"
    Private Const tbCD068Name As String = "CD068"
    Private Const tbStopAllName As String = "StopAll"
    Private Const tbInputACTHNOName As String = "InputACTHNO"
    Private ReplyType As AuthType = AuthType.ErrorType
    Private Const TxtDirName As String = "TXT"
    Private StopAll As Boolean = False
    Private _dsInputData As DataSet = Nothing
    Private OracleDate As DateTime
    Private ContextUpdDate As String = Nothing
    Private SucessCount As Integer = 0
    Private CD008Where As String = String.Empty
    Private CanUpdateOldAuth As Boolean = False
    Private UpdateOldAuthRowIds As String = Nothing
    Private OldAuthNewRowId As String = Nothing
    Private NewMasterIdSeq As Object
    Private SourceType As FormatType
    Private Language As New CableSoft.BLL.Language.SO61.ACHAuthInLanguage
    Private FNowDate As Date = Date.Now
    Friend Enum AuthType
        Auth = 0
        CancelAuth = 1
        OldAuth = 2
        ErrorType = 3
    End Enum
    Friend Enum FormatType
        NewType = 2
        OldType = 1
    End Enum
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
    Public Function Launch(ByVal dsInputData As DataSet,
                            ByVal fileName As String,
                            ByVal StopDate1 As String,
                            ByVal StopDate2 As String) As RIAResult
        Dim Result As New RIAResult
        Dim tbSO106 As DataTable = Nothing
        Dim tbSO106A As DataTable = Nothing
        Dim LogContext As New System.Text.StringBuilder()
        Dim InputAchtNo As String = Nothing
        Dim ErrorMsg As String = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim MasterId As String = Nothing
        Dim TimeSpend As New Stopwatch()
        Dim ContextString As String = Nothing
        TimeSpend.Start()
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
        _dsInputData = dsInputData

        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId)
        OracleDate = Date.Parse(DAO.ExecSclr(_DAL.QueryOracleDate))
        SucessCount = 0
        SourceType = FormatType.OldType
        If Integer.Parse(_dsInputData.Tables(tbFormatTypeName).Rows(0).Item("CodeNo")) = 2 Then
            SourceType = FormatType.NewType
        End If

        Try
            ContextString = System.IO.File.ReadAllText(CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & "TXT\" & fileName)
            For Each ACHTNo As DataRow In _dsInputData.Tables(tbInputACTHNOName).Rows
                If String.IsNullOrEmpty(InputAchtNo) Then
                    InputAchtNo = String.Format("'{0}'", ACHTNo("ACHTNO"))
                Else
                    InputAchtNo = String.Format("{0},'{1}'", InputAchtNo, ACHTNo("ACHTNO"))
                End If
            Next
            For Each conText As String In ContextString.Split(Environment.NewLine)
                conText = conText.Replace(Chr(10), "").Replace(Chr(13), "")
                If Not String.IsNullOrEmpty(conText) Then
                    If (conText.Substring(0, 3).ToUpper = "BOF".ToUpper) OrElse (conText.Substring(0, 3).ToUpper = "EOF".ToUpper) Then
                        ContextUpdDate = (Integer.Parse(conText.Substring(9, 8)) + 19110000).ToString
                    Else
                        If (conText.Length < 108) OrElse (conText.Substring(106, 1).ToUpper <> "R".ToUpper) Then
                            Return New RIAResult With {.ErrorCode = -1, .ErrorMessage = Language.FormatError, .ResultBoolean = False}
                            Exit For
                        End If
                        ReplyType = chkReplyType(conText)
                        If ReplyType = AuthType.ErrorType Then
                            Return New RIAResult With {.ErrorCode = -2, .ErrorMessage = Language.ReplyFormatError, .ResultBoolean = False}
                            Exit For
                        End If
                      
                        tbSO106 = DAO.ExecQry(_DAL.QuerySO106Data(ReplyType), New Object() {
                                              GetACHCustId(conText),
                                              GetAccountId(conText),
                                              GetContextStringACH(conText)})
                        If tbSO106.Rows.Count = 0 Then
                            ErrorMsg = Language.ACHCustIdNotInDB
                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, ""))
                            Continue For
                        End If

                        If String.IsNullOrEmpty(StopDate1) Then
                            StopDate1 = Nothing
                            StopDate1 = Date.Today.ToString("yyyyMMdd")

                        End If
                        If String.IsNullOrEmpty(StopDate2) Then
                            StopDate2 = Nothing
                            StopDate2 = Date.Today.ToString("yyyyMMdd")
                            'StopDate2 = OracleDate.ToShortDateString.Replace("/", "")
                        End If
                        If Not IsDate(StopDate1) Then
                            StopDate1 = Date.ParseExact(StopDate1, "yyyyMMdd", Globalization.CultureInfo.InvariantCulture)
                        End If
                        If Not IsDate(StopDate2) Then
                            StopDate2 = Date.ParseExact(StopDate2, "yyyyMMdd", Globalization.CultureInfo.InvariantCulture)
                        End If

                        'StopDate1 = StopDate1.Replace("/", "")
                        'StopDate2 = StopDate2.Replace("/", "")
                        'If StopDate1.Length < "20141014235900".Length Then
                        '    StopDate1 = String.Format("{0}000000", StopDate1.Replace("/", ""))
                        'End If
                        'If StopDate2.Length < "20141014235900".Length Then
                        '    StopDate2 = String.Format("{0}235959", StopDate2.Replace("/", ""))
                        'End If
                        
                        If SourceType = FormatType.NewType Then
                            Select Case ReplyType
                                Case AuthType.CancelAuth
                                    tbSO106A = DAO.ExecQry(_DAL.QuerySO106A(ReplyType, InputAchtNo, StopDate1.Replace("/", ""), StopDate2.Replace("/", "")),
                                     New Object() {tbSO106.Rows(0).Item("MasterId"),
                                                   GetContextStringACH(conText), Date.Parse(StopDate1), _
                                                                     Date.Parse(StopDate2)})
                                Case Else
                                    tbSO106A = DAO.ExecQry(_DAL.QuerySO106A(ReplyType, InputAchtNo, StopDate1.Replace("/", ""), StopDate2.Replace("/", "")),
                                    New Object() {tbSO106.Rows(0).Item("MasterId"),
                                                  GetContextStringACH(conText)})
                            End Select
                          


                            If tbSO106A.Rows.Count = 0 Then
                                ErrorMsg = Language.NotFoundSO106A
                                LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                Continue For
                            End If
                        Else
                            tbSO106A = DAO.ExecQry(_DAL.GetEmptySO106A)
                            Dim rw106a = tbSO106A.NewRow
                            rw106a("masterid") = "-1"
                            tbSO106A.Rows.Add(rw106a)
                        End If


                        '回覆失敗
                        If Not IsAuthInOk(conText, ErrorMsg) Then
                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                            UpdateSO106A(AuthType.ErrorType, tbSO106A.Copy, ErrorMsg)
                            InsertOrUpdSO106(AuthType.ErrorType, conText, ErrorMsg)
                            If (ReplyType = AuthType.Auth) Then
                                ClearSO106(tbSO106.Rows(0), tbSO106A.Rows(0))
                            End If
                            Continue For
                        End If
                        '回覆成功
                        Dim BillHeadString As String = Nothing
                        For Each rwBillHead As DataRow In _dsInputData.Tables(tbBillHeadFmtName).Rows
                            If String.IsNullOrEmpty(BillHeadString) Then
                                BillHeadString = String.Format("'{0}'", rwBillHead.Item("BillHeadFmt"))
                            Else
                                BillHeadString = String.Format("{0},'{1}'", BillHeadString, rwBillHead.Item("BillHeadFmt"))
                            End If
                        Next
                        If ReplyType = AuthType.Auth Then
                            CD008Where = String.Format(_DAL.GetCD008Where, BillHeadString)

                        End If
                        CD008Where = " And 1= 1 "
                        Select Case ReplyType
                            '授權
                            Case AuthType.Auth
                                Select Case SourceType
                                    Case FormatType.NewType
                                        '新格式
                                        If MasterId <> tbSO106A.Rows(0).Item("MasterId") Then
                                            If Not InsertOrUpdSO106(AuthType.Auth, conText, Nothing) Then
                                                ErrorMsg = Language.UpdSO106AError
                                                LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                                Continue For
                                            End If

                                            'Cancel to update so002a by kin 2017/03/21
                                            'If Not InsertOrUpdSO002A(tbSO106.Rows(0).Item("MasterId"),
                                            '                         tbSO106.Rows(0).Item("CustId"), GetAccountId(conText)) Then
                                            '    ErrorMsg = Language.InsSO002AError
                                            '    LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                            '    Continue For
                                            'End If
                                        End If
                                    Case Else
                                        '舊格式
                                        If Not InsertOrUpdSO106(AuthType.Auth, conText, Nothing) Then
                                            ErrorMsg = Language.UpdSO106AError
                                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                            Continue For
                                        End If
                                        'Cancel to update so002a by kin 2017/03/21
                                        'If Not InsertOrUpdSO002A(tbSO106.Rows(0).Item("MasterId"),
                                        '                           tbSO106.Rows(0).Item("CustId"), GetAccountId(conText)) Then
                                        '    ErrorMsg = Language.InsSO002AError
                                        '    LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        '    Continue For
                                        'End If
                                End Select
                                If SourceType = FormatType.NewType Then
                                 
                                    '更新SO003C                                    
                                    If Not UpdACHSO003C(tbSO106.Rows(0).Item("MasterId")) Then
                                        ErrorMsg = Language.UpdSO003CError
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                    If Not UpdateSO003(tbSO106.Rows(0).Item("Masterid")) Then
                                        ErrorMsg = Language.UpdSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                    '更新非週期收費項目
                                    If Not UpdNonePeriod(tbSO106.Rows(0).Item("MasterId")) Then
                                        ErrorMsg = Language.UpdNoneSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                Else
                                   
                                    '更新SO003C
                                    If Not UpdACHSO003C(tbSO106.Rows(0).Item("MasterId")) Then
                                        ErrorMsg = Language.UpdSO003CError
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                    If Not UpdateSO003(tbSO106.Rows(0).Item("MasterId")) Then
                                        ErrorMsg = Language.UpdSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                    '更新非週期收費項目
                                    If Not UpdNonePeriod(tbSO106.Rows(0).Item("MasterId")) Then
                                        ErrorMsg = Language.UpdNoneSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                End If

                                'If Not InsertSO004(tbSO106.Rows(0)) Then
                                '    ErrorMsg = Language.InsSO004Error
                                '    LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId2")))
                                '    Continue For
                                'End If
                                If SourceType = FormatType.NewType Then
                                    If Not UpdateSO106A(AuthType.Auth, tbSO106A.Copy, Nothing) Then
                                        ErrorMsg = Language.UpdSO106AError2
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    Else
                                        SucessCount += 1
                                    End If
                                    MasterId = tbSO106A.Rows(0).Item("MasterId").ToString
                                Else
                                    SucessCount += 1
                                End If

                            Case AuthType.CancelAuth
                                '取消授權
                                If SourceType = FormatType.NewType Then
                                    If Not UpdateSO106A(AuthType.CancelAuth, tbSO106A.Copy, Nothing) Then
                                        ErrorMsg = Language.UpdSO106AError2
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                End If

                                If Not InsertOrUpdSO106(AuthType.CancelAuth, conText, Nothing) Then
                                    ErrorMsg = Language.UpdSO106Error
                                    LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                    Continue For
                                End If
                                If SourceType = FormatType.NewType Then
                                    Dim aPTCode As Integer = tbSO106.Rows(0).Item("PTCode")
                                    Dim aPTName As String = tbSO106.Rows(0).Item("PTName")
                                    Dim aCMCode As Integer = tbSO106.Rows(0).Item("CMCode")
                                    Dim aCMName As String = tbSO106.Rows(0).Item("CMName")
                                    Using tb As DataTable = DAO.ExecQry(_DAL.GetPTCode)
                                        If tb IsNot Nothing Then
                                            If tb.Rows.Count > 0 Then
                                                aPTCode = tb.Rows(0).Item("CODENO")
                                                aPTName = tb.Rows(0).Item("Description")
                                            End If
                                        End If
                                    End Using
                                    'cmd.CommandText = _DAL.GetPTCode
                                    'Using dr As DbDataReader = cmd.ExecuteReader()
                                    '    dr.Read()
                                    '    aPTCode = dr.Item("CODENO")
                                    '    aPTName = dr.Item("Description")
                                    'End Using
                                    Using tb As DataTable = DAO.ExecQry(_DAL.GetDefCMCode(LoginInfo, String.Empty))
                                        If tb IsNot Nothing Then
                                            If tb.Rows.Count > 0 Then
                                                aCMCode = tb.Rows(0).Item("CODENO")
                                                aCMName = tb.Rows(0).Item("Description")
                                            End If
                                        End If
                                    End Using


                                    'cmd.CommandText = _DAL.GetDefCMCode(LoginInfo, String.Empty)

                                    'Using dr As DbDataReader = cmd.ExecuteReader
                                    '    dr.Read()
                                    '    aCMCode = dr.Item("CODENO")
                                    '    aCMName = dr.Item("Description")
                                    'End Using                                
                                    If Not StopSO003(aCMCode, aCMName, aPTCode, aPTName, tbSO106.Rows(0)) Then
                                        ErrorMsg = Language.StopSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For

                                    End If
                                    If Not StopNonePeriod(aCMCode, aCMName, aPTCode, aPTName, tbSO106.Rows(0)) Then
                                        ErrorMsg = Language.StopNonePeriodError
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                    If Not StopACHSO003C(tbSO106.Rows(0).Item("MasterId"), aCMCode, aCMName, aPTCode, aPTName) Then
                                        ErrorMsg = Language.StopSO003CError
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                    SucessCount += 1
                                Else
                                    SucessCount += 1
                                End If

                            Case AuthType.OldAuth
                                '新增舊有已簽約委繳戶資料
                                If SourceType = FormatType.NewType Then
                                    If MasterId <> tbSO106A.Rows(0).Item("MasterId") Then
                                        If (CanUpdateOldAuth) AndAlso (Not String.IsNullOrEmpty(UpdateOldAuthRowIds)) Then
                                            UpdateOldAuth()
                                        End If
                                        MasterId = tbSO106A.Rows(0).Item("MasterId")

                                        If InsertOrUpdSO106(AuthType.OldAuth, conText, Nothing) Then
                                            SucessCount += 1
                                            CanUpdateOldAuth = True
                                        Else
                                            ErrorMsg = Language.UpdSO106Error
                                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                            Continue For
                                        End If
                                    Else
                                        MasterId = tbSO106A.Rows(0).Item("MasterId")
                                    End If
                                Else
                                    If InsertOrUpdSO106(AuthType.OldAuth, conText, Nothing) Then
                                        UpdateOldAuth()
                                        SucessCount += 1
                                    Else
                                        ErrorMsg = Language.UpdSO106Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                End If

                                If SourceType = FormatType.NewType Then
                                    If Not UpdateSO106A(AuthType.OldAuth, tbSO106A.Copy, Nothing) Then
                                        ErrorMsg = Language.NotFoundSO106A
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("ID")))
                                        Continue For
                                    End If
                                End If

                        End Select
                    End If
                End If
            Next
            If Not String.IsNullOrEmpty(UpdateOldAuthRowIds) Then
                UpdateOldAuth()
            End If
            Dim LogFileName As String = String.Empty

            If Not String.IsNullOrEmpty(LogContext.ToString) Then
                'LogFileName = WriteLogFile(LogContext.ToString)
                LogFileName = Now.ToString("yyyyMMddHHmmssff") & "-ACHLog.zip"
                Dim zipFileName As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & TxtDirName & "\" & LogFileName
                Using zip As New Ionic.Zip.ZipFile(zipFileName,
                                                   System.Text.Encoding.GetEncoding(950))
                    zip.AddEntry(Me.LoginInfo.EntryId & "-" & "ACHLog.Txt", LogContext.ToString)
                    zip.Save()
                    zip.Dispose()
                    Result.DownloadFileName = TxtDirName & "\" & LogFileName
                End Using
            End If
            trans.Commit()
            'trans.Rollback()
            TimeSpend.Stop()
            Result.ResultBoolean = True
            Result.ResultXML = String.Format(Language.RunTotalRecord & _
                                                            Environment.NewLine & _
                                                            Language.RunErrorRecord & Environment.NewLine & _
                                                            Language.RunSpendTime, SucessCount,
                                                            LogContext.ToString.Split(Environment.NewLine).Count - 1,
                                                            Math.Round(TimeSpend.Elapsed.TotalSeconds, 1))

            'If Not String.IsNullOrEmpty(LogFileName) Then
            '    Result.ResultXML = String.Format("{0}", Result.ResultXML)
            'End If
        Catch ex As Exception
            trans.Rollback()
            Return New RIAResult With {.ErrorCode = -99, .ErrorMessage = ex.ToString, .ResultBoolean = False}
        Finally
            TimeSpend.Reset()
            If tbSO106 IsNot Nothing Then
                tbSO106.Dispose()
                tbSO106 = Nothing
            End If
            If tbSO106A IsNot Nothing Then
                tbSO106A.Dispose()
                tbSO106A = Nothing
            End If
            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                End If
            End If
        End Try
        Return Result
    End Function
    Public Function Execute(ByVal dsInputData As DataSet,
                            ByVal ContextString As String,
                            ByVal StopDate1 As String,
                            ByVal StopDate2 As String) As RIAResult

        Dim Result As New RIAResult
        Dim tbSO106 As DataTable = Nothing
        Dim tbSO106A As DataTable = Nothing
        Dim LogContext As New System.Text.StringBuilder()
        Dim InputAchtNo As String = Nothing
        Dim ErrorMsg As String = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim MasterId As String = Nothing
        Dim TimeSpend As New Stopwatch()
        TimeSpend.Start()
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
        _dsInputData = dsInputData

        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans

        OracleDate = Date.Parse(DAO.ExecSclr(_DAL.QueryOracleDate))
        SucessCount = 0
        SourceType = FormatType.OldType
        If Integer.Parse(_dsInputData.Tables(tbFormatTypeName).Rows(0).Item("CodeNo")) = 2 Then
            SourceType = FormatType.NewType
        End If

        Try
            For Each ACHTNo As DataRow In _dsInputData.Tables(tbInputACTHNOName).Rows
                If String.IsNullOrEmpty(InputAchtNo) Then
                    InputAchtNo = String.Format("'{0}'", ACHTNo("ACHTNO"))
                Else
                    InputAchtNo = String.Format("{0},'{1}'", InputAchtNo, ACHTNo("ACHTNO"))
                End If
            Next
            For Each conText As String In ContextString.Split(Environment.NewLine)
                conText = conText.Replace(Chr(10), "").Replace(Chr(13), "")
                If Not String.IsNullOrEmpty(conText) Then
                    If (conText.Substring(0, 3).ToUpper = "BOF".ToUpper) OrElse (conText.Substring(0, 3).ToUpper = "EOF".ToUpper) Then
                        ContextUpdDate = (Integer.Parse(conText.Substring(9, 8)) + 19110000).ToString
                    Else
                        If (conText.Length < 108) OrElse (conText.Substring(106, 1).ToUpper <> "R".ToUpper) Then
                            Return New RIAResult With {.ErrorCode = -1, .ErrorMessage = Language.FormatError, .ResultBoolean = False}
                            Exit For
                        End If
                        ReplyType = chkReplyType(conText)
                        Dim aAction As String = Nothing
                        Select Case ReplyType
                            Case AuthType.Auth
                                aAction = Language.AuthClientInfo
                            Case AuthType.CancelAuth
                                aAction = Language.CancelClientInfo
                            Case AuthType.OldAuth
                                aAction = Language.OldClientInfo
                            Case Else
                                aAction = Language.AuthClientInfo
                        End Select
                        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, aAction)
                        If ReplyType = AuthType.ErrorType Then
                            Return New RIAResult With {.ErrorCode = -2, .ErrorMessage = Language.ReplyFormatError, .ResultBoolean = False}
                            Exit For
                        End If
                        tbSO106 = DAO.ExecQry(_DAL.QuerySO106Data(ReplyType), New Object() {
                                              GetACHCustId(conText),
                                              GetAccountId(conText),
                                              GetContextStringACH(conText)})
                        If tbSO106.Rows.Count = 0 Then
                            ErrorMsg = Language.ACHCustIdNotInDB
                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, ""))
                            Continue For
                        End If
                        If String.IsNullOrEmpty(StopDate1) Then
                            StopDate1 = OracleDate.ToShortDateString.Replace("/", "")
                        End If
                        If String.IsNullOrEmpty(StopDate2) Then
                            StopDate2 = OracleDate.ToShortDateString.Replace("/", "")
                        End If
                        StopDate1 = StopDate1.Replace("/", "")
                        StopDate2 = StopDate2.Replace("/", "")
                        If StopDate1.Length < "20141014235900".Length Then
                            StopDate1 = String.Format("{0}000000", StopDate1.Replace("/", ""))
                        End If
                        If StopDate2.Length < "20141014235900".Length Then
                            StopDate2 = String.Format("{0}235959", StopDate2.Replace("/", ""))
                        End If

                        If SourceType = FormatType.NewType Then
                            tbSO106A = DAO.ExecQry(_DAL.QuerySO106A(ReplyType, InputAchtNo, StopDate1.Replace("/", ""), StopDate2.Replace("/", "")),
                                       New Object() {tbSO106.Rows(0).Item("MasterId"),
                                                     GetContextStringACH(conText)})


                            If tbSO106A.Rows.Count = 0 Then
                                ErrorMsg = Language.NotFoundSO106A
                                LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                Continue For
                            End If
                        End If


                        '回覆失敗
                        If Not IsAuthInOk(conText, ErrorMsg) Then
                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                            UpdateSO106A(AuthType.ErrorType, tbSO106A, ErrorMsg)
                            InsertOrUpdSO106(AuthType.ErrorType, conText, ErrorMsg)
                            If (ReplyType = AuthType.Auth) Then
                                ClearSO106(tbSO106.Rows(0), tbSO106A.Rows(0))
                            End If
                            Continue For
                        End If
                        '回覆成功
                        Dim BillHeadString As String = Nothing
                        For Each rwBillHead As DataRow In _dsInputData.Tables(tbBillHeadFmtName).Rows
                            If String.IsNullOrEmpty(BillHeadString) Then
                                BillHeadString = String.Format("'{0}'", rwBillHead.Item("BillHeadFmt"))
                            Else
                                BillHeadString = String.Format("{0},'{1}'", BillHeadString, rwBillHead.Item("BillHeadFmt"))
                            End If
                        Next
                        If ReplyType = AuthType.Auth Then
                            CD008Where = String.Format("And Exists(Select CitemCode From SO003 B Where " &
                                                                                 "B.Custid = SO106.Custid " &
                                                                                    " And B.CompCode=SO106.CompCode " &
                                                                                     " And instr(','||SO106.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " &
                                                                                     " And Exists(Select * From  CD068 C Where " &
                                                                                                         " instr(','||C.Citemcodestr||',',','||B.CitemCode||',')>0 " &
                                                                                                         " And C.BillHeadFmt In({0}) And C.ACHType=1 ))", BillHeadString)

                        End If
                        Select Case ReplyType
                            '授權
                            Case AuthType.Auth
                                Select Case SourceType
                                    Case FormatType.NewType
                                        '新格式
                                        If MasterId <> tbSO106A.Rows(0).Item("MasterId") Then
                                            If Not InsertOrUpdSO106(AuthType.Auth, conText, Nothing) Then
                                                ErrorMsg = Language.UpdSO106AError
                                                LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                                Continue For
                                            End If

                                            If Not InsertOrUpdSO002A(tbSO106.Rows(0).Item("MasterId"),
                                                                     tbSO106.Rows(0).Item("CustId"), GetAccountId(conText)) Then
                                                ErrorMsg = Language.InsSO002AError
                                                LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                                Continue For
                                            End If
                                        End If
                                    Case Else
                                        '舊格式
                                        If Not InsertOrUpdSO106(AuthType.Auth, conText, Nothing) Then
                                            ErrorMsg = Language.UpdSO106AError
                                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                            Continue For
                                        End If

                                        If Not InsertOrUpdSO002A(tbSO106.Rows(0).Item("MasterId"),
                                                                   tbSO106.Rows(0).Item("CustId"), GetAccountId(conText)) Then
                                            ErrorMsg = Language.InsSO002AError
                                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                            Continue For
                                        End If
                                End Select
                                If SourceType = FormatType.NewType Then
                                    If Not UpdateSO003(conText, tbSO106.Rows(0).Item("CustId"),
                                                  tbSO106A.Rows(0)) Then
                                        ErrorMsg = Language.UpdSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        Continue For
                                    End If
                                Else
                                    If Not UpdateSO003(conText, tbSO106.Rows(0).Item("CustId"), InputAchtNo, tbSO106.Rows(0).Item("CitemStr")) Then
                                        ErrorMsg = Language.UpdSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        Continue For
                                    End If
                                End If

                                If Not InsertSO004(tbSO106.Rows(0)) Then
                                    ErrorMsg = Language.InsSO004Error
                                    LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                    Continue For
                                End If
                                If SourceType = FormatType.NewType Then
                                    If Not UpdateSO106A(AuthType.Auth, tbSO106A, Nothing) Then
                                        ErrorMsg = Language.UpdSO106AError2
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        Continue For
                                    Else
                                        SucessCount += 1
                                    End If
                                    MasterId = tbSO106A.Rows(0).Item("MasterId").ToString
                                Else
                                    SucessCount += 1
                                End If

                            Case AuthType.CancelAuth
                                '取消授權
                                If SourceType = FormatType.NewType Then
                                    If Not UpdateSO106A(AuthType.CancelAuth, tbSO106A, Nothing) Then
                                        ErrorMsg = Language.UpdSO106AError2
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        Continue For
                                    End If
                                End If

                                If Not InsertOrUpdSO106(AuthType.CancelAuth, conText, Nothing) Then
                                    ErrorMsg = Language.UpdSO106Error
                                    LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                    Continue For
                                End If
                                If SourceType = FormatType.NewType Then
                                    If Not StopSO003(tbSO106.Rows(0), tbSO106A.Rows(0)) Then
                                        ErrorMsg = Language.StopSO003Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        Continue For
                                    Else
                                        SucessCount += 1
                                    End If
                                Else
                                    SucessCount += 1
                                End If

                            Case AuthType.OldAuth
                                '新增舊有已簽約委繳戶資料
                                If SourceType = FormatType.NewType Then
                                    If MasterId <> tbSO106A.Rows(0).Item("MasterId") Then
                                        If (CanUpdateOldAuth) AndAlso (Not String.IsNullOrEmpty(UpdateOldAuthRowIds)) Then
                                            UpdateOldAuth()
                                        End If
                                        MasterId = tbSO106A.Rows(0).Item("MasterId")
                                        If InsertOrUpdSO106(AuthType.OldAuth, conText, Nothing) Then
                                            SucessCount += 1
                                            CanUpdateOldAuth = True
                                        Else
                                            ErrorMsg = Language.UpdSO106Error
                                            LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                            Continue For
                                        End If
                                    Else
                                        MasterId = tbSO106A.Rows(0).Item("MasterId")
                                    End If
                                Else
                                    If InsertOrUpdSO106(AuthType.OldAuth, conText, Nothing) Then
                                        UpdateOldAuth()
                                        SucessCount += 1
                                    Else
                                        ErrorMsg = Language.UpdSO106Error
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        Continue For
                                    End If
                                End If

                                If SourceType = FormatType.NewType Then
                                    If Not UpdateSO106A(AuthType.OldAuth, tbSO106A, Nothing) Then
                                        ErrorMsg = Language.NotFoundSO106A
                                        LogContext.AppendLine(GetErrMsg(conText, ErrorMsg, tbSO106.Rows(0).Item("CustId")))
                                        Continue For
                                    End If
                                End If

                        End Select
                    End If
                End If
            Next
            If Not String.IsNullOrEmpty(UpdateOldAuthRowIds) Then
                UpdateOldAuth()
            End If
            Dim LogFileName As String = String.Empty
            If Not String.IsNullOrEmpty(LogContext.ToString) Then
                LogFileName = WriteLogFile(LogContext.ToString)
            End If
            trans.Commit()
            TimeSpend.Stop()
            Result.ResultBoolean = True
            Result.ResultXML = String.Format(Language.RunTotalRecord & _
                                                            Environment.NewLine & _
                                                            Language.RunErrorRecord & Environment.NewLine & _
                                                            Language.RunSpendTime, SucessCount,
                                                            LogContext.ToString.Split(Environment.NewLine).Count - 1,
                                                            Math.Round(TimeSpend.Elapsed.TotalSeconds, 1))
            If Not String.IsNullOrEmpty(LogFileName) Then
                Result.ResultXML = String.Format("{0};{1}", Result.ResultXML, LogFileName)
            End If
        Catch ex As Exception
            trans.Rollback()
            Return New RIAResult With {.ErrorCode = -99, .ErrorMessage = ex.ToString, .ResultBoolean = False}
        Finally
            TimeSpend.Reset()
            If tbSO106 IsNot Nothing Then
                tbSO106.Dispose()
                tbSO106 = Nothing
            End If
            If tbSO106A IsNot Nothing Then
                tbSO106A.Dispose()
                tbSO106A = Nothing
            End If
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                End If
            End If
        End Try
        Return Result

    End Function
    Private Function InsertSO004(ByVal rwSO106 As DataRow) As Boolean
        If SourceType = FormatType.NewType Then
            Return True
        End If
        Try
            DAO.ExecNqry(_DAL.InsertSO004, New Object() {
                         rwSO106.Item("AccountID"),
                         rwSO106.Item("BankCode"),
                         rwSO106.Item("BankName"),
                         rwSO106.Item("PTCode"),
                         rwSO106.Item("PTname"),
                         rwSO106.Item("CMCode"),
                         rwSO106.Item("CMName"),
                         rwSO106.Item("MasterId")
                         })

        Catch ex As Exception
            Throw
        End Try
        Return True
    End Function
    Private Function UpdateOldAuth() As Boolean
        If String.IsNullOrEmpty(UpdateOldAuthRowIds) Then
            Return True
        End If
        Try
            DAO.ExecNqry(_DAL.UpdateOldAuth(UpdateOldAuthRowIds))
        Catch ex As Exception
        Finally
            UpdateOldAuthRowIds = Nothing
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
    Private Overloads Function StopSO003(ByVal rwSO106 As DataRow, ByVal rwSO106A As DataRow) As Boolean
        Try
            DAO.ExecNqry(_DAL.StopSO003(rwSO106A.Item("CitemCodeStr")), New Object() {
                         Me.LoginInfo.CompCode,
                         rwSO106.Item("CustId"),
                         rwSO106.Item("ACCOUNTID")})
        Catch ex As Exception
            Throw
        Finally

        End Try
        Return True
    End Function
    Private Function ClearSO106(ByVal rwSO106 As DataRow, ByVal rwSO106A As DataRow) As Boolean
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
    Private Function InsertOrUpdSO002A(ByVal MasterId As Integer, ByVal CustId As Integer,
                                       ByVal AccountId As String) As Boolean
        Dim tbUpdData As DataTable = Nothing
        Try
            tbUpdData = DAO.ExecQry(_DAL.QueryUpdSO002AData, New Object() {MasterId,
                                                                           AccountId, Me.LoginInfo.CompCode})
            If tbUpdData.Rows.Count > 0 Then
                If Integer.Parse(DAO.ExecSclr(_DAL.IsExistsSO002A, New Object() {
                                         AccountId, CustId, Me.LoginInfo.CompCode})) = 0 Then

                    DAO.ExecNqry(_DAL.InserSO002A, New Object() {
                                 CustId,
                                 Me.LoginInfo.CompCode,
                                 tbUpdData.Rows(0).Item("BankCode"),
                                 tbUpdData.Rows(0).Item("BankName"),
                                 tbUpdData.Rows(0).Item("AccountID"),
                                 tbUpdData.Rows(0).Item("ChargeAddrNo"),
                                 tbUpdData.Rows(0).Item("ChargeAddress"),
                                 tbUpdData.Rows(0).Item("MailAddrNo"),
                                 tbUpdData.Rows(0).Item("MailAddress"),
                                 tbUpdData.Rows(0).Item("AccountName"),
                                 tbUpdData.Rows(0).Item("InvNo"),
                                 tbUpdData.Rows(0).Item("InvTitle"),
                                 tbUpdData.Rows(0).Item("InvAddress"),
                                 tbUpdData.Rows(0).Item("InvoiceType")})
                Else
                    DAO.ExecNqry(_DAL.UpdateSO002A, New Object() {
                                  tbUpdData.Rows(0).Item("BankCode"),
                                 tbUpdData.Rows(0).Item("BankName"),
                                 tbUpdData.Rows(0).Item("AccountID"),
                                  tbUpdData.Rows(0).Item("ChargeAddrNo"),
                                 tbUpdData.Rows(0).Item("ChargeAddress"),
                                 tbUpdData.Rows(0).Item("MailAddrNo"),
                                 tbUpdData.Rows(0).Item("MailAddress"),
                                  CustId, AccountId, Me.LoginInfo.CompCode})
                End If
                If Integer.Parse(DAO.ExecSclr(_DAL.IsExistsSO002AD, New Object() {
                                CustId,
                                tbUpdData.Rows(0).Item("AccountID"),
                                Me.LoginInfo.CompCode,
                                tbUpdData.Rows(0).Item("InvSeqNo")})) = 0 Then

                    DAO.ExecNqry(_DAL.InsertSO002AD, New Object() {
                                 tbUpdData.Rows(0).Item("AccountID"),
                                 Me.LoginInfo.CompCode,
                                 CustId,
                                 tbUpdData.Rows(0).Item("InvSeqNo")})

                End If
            End If

        Catch ex As Exception
            Throw
        Finally
            If tbUpdData IsNot Nothing Then
                tbUpdData.Dispose()
                tbUpdData = Nothing
            End If
        End Try
        Return True
    End Function
    Private Function GetReStatus(ByVal ConText As String) As String
        Select Case Mid(ConText, 105, 1)
            Case "P"
                Return Language.GetReStatusP
                'Return "P:已發送授權書及授權扣款檔"
            Case "R"
                Return Language.GetReStatusR
                'Return "R:先收到回覆訊息但未收到授權書"
            Case "Y"
                Return Language.GetReStatusY
                'Return "Y:先收到回覆訊息後收到授權書"
            Case "M"
                Return Language.GetReStatusM
                'Return "M:先收到授權書但未收到回覆訊息"
            Case "S"
                Return Language.GetReStatusS
                'Return "S:先收到授權書後收到回覆訊息"
            Case "C"
                Return Language.GetReStatusC
                'Return "C:已收到舊件轉檔回覆訊息"
            Case "D"
                Return Language.GetReStatusD
                'Return "D:已收到取消授權扣款回覆訊息"
            Case Else
                Return String.Empty
        End Select
    End Function
    Private Function GetTxtDate(ByVal conText As String) As String
        Return (Integer.Parse(conText.Substring(71, 8).Replace(" ", "").ToString()) + 19110000).ToString()
    End Function
    Private Function InsertOrUpdSO106(ByVal ReplyType As AuthType, ByVal ConText As String,
                                ByVal ErrorMsg As String) As Boolean
        Dim UpdateCount As Integer = 1
        Try
            Select Case ReplyType
                Case AuthType.Auth
                    UpdateCount = Integer.Parse(DAO.ExecNqry(String.Format("{0}{1}",
                                                               _DAL.UpdateSO106(AuthType.Auth),
                                                               CD008Where), New Object() {ContextUpdDate,
                                                                                          GetReStatus(ConText),
                                                                                          Me.LoginInfo.EntryName,
                                                                                         CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                                                                          OracleDate.ToString("yyyyMMddHHmmss"),
                                                                                          GetACHCustId(ConText), GetTxtDate(ConText),
                                                                                          GetAccountId(ConText),
                                                                                          GetContextStringACH(ConText)}))
                Case AuthType.CancelAuth
                    If StopAll Then
                        UpdateCount = Integer.Parse(DAO.ExecNqry(_DAL.UpdateSO106(AuthType.CancelAuth), New Object() {
                                                                                        GetReStatus(ConText),
                                                                                          Me.LoginInfo.EntryName,
                                                                                         CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                                                                        OracleDate.ToString("yyyyMMddHHmmss"),
                                                                                          GetACHCustId(ConText),
                                                                                          GetAccountId(ConText)}))


                    End If
                Case AuthType.OldAuth
                    Using tbUpdData As DataTable = DAO.ExecQry(_DAL.QueryUpdOldAch, New Object() {
                                                               GetAccountId(ConText),
                                                               GetACHCustId(ConText)})

                        If tbUpdData.Rows.Count > 0 Then
                            NewMasterIdSeq = DAO.ExecSclr(_DAL.GetMasterIdSeq)
                            DAO.ExecNqry(_DAL.InsertSO106, New Object() {
                                         ContextUpdDate,
                                         tbUpdData.Rows(0).Item("SendDate"),
                                         tbUpdData.Rows(0).Item("SendDate"),
                                         GetReStatus(ConText),
                                         Me.LoginInfo.EntryName,
                                         CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                         NewMasterIdSeq, tbUpdData.Rows(0).Item("AcceptName"),
                                         tbUpdData.Rows(0).Item("Proposer"), tbUpdData.Rows(0).Item("ID"),
                                         tbUpdData.Rows(0).Item("BankCode"), tbUpdData.Rows(0).Item("BankName"),
                                         tbUpdData.Rows(0).Item("CardCode"), tbUpdData.Rows(0).Item("CardName"),
                                         tbUpdData.Rows(0).Item("StopYM"), tbUpdData.Rows(0).Item("AccountID"),
                                         tbUpdData.Rows(0).Item("AccountName"), tbUpdData.Rows(0).Item("AccountNameID"),
                                         tbUpdData.Rows(0).Item("MediaCode"), tbUpdData.Rows(0).Item("MediaName"),
                                         tbUpdData.Rows(0).Item("IntroID"), tbUpdData.Rows(0).Item("IntroName"),
                                         tbUpdData.Rows(0).Item("Note"), tbUpdData.Rows(0).Item("UpdateFlag"),
                                         tbUpdData.Rows(0).Item("CompCode"), tbUpdData.Rows(0).Item("CustId"),
                                         tbUpdData.Rows(0).Item("CMCode"), tbUpdData.Rows(0).Item("CMName"),
                                         tbUpdData.Rows(0).Item("Alien"), tbUpdData.Rows(0).Item("AccountAlien"),
                                         tbUpdData.Rows(0).Item("AcceptEn"), tbUpdData.Rows(0).Item("CVC2"),
                                         tbUpdData.Rows(0).Item("CitemStr"), tbUpdData.Rows(0).Item("CitemStr2"),
                                         tbUpdData.Rows(0).Item("AddCitemAccount"), tbUpdData.Rows(0).Item("PTCode"),
                                         tbUpdData.Rows(0).Item("PTName"), tbUpdData.Rows(0).Item("ACHCustId"),
                                         tbUpdData.Rows(0).Item("ACHSN"), tbUpdData.Rows(0).Item("ACHTNo"),
                                         tbUpdData.Rows(0).Item("ACHTDESC")
                                     })

                            Using tbSO106 As DataTable = DAO.ExecQry(_DAL.QuerySO106AllData, New Object() {NewMasterIdSeq})
                                OldAuthNewRowId = tbSO106.Rows(0).Item("RowId")
                                If String.IsNullOrEmpty(UpdateOldAuthRowIds) Then
                                    UpdateOldAuthRowIds = String.Format("'{0}'", tbUpdData.Rows(0).Item("RowId").ToString)
                                Else
                                    UpdateOldAuthRowIds = String.Format("{0},'{1}'", UpdateOldAuthRowIds,
                                                                        tbUpdData.Rows(0).Item("RowId").ToString)
                                End If
                            End Using


                        End If



                    End Using

                Case AuthType.ErrorType
                    If (StopAll) OrElse (SourceType = FormatType.OldType) Then
                        If _dsInputData.Tables(tbStopAllName).Rows(0).Item("StopAll") Then
                            UpdateCount = DAO.ExecNqry(_DAL.StopSO106(True), New Object() {Me.LoginInfo.EntryName,
                                                                           CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                                                            OracleDate,
                                                                             ErrorMsg,
                                                                             ErrorMsg & Language.failDate & OracleDate.ToString("yyyy/MM/dd"),
                                                                            OracleDate.ToString("yyyyMMddHHmmss"),
                                                                             GetAccountId(ConText), GetACHCustId(ConText)})


                        Else
                            DAO.ExecNqry(_DAL.StopSO106(False), New Object() {Me.LoginInfo.EntryName,
                                                                           CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                                                            ErrorMsg,
                                                                             ErrorMsg & Language.failDate & OracleDate.ToString("yyyy/MM/dd"),
                                                                              OracleDate.ToString("yyyyMMddHHmmss"),
                                                                             GetAccountId(ConText), GetACHCustId(ConText)})

                        End If
                    End If
            End Select
        Catch ex As Exception
            Throw
        Finally

        End Try

        Return UpdateCount > 0


    End Function
    Private Overloads Function UpdateSO003(ByVal ConText As String, ByVal CustId As Integer,
                                           ByVal ACTHNo As String, ByVal CitemCodeStr As String) As Boolean
        Dim tbUpdData As DataTable = Nothing
        Try
            tbUpdData = DAO.ExecQry(_DAL.QueryUpdSO003Data(ACTHNo), New Object() {
                                    CustId,
                                    GetAccountId(ConText),
                                    GetACHCustId(ConText),
                                    ContextUpdDate})

            If tbUpdData.Rows.Count > 0 Then
                DAO.ExecNqry(_DAL.UpdateSO003(CitemCodeStr), New Object() {
                             tbUpdData.Rows(0).Item("BankCode"),
                             tbUpdData.Rows(0).Item("BANKNAME"),
                             tbUpdData.Rows(0).Item("ACCOUNTID"),
                             tbUpdData.Rows(0).Item("PTCode"),
                             tbUpdData.Rows(0).Item("PTName"),
                             tbUpdData.Rows(0).Item("CMCode"),
                             tbUpdData.Rows(0).Item("CMName"),
                             tbUpdData.Rows(0).Item("InvSeqNo"),
                             CustId})


            Else
                Return False
            End If
        Catch ex As Exception
            Throw
        Finally
            If tbUpdData IsNot Nothing Then
                tbUpdData.Dispose()
                tbUpdData = Nothing
            End If
        End Try
        Return True
    End Function
    Private Function UpdACHSO003C(ByVal MasterId As String) As Boolean
        Try
            Using tbSO106 = DAO.ExecQry(_DAL.QueryUniqueSO106(), New Object() {Integer.Parse(MasterId)})
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
            Throw
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
            Throw
        Finally
            If tbSO106 IsNot Nothing Then
                tbSO106.Dispose()
                tbSO106 = Nothing
            End If
        End Try
        Return True
    End Function
    Private Overloads Function UpdateSO003(ByVal ConText As String, ByVal CustId As Integer, ByVal rwSO106A As DataRow) As Boolean

        Return UpdateSO003(ConText, CustId, rwSO106A("ACHTNO"), rwSO106A("CitemCodeStr"))

    End Function
    Private Function UpdateSO106A(ByVal ReplyType As AuthType,
                                  ByVal tblSO106A As DataTable, ByVal ErrMsg As String) As Boolean
        Dim ErrNote As String = Nothing
        If SourceType = FormatType.OldType Then
            Return True
        End If
        Try
            Select Case ReplyType
                Case AuthType.Auth
                    For Each rwSO106A In tblSO106A.Rows
                        DAO.ExecNqry(_DAL.UpdateSO016A(AuthType.Auth), New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                                     Me.LoginInfo.EntryName, rwSO106A.Item("CTID")})
                    Next

                Case AuthType.CancelAuth
                    For Each rwSO106A In tblSO106A.Rows                        
                        Using tbStopAll As DataTable = DAO.ExecQry(_DAL.ChkCancelAuthStopAll, New Object() {rwSO106A.Item("MasterId")})
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
                        DAO.ExecNqry(_DAL.UpdateSO016A(AuthType.CancelAuth), New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                                     Me.LoginInfo.EntryName, rwSO106A.Item("CTID")})
                    Next

                Case AuthType.ErrorType
                    For Each rwSO106A In tblSO106A.Rows
                        DAO.ExecNqry(_DAL.UpdateSO016A(ReplyType), New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                                     Me.LoginInfo.EntryName,
                                                     ErrMsg & Language.failDate & OracleDate.ToShortDateString, rwSO106A.Item("CTID")})
                        If Integer.Parse(DAO.ExecSclr(_DAL.ChkSO106AAllFail,
                                   New Object() {rwSO106A.Item("AchtNO"),
                                                rwSO106A.Item("ACHDesc"),
                                                rwSO106A.Item("MasterId")})) = 0 Then
                            StopAll = True
                            Using tbSO106A As DataTable = DAO.ExecQry(_DAL.QuerySO106AErrNote,
                                                                      New Object() {rwSO106A.Item("MasterId")})
                                For Each rw As DataRow In tbSO106A.Rows
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
                            StopAll = False
                        End If
                    Next






                Case AuthType.OldAuth
                    For Each rwSO106A In tblSO106A.Rows
                        If Not String.IsNullOrEmpty(OldAuthNewRowId) Then
                            With rwSO106A
                                DAO.ExecNqry(_DAL.InsertSO106A, New Object() {
                                       OldAuthNewRowId,
                                       .Item("ACHTNO"),
                                       .Item("Notes"),
                                       .Item("CitemCodeStr"),
                                       .Item("CitemNameStr"),
                                       .Item("StopFlag"),
                                       .Item("StopDate"),
                                       Me.LoginInfo.EntryName,
                                       CableSoft.BLL.Utility.DateTimeUtility.GetDTString(OracleDate),
                                       Me.LoginInfo.EntryName,
                                       .Item("RecordType"),
                                       .Item("AuthorizeStatus"),
                                       .Item("AchDesc"),
                                       NewMasterIdSeq})
                            End With

                        End If
                    Next

            End Select


        Catch ex As Exception
            Throw
        Finally

        End Try

        Return True
    End Function

    Private Function IsAuthInOk(ByVal ConText As String, ByRef AuthErrorMsg As String) As Boolean
        Dim Result As Boolean = False
        Select Case ConText.Substring(107, 1)
            Case "0"
                AuthErrorMsg = Nothing
                Result = True
            Case "1" '                1 =印鑑不符
                AuthErrorMsg = Language.AuthErrorMsg1
            Case "2" '                2 = 無此帳號
                AuthErrorMsg = Language.AuthErrorMsg2
            Case "3" '                3 = 委繳戶統編不存在
                AuthErrorMsg = Language.AuthErrorMsg3
            Case "4" '                4 = 資料重覆
                AuthErrorMsg = Language.AuthErrorMsg4
            Case "5" '                5 = 原交易不存在
                AuthErrorMsg = Language.AuthErrorMsg5
            Case "6" '                6 = 電子資料與授書內容不符
                AuthErrorMsg = Language.AuthErrorMsg6
            Case "7" '                7 = 帳戶已結清
                AuthErrorMsg = Language.AuthErrorMsg7
            Case "8" '                8 = 印鑑不清
                AuthErrorMsg = Language.AuthErrorMsg8
            Case "A"
                AuthErrorMsg = Language.AuthErrorMsgA
            Case "B"
                AuthErrorMsg = Language.AuthErrorMsgB
            Case "C"
                AuthErrorMsg = Language.AuthErrorMsgC
            Case "D"
                AuthErrorMsg = Language.AuthErrorMsgD
            Case Else '                9 = 其他
                AuthErrorMsg = Language.AuthErrorMsg9
        End Select
        Return Result
    End Function
    Private Function GetContextStringACH(ByVal ConText As String) As String
        Return ConText.Substring(6, 3).Replace(" ", "").PadLeft(3, " ")
    End Function
    Private Function GetErrMsg(ByVal ConText As String, ByVal ErrDescription As String,
                               ByVal ID As String) As String
        Return String.Format(Language.GetErrMsg, ID, GetAccID(ConText), ErrDescription)
    End Function
    Private Function GetAccID(ByVal ConText As String) As String
        On Error Resume Next
        Dim ActLength As Object
        ActLength = DAO.ExecSclr(_DAL.QueryActLength, New Object() {GetBankCode(ConText)})
        If IsDBNull(ActLength) Then
            ActLength = 0
        Else
            ActLength = Integer.Parse(ActLength)
        End If
        Return Right(ConText.Substring(26, 14), ActLength).Trim()
    End Function
    Private Function GetACHCustId(ByVal ConText As String) As String
        Return ConText.Substring(50, 20).Replace(" ", "")
    End Function
    Private Function GetAccountId(ByVal ConText As String) As String
        Return ConText.Substring(26, 14).Replace(" ", "").PadLeft(14, "0")
    End Function
    Private Function GetBankCode(ByVal ConText As String) As String
        Return ConText.Substring(19, 7).Replace(" ", "")
    End Function
    Private Function WriteLogFile(ByVal FileContext As String) As String
        'Dim Path As String = System.Web.HttpContext.Current.Server.MapPath("~\") & TxtDirName
        Dim Path As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & TxtDirName
        Dim FileName As String = Me.LoginInfo.EntryId & "-" & Now.ToString("yyyyMMddHHmmssff") & ".Txt"
        System.IO.File.WriteAllText(String.Format("{0}\{1}", Path, FileName), FileContext, System.Text.Encoding.GetEncoding(950))
        Return FileName
    End Function
    Public Function DeleteLogFile(ByVal FileName As String) As Boolean
        Try
            Dim Path As String = System.Web.HttpContext.Current.Server.MapPath("~\") & TxtDirName
            FileName = String.Format("{0}\{1}", Path, FileName)
            System.IO.File.Delete(FileName)
        Catch ex As Exception
            Return True
        End Try
        Return True
    End Function
    Private Function chkReplyType(ByVal str As String) As AuthType
        Select Case str.Substring(70, 1).ToUpper
            Case "A"
                Return AuthType.Auth
            Case "O"
                Return AuthType.OldAuth
            Case "D"
                Return AuthType.CancelAuth
            Case Else
                Return AuthType.ErrorType
        End Select
    End Function

    Public Function QueryAllData() As DataSet
        Dim dsReturn As New DataSet
        Dim tbCompCode As DataTable = Nothing
        Dim tbFormatType As DataTable = Nothing
        Dim tbBankId As DataTable = Nothing
        Dim tbBillHeadFmt As DataTable = Nothing
        Dim tbCitemCode As DataTable = Nothing
        Dim tbCD068 As DataTable = Nothing
        Try
            tbCompCode = QueryCompCode.Copy
            tbCompCode.TableName = tbCompCodeName
            tbFormatType = QueryFormatType.Copy
            tbFormatType.TableName = tbFormatTypeName
            tbBankId = QueryBankId.Copy
            tbBankId.TableName = tbBankIdName
            tbBillHeadFmt = QueryBillHeadFmt.Copy
            tbBillHeadFmt.TableName = tbBillHeadFmtName
            tbCitemCode = QueryCitemCode.Copy
            tbCitemCode.TableName = tbCitemCodeName
            tbCD068 = QueryCD068.Copy
            tbCD068.TableName = tbCD068Name
            With dsReturn.Tables
                .Add(tbBankId)
                .Add(tbBillHeadFmt)
                .Add(tbCD068)
                .Add(tbCitemCode)
                .Add(tbCompCode)
                .Add(tbFormatType)
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
    Public Function QueryFormatType() As DataTable
        Return DAO.ExecQry(_DAL.QueryFormatType, New Object() {Me.LoginInfo.CompCode})
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
                    Me.DAO = Nothing
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
