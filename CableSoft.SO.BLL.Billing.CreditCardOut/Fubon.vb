Imports CableSoft.BLL.Utility
Imports System.Web
Imports System.Xml
Imports System.Data.Common
Public Class Fubon
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New FubonDALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.SO3272A3Language
    Private FNowDate As Date = Date.Now
    Private strSO033Where As String = Nothing
    Private blnUpdUCCode As Boolean = False
    Private strUCCode As String = Nothing
    Private strUCName As String = Nothing
    Private strCMCode As String = Nothing
    Private strCMName As String = Nothing
    Private Const TxtDirName As String = "TXT"
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
    Public Function QueryCD019(ByVal BillHeadFmt As String) As DataSet
        Try
            Dim ds As New DataSet
            ds.Tables.Add(DAO.ExecQry(_DAL.QueryCD019, New Object() {BillHeadFmt}).Copy)
            Return ds.Copy

        Catch ex As Exception
            Throw
        End Try
        

    End Function

    Private Sub DropView(ByVal viewName As String)


        Try
            DAO.ExecNqry(_DAL.dropView(viewName))
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub
    
    Public Function ChkAuthority(ByVal Mid As String) As RIAResult
        Dim result As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        Try
            Using obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
                result = obj.ChkPriv(LoginInfo.EntryId, Mid)
                obj.Dispose()
            End Using
           

        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -2
        Finally

        End Try
        Return result

    End Function
    Private Function getConditionDataSet(ByVal dsSource As DataSet) As DataSet
        Dim dsResult As New DataSet
        Dim dtResult As New DataTable("Condition")
        dtResult.Columns.Add("ShouldDate1", GetType(String))
        dtResult.Columns.Add("ShouldDate2", GetType(String))
        dtResult.Columns.Add("CreateTime1", GetType(String))
        dtResult.Columns.Add("CreateTime2", GetType(String))
        dtResult.Columns.Add("CMCode", GetType(String))
        dtResult.Columns.Add("AreaCode", GetType(String))
        dtResult.Columns.Add("ServCode", GetType(String))
        dtResult.Columns.Add("ClctEn", GetType(String))
        dtResult.Columns.Add("OldClctEn", GetType(String))
        dtResult.Columns.Add("PayKind", GetType(String))
        dtResult.Columns.Add("CustId", GetType(String))
        dtResult.Columns.Add("BILLNOTYPE", GetType(String))
        dtResult.Columns.Add("CreateEn", GetType(String))
        dtResult.Columns.Add("MDUIDE", GetType(String))
        dtResult.Columns.Add("MDUIDN", GetType(String))
        dtResult.Columns.Add("ISOTHER", GetType(String))
        dtResult.Columns.Add("UCCODE", GetType(String))
        dtResult.Columns.Add("BillHeadFmt", GetType(String))
        dtResult.Columns.Add("BANKCODE", GetType(String))
        dtResult.Columns.Add("ExcI", GetType(String))
        dtResult.Columns.Add("CustStatusCode", GetType(String))
        dtResult.Columns.Add("ClassCode1", GetType(String))
        dtResult.Columns.Add("AMduId", GetType(String))
        dtResult.Columns.Add("ACCEPTDATE", GetType(String))
        dtResult.Columns.Add("BILLMEMO", GetType(String))
        dtResult.Columns.Add("ClientId", GetType(String))
        dtResult.Columns.Add("AUTHBATCH", GetType(String))
        dtResult.Columns.Add("STORENUM", GetType(String))
        dtResult.Columns.Add("IsZero", GetType(String))
        dtResult.Columns.Add("IsFubonIntegrate", GetType(String))
        dtResult.Columns.Add("IGNORCREDITCARD", GetType(String))
        Dim rwResult As DataRow = dtResult.NewRow
        With dsSource.Tables("Condition")
            For Each rw As DataRow In .Rows
                Select Case rw.Item("FieldName").ToString.ToUpper
                    Case "ShouldDate_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ShouldDate1") = rw.Item("FIELDVALUE")
                        End If
                    Case "ShouldDate_2".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ShouldDate2") = rw.Item("FIELDVALUE")
                        End If
                    Case "CREATETIME_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("CreateTime1") = rw.Item("FIELDVALUE")
                        End If
                    Case "CREATETIME_2".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("CreateTime2") = rw.Item("FIELDVALUE")
                        End If
                    Case "CMCode_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("CMCode") = rw.Item("FIELDVALUE")
                        End If
                    Case "AreaCode_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("AreaCode") = rw.Item("FIELDVALUE")
                        End If
                    Case "ServCode_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ServCode") = rw.Item("FIELDVALUE")
                        End If
                    Case "ClctEn_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ClctEn") = rw.Item("FIELDVALUE")
                        End If
                    Case "OldClctEn_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("OldClctEn") = rw.Item("FIELDVALUE")
                        End If
                    Case "PayType_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("PayKind") = rw.Item("FIELDVALUE")
                        End If
                    Case "BillType_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("BILLNOTYPE") = rw.Item("FIELDVALUE")
                        End If
                    Case "CustId_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("CustId") = rw.Item("FIELDVALUE")
                        End If
                    Case "CreateEn_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("CreateEn") = rw.Item("FIELDVALUE")
                        End If
                    Case "EMduid_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("MDUIDE") = rw.Item("FIELDVALUE")
                        End If
                    Case "NMduid_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("MDUIDN") = rw.Item("FIELDVALUE")
                        End If
                    Case "Normal_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ISOTHER") = rw.Item("FIELDVALUE")
                        End If
                    Case "FubonIntegrate_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("IsFubonIntegrate") = rw.Item("FIELDVALUE")
                        End If
                    Case "UCCode_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("UCCODE") = rw.Item("FIELDVALUE")
                        End If
                    Case "BillHeadFmt_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("BillHeadFmt") = rw.Item("FIELDDESC")
                        End If
                    Case "BankCode_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("BANKCODE") = rw.Item("FIELDVALUE")
                        End If
                    Case "ExcI_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ExcI") = rw.Item("FIELDVALUE")
                        End If
                    Case "IgnorCredit_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("IGNORCREDITCARD") = rw.Item("FIELDVALUE")
                        End If
                    Case "CustStatus_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("CustStatusCode") = rw.Item("FIELDVALUE")
                        End If
                    Case "CustClass_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ClassCode1") = rw.Item("FIELDVALUE")
                        End If
                    Case "AMduId_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("AMduId") = rw.Item("FIELDVALUE")
                        End If
                    Case "SendDate_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ACCEPTDATE") = rw.Item("FIELDVALUE")
                        End If
                    Case "AuthBatch_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("AUTHBATCH") = rw.Item("FIELDVALUE")
                        End If
                    Case "ClientId_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("ClientId") = rw.Item("FIELDVALUE")
                        End If
                    Case "BillMemo_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("BILLMEMO") = rw.Item("FIELDVALUE")
                        End If
                    Case "StoreNum_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("STORENUM") = rw.Item("FIELDVALUE")
                        End If
                    Case "IsZero_1".ToUpper
                        If Not DBNull.Value.Equals(rw.Item("FIELDVALUE")) Then
                            rwResult("IsZero") = rw.Item("FIELDVALUE")
                        End If
                End Select
            Next
        End With
        dtResult.Rows.Add(rwResult)
        dsResult.Tables.Add(dtResult)
        Return dsResult
    End Function
    Public Function Execute(ByVal dsSource As DataSet, ByVal SEQNO As Integer) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim blnBeginTransaction As Boolean = False
        Dim riaResult As New RIAResult With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        Dim aSQL As String = Nothing
        Dim sqlTmpViewName = Nothing
        Dim isCrossCustCombine As Boolean = False
        Dim lngTotalCount As Integer = 0
        Dim lngErrCount As Integer = 0
        Dim lngCount As Integer = 0
        Dim sbErr As New System.Text.StringBuilder
        Dim sbOK As New System.Text.StringBuilder
        Dim aCustId As String = Nothing
        Dim aCustName As String = Nothing
        Dim aPayKindName As String = Nothing
        Dim aCreditDate As String = Nothing
        Dim strBaseKey As String = Nothing
        Dim strPrepareEncrypt As String = Nothing
        Dim strHeader As String = Nothing
        Dim strFinal As String = Nothing
        Dim ParentSeqNo As Double = -1
        Dim BllUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        FNowDate = Date.Now
        Dim RunTime As New Stopwatch()
        RunTime.Start()
        Dim dsCondition As DataSet = getConditionDataSet(dsSource)
        Dim IsFubonIntegrate As Boolean = False
        Try
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
            DAO.AutoCloseConn = False
            If blnAutoClose Then
                Dim aAction As String = Nothing
                aAction = Language.ClientInfoString
                CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, aAction)
            End If
            isCrossCustCombine = Integer.Parse(DAO.ExecSclr(_DAL.QueryIsCrossCustCombine)) = 1
            If isCrossCustCombine Then
                Dim viewNum As String = DAO.ExecSclr(_DAL.GetViewName)
                sqlTmpViewName = "TMP_" & viewNum
                DropView(sqlTmpViewName)
            End If
            strSO033Where = Nothing
            Dim aWhere As String = _DAL.BuildCondition(dsCondition, Me.LoginInfo.CompCode,strSO033Where)
            Using tbUpdUCCode As DataTable = DAO.ExecQry(_DAL.QueryUpdUCCode)
                If tbUpdUCCode.Rows.Count > 0 Then
                    strUCCode = tbUpdUCCode.Rows(0).Item("CodeNo")
                    strUCName = tbUpdUCCode.Rows(0).Item("Description")
                End If
            End Using
            If Integer.Parse(dsCondition.Tables(0).Rows(0).Item("IsFubonIntegrate")) = 1 Then
                IsFubonIntegrate = True
                'Using tbCD031 As DataTable = DAO.ExecQry(_DAL.QueryCD031RefNo5)
                '    If tbCD031.Rows.Count > 0 Then
                '        IsFubonIntegrate = True
                '        strCMCode = tbCD031.Rows(0).Item("CodeNo")
                '        strCMName = tbCD031.Rows(0).Item("Description")
                '    End If
                'End Using
            End If
            strCMCode = Nothing
            strCMName = Nothing


            aSQL = _DAL.QuerySO033Data(strSO033Where,
                                       Integer.Parse(dsCondition.Tables("Condition").Rows(0).Item("IsZero")),
                                       isCrossCustCombine, LoginInfo.CompCode)
            If Not String.IsNullOrEmpty(sqlTmpViewName) Then
                DAO.ExecNqry(_DAL.CreateView(sqlTmpViewName, aSQL))
                aSQL = _DAL.QueryViewSQL(sqlTmpViewName)
            End If
            If SEQNO = 0 Then
                BllUtility.InsertProgramLog("SO3272A3", dsSource, SO.BLL.Utility.ExecType.TextFile,
                                        Nothing, True, Nothing, ParentSeqNo)
            End If
            Dim aMediaIsNull As String = _DAL.QueryMediaIsNull(strSO033Where, _
                                                              Integer.Parse(dsCondition.Tables("Condition").Rows(0).Item("IsZero")))
            blnBeginTransaction = True
            Using tbMediaIsNull As DataTable = DAO.ExecQry(aMediaIsNull)
                For i As Integer = 0 To tbMediaIsNull.Rows.Count - 1
                    Dim strSequenceNumber As String = DAO.ExecSclr(_DAL.GetInvoiceNo2("SO033_MediaBillNo"))
                    DAO.ExecNqry(_DAL.UpdMediabillNo(strSO033Where), New Object() {
                                    strSequenceNumber, tbMediaIsNull.Rows(i)("BillNO"), tbMediaIsNull.Rows(i)("AccountNO")
                                 })
                Next
            End Using
            Using tbSO033 As DataTable = DAO.ExecQry(aSQL)
                'lngTotalCount = tbSO033.Rows.Count
                Dim ds As New DataTable
                Dim o As New System.IO.MemoryStream()



                If tbSO033.Rows.Count = 0 Then
                    'No Data
                    RunTime.Stop()
                    Math.Round(RunTime.Elapsed.TotalSeconds, 1)
                    riaResult.ErrorCode = -1
                    riaResult.ErrorMessage = String.Format(Language.PrcResult, 0, 0, Math.Round(RunTime.Elapsed.TotalSeconds, 1))
                    If SEQNO = 0 Then
                        DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, riaResult.ErrorMessage, DBNull.Value, ParentSeqNo})
                    End If
                    Return riaResult
                Else
                    aPayKindName = DAO.ExecSclr(_DAL.QueryPayKindName)
                    For i As Integer = 0 To tbSO033.Rows.Count - 1
                        With tbSO033.Rows(i)
                            '-----------------------------check data---------------------------------------------------
                            If isCrossCustCombine Then
                                If Integer.Parse(.Item("CustId") & "") = -1 Then
                                    lngErrCount = lngErrCount + 1
                                    sbErr.AppendLine(String.Format(Language.NoFoundComCustId, .Item("BillNo")))
                                    Continue For
                                End If
                            End If
                            Using tbCustId As DataTable = DAO.ExecQry(_DAL.QueryCustIDAndName, New Object() {.Item("BILLNO")})
                                aCustId = -1
                                aCustName = String.Empty
                                If tbCustId.Rows.Count > 0 Then
                                    aCustId = tbCustId.Rows(0).Item("CustId")
                                    aCustName = tbCustId.Rows(0).Item("CustName")
                                End If
                                tbCustId.Dispose()
                            End Using
                            If (Not IsPayKindOK(Integer.Parse(.Item("PayKind")), .Item("REALSTOPDATE"),
                                                    dsCondition.Tables("Condition").Rows(0).Item("ACCEPTDATE"))) Then
                                lngErrCount += 1
                                sbErr.AppendLine(String.Format(Language.WatchDateWrong, .Item("BillNo"),
                                                               aCustId, aCustName, .Item("RealStopDate"), aPayKindName))

                                Continue For
                            End If

                            If DBNull.Value.Equals(.Item("AccountNO")) Then
                                lngErrCount += 1
                                sbErr.AppendLine(String.Format(Language.CreditCardEmpty, .Item("BillNo"), aCustName))
                                Continue For
                            End If
                            If DBNull.Value.Equals(.Item("CardExpDate")) Then
                                lngErrCount += 1
                                sbErr.AppendLine(String.Format(Language.CreditCardDateEmpty, .Item("BillNo"),
                                                               aCustId, aCustName, .Item("RealStopDate"), aPayKindName))

                                Continue For
                            End If
                            aCreditDate = Left(.Item("CardExpDate"), 4)
                            If .Item("CardExpDate").ToString.Length = 5 Then
                                aCreditDate = aCreditDate & "/0" & Right(.Item("CardExpDate"), 1)
                            Else
                                aCreditDate = aCreditDate & "/" & Right(.Item("CardExpDate"), 2)
                            End If

                            If Integer.Parse(aCreditDate.Replace("/", "") & "12") <
                                    Integer.Parse(dsCondition.Tables("Condition").Rows(0).Item("ACCEPTDATE").ToString.Replace("/", "")) Then
                                lngErrCount += 1
                                sbErr.AppendLine(String.Format(Language.CreditCardDue, .Item("BillNo"),
                                                             aCustName, .Item("AccountNO"), aCreditDate))

                                If Integer.Parse(dsCondition.Tables("Condition").Rows(0).Item("IGNORCREDITCARD")) = 0 Then Continue For
                            End If
                            If (Integer.Parse(.Item("ShouldAmt").ToString & "") <= 0) AndAlso
                                    (Integer.Parse(dsCondition.Tables("Condition").Rows(0).Item("ISZERO")) = 0) Then
                                lngErrCount += 1
                                sbErr.AppendLine(String.Format(Language.ZeroAmt, .Item("BillNo"),
                                                              aCustName, .Item("ShouldAmt")))

                                Continue For
                            End If
                            '--------------------------------------------------------------------------------------------------------------
                            lngCount += 1
                            strBaseKey = dsCondition.Tables("Condition").Rows(0).Item("ACCEPTDATE").ToString.Replace("/", "") & "080000" & _
                                     Right("0000" & dsCondition.Tables("Condition").Rows(0).Item("AUTHBATCH").ToString(), 4) & _
                                     .Item("BillNo").ToString.Substring(0, 8)

                            If sbOK.Length > 0 Then sbOK.Append(Environment.NewLine)
                            '1-1
                            sbOK.Append("D")
                            '2-17
                            sbOK.Append(Left(dsCondition.Tables("Condition").Rows(0).Item("ClientId").ToString & Space(16), 16))
                            '18-42
                            sbOK.Append(Left(.Item("BillNo").ToString & Space(25), 25))
                            '43-44
                            sbOK.Append("00")
                            '45-45
                            sbOK.Append("0")
                            '45-56
                            sbOK.Append(Right("00000000000" & .Item("ShouldAmt"), 11))

                            strPrepareEncrypt = Left(.Item("AccountNO").ToString & Space(19), 19) & _
                                            Right(.Item("CardExpDate").ToString, 2) & Right(Left(.Item("CardExpDate").ToString, 4), 2) & _
                                            Space(3)
                            '57-92
                            sbOK.Append(System.Convert.ToBase64String(xorEncrpty(strPrepareEncrypt, strBaseKey)))
                            '93-94
                            sbOK.Append(Space(2))
                            '95-134
                            sbOK.Append(Space(40))
                            '135-144
                            sbOK.Append(Space(10))
                            '145-184
                            If DBNull.Value.Equals(dsCondition.Tables("Condition").Rows(0).Item("BILLMEMO").ToString) Then
                                sbOK.Append(Space(40))
                            Else
                                Dim btyText() As Byte = System.Text.Encoding.Default.GetBytes(dsCondition.Tables("Condition").Rows(0).Item("BILLMEMO").ToString)
                                If btyText.Length > 40 Then
                                    sbOK.Append(System.Text.Encoding.Default.GetString(btyText, 0, 40))
                                Else
                                    sbOK.Append(System.Text.Encoding.Default.GetString(btyText, 0, btyText.Length))
                                End If
                            End If
                            '185-192
                            sbOK.Append(Space(8))
                            '193-300
                            '8575 stop checking the bank by kin 2020/02/21
                            If IsFubonIntegrate AndAlso 1 = 0 Then
                                sbOK.Append(Space(13))
                                sbOK.Append(Left(Right((Integer.Parse(Date.Now.ToString("yyyyMMdd")) - 19110000).ToString(), 4) &
                                                  Right("00" & dsCondition.Tables("Condition").Rows(0).Item("AUTHBATCH").ToString(), 1) &
                                                 .Item("BillNo").ToString & "0000000000000000", 16))

                                ' dsCondition.Tables("Condition").Rows(0).Item("AUTHBATCH").ToString

                                sbOK.Append(Space(79))
                            Else
                                sbOK.Append(Space(108))
                            End If

                            'If IsFubonIntegrate AndAlso Not String.IsNullOrEmpty(strCMCode) Then
                            '    DAO.ExecNqry(_DAL.UpdUCCode(IsFubonIntegrate), New Object() {strUCCode,
                            '                                          strUCName,
                            '                                          LoginInfo.EntryName,
                            '                                           CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                            '                                           strCMCode, strCMName,
                            '                                          .Item("BillNo"), .Item("AccountNo")})
                            'Else
                            '    DAO.ExecNqry(_DAL.UpdUCCode(IsFubonIntegrate AndAlso Not String.IsNullOrEmpty(strCMCode)), New Object() {strUCCode,
                            '                                            strUCName,
                            '                                            LoginInfo.EntryName,
                            '                                             CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                            '                                            .Item("BillNo"), .Item("AccountNo")})
                            'End If
                            DAO.ExecNqry(_DAL.UpdUCCode(False), New Object() {strUCCode,
                                                                        strUCName,
                                                                        LoginInfo.EntryName,
                                                                         CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                                                                        .Item("BillNo"), .Item("AccountNo")})
                            lngTotalCount = lngTotalCount + Integer.Parse(.Item("ShouldAmt"))

                        End With

                    Next

                End If
                strHeader = String.Empty
                strFinal = String.Empty
                strHeader = "HBA" & Left(dsCondition.Tables("Condition").Rows(0).Item("STORENUM").ToString & Space(16), 16) & _
                            dsCondition.Tables("Condition").Rows(0).Item("ACCEPTDATE").ToString.Replace("/", "") & "080000" & _
                            Right("0000" & dsCondition.Tables("Condition").Rows(0).Item("AUTHBATCH").ToString, 4) & Space(263)

                strFinal = "T"
                strFinal = "T" & Right("00000000" & lngCount, 8) & _
                            Right("00000000" & lngCount, 8) & _
                           "00000000" & _
                           Right("0000000000000" & lngTotalCount, 13) & _
                           Right("0000000000000" & lngTotalCount, 13) & _
                           "0000000000000" & _
                            "00000000" & _
                            "00000000" & _
                            "00000000" & _
                            "00000000" & _
                            "0000000000000" & _
                            "0000000000000" & _
                            "0000000000000" & _
                            "0000000000000" & _
                             "00000000" & _
                             "0000000000000" & Space(131)


            End Using

            Dim fileName As String = "BA" & dsCondition.Tables("Condition").Rows(0).Item("STORENUM").ToString & "_" & _
                       dsCondition.Tables("Condition").Rows(0).Item("ACCEPTDATE").ToString.Replace("/", "") & _
                       Right("0000" & dsCondition.Tables("Condition").Rows(0).Item("AUTHBATCH").ToString, 4) & ".dat"
            Dim downladFile As String = Me.LoginInfo.EntryId & "-" & Now.ToString("yyyyMMddHHmmssff") & ".zip"
            Dim zipFile As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & "\" & TxtDirName & "\" & downladFile


            Using zip As New Ionic.Zip.ZipFile(zipFile, _
                                                    System.Text.Encoding.GetEncoding(950))
                If sbErr.Length > 0 Then
                    zip.AddEntry("ErrorLog", sbErr.ToString)
                End If
                If sbOK.Length > 0 Then
                    zip.AddEntry(fileName, strHeader & Environment.NewLine & sbOK.ToString & Environment.NewLine & strFinal)
                End If

                zip.Save()
            End Using
            riaResult.ResultBoolean = True
            riaResult.ErrorMessage = Nothing
            riaResult.DownloadFileName = TxtDirName & "\" & downladFile
            RunTime.Stop()
            Math.Round(RunTime.Elapsed.TotalSeconds, 1)
            riaResult.ResultXML = riaResult.ResultXML & _
                    String.Format(Language.PrcResult, lngCount, lngErrCount, Math.Round(RunTime.Elapsed.TotalSeconds, 1))
            If SEQNO = 0 Then
                DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, riaResult.ResultXML, riaResult.DownloadFileName, ParentSeqNo})
            End If
            If blnAutoClose Then
                trans.Commit()
            End If

        Catch ex As Exception
            If (trans IsNot Nothing) AndAlso (blnAutoClose) Then
                If blnBeginTransaction Then
                    trans.Rollback()
                End If
            End If
            riaResult.ResultBoolean = False
            riaResult.ErrorCode = -999
            riaResult.ErrorMessage = ex.ToString
        Finally

            If Not String.IsNullOrEmpty(sqlTmpViewName) Then
                DropView(sqlTmpViewName)
            End If
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                DAO.AutoCloseConn = True
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
            End If
            If dsCondition IsNot Nothing Then
                dsCondition.Dispose()
                dsCondition = Nothing
            End If
        End Try
        If _DAL IsNot Nothing Then
            _DAL.Dispose()
            _DAL = Nothing
        End If
        If BllUtility IsNot Nothing Then
            BllUtility.Dispose()
            BllUtility = Nothing
        End If
        Return riaResult
    End Function
    Private Function WriteText(ByVal path As String, ByVal OKContext As String, ErrConText As String, isGateway As Boolean)
        Try

        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function xorEncrpty(ByVal source As String, Key As String) As Byte()
        Try
            Dim text As Byte() = System.Text.Encoding.ASCII.GetBytes(source)
            Dim res As Byte() = New Byte(text.Length - 1) {}
            Dim c As Integer = 0
            Dim keyBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(Key)
            For c = 0 To UBound(Text)
                res(c) = CByte((Text(c)) Xor keyBytes(c Mod (UBound(keyBytes) + 1)))
            Next
            Return res
        Catch ex As Exception
            Throw
        End Try

    End Function
    Private Function xorEncrpty(text() As Byte, Key As String) As Byte()
        Try
            Dim res As Byte() = New Byte(text.Length - 1) {}
            Dim c As Integer = 0
            Dim keyBytes As Byte() = System.Text.Encoding.Unicode.GetBytes(Key)
            For c = 0 To UBound(text)
                res(c) = CByte((text(c)) Xor keyBytes(c Mod (UBound(keyBytes) + 1)))
            Next
            Return res
        Catch ex As Exception
            Throw
        End Try

    End Function

    
    Private Function IsPayKindOK(ByVal PayKind As Integer, ByVal REALSTOPDATE As String, ByVal AcceptDate As String) As Boolean
        If PayKind = 0 Then Return True
        If Integer.Parse(REALSTOPDATE.Replace("/", "")) <= Integer.Parse(AcceptDate.Replace("/", "")) Then
            Return True
        Else
            Return False
        End If
      
    End Function
    Public Function QueryAllData() As DataSet
        Dim dsReturn As New DataSet
        Dim tbCompCode As DataTable = Nothing
        Dim tbCD068 As DataTable = Nothing
        Dim tbCD018 As DataTable = Nothing
        Dim tbCD031 As DataTable = Nothing
        Dim tbCD031REFNO5 As DataTable = Nothing
        Dim tbCD001 As DataTable = Nothing
        Dim tbCD002 As DataTable = Nothing
        Dim tbCD035 As DataTable = Nothing
        Dim tbCD004 As DataTable = Nothing
        Dim tbCM003 As DataTable = Nothing
        Dim tbCD013 As DataTable = Nothing
        Dim tbBillType As DataTable = Nothing
        Dim tbCD112 As DataTable = Nothing
        Dim tbSO041 As DataTable = Nothing
        Dim tbSO202 As DataTable = Nothing
        Dim tbSO017 As DataTable = Nothing
        Dim tbOtherCondition As New DataTable("OTHER")
        Dim xmlString As Object = DAO.ExecSclr(_DAL.QuerySO1108A(), New Object() {Me.LoginInfo.EntryId})
        Dim dsSO108A As New DataSet
        If xmlString IsNot Nothing AndAlso Not DBNull.Value.Equals(xmlString) Then
            Dim xmlRd As System.IO.TextReader = New System.IO.StringReader(xmlString)
            dsSO108A.ReadXml(xmlRd)
        End If
        'dtResult.Columns.Add("ACCEPTDATE", GetType(String))
        'dtResult.Columns.Add("BILLMEMO", GetType(String))
        'dtResult.Columns.Add("ClientId", GetType(String))
        'dtResult.Columns.Add("AUTHBATCH", GetType(String))
        'dtResult.Columns.Add("STORENUM", GetType(String))
        tbOtherCondition.Columns.Add("CLIENTID", GetType(String))
        tbOtherCondition.Columns.Add("AUTHBATCH", GetType(String))
        tbOtherCondition.Columns.Add("STORENUM", GetType(String))
        If dsSO108A.Tables.Count > 0 AndAlso dsSO108A.Tables(0).Rows.Count > 0 Then
            Dim rwNew As DataRow = tbOtherCondition.NewRow
            For Each rwSO108A As DataRow In dsSO108A.Tables(0).Rows
                Select Case rwSO108A.Item("FieldName").ToString.ToUpper
                    Case "StoreNum_1".ToUpper
                        rwNew.Item("STORENUM") = rwSO108A.Item("FieldValue").ToString
                    Case "ClientId_1".ToUpper
                        rwNew.Item("CLIENTID") = rwSO108A.Item("FieldValue").ToString
                    Case "AUTHBATCH_1".ToUpper
                        rwNew.Item("AUTHBATCH") = rwSO108A.Item("FieldValue").ToString
                End Select
            Next
            tbOtherCondition.Rows.Add(rwNew)
        End If
        Try
            tbCompCode = DAO.ExecQry(_DAL.QueryCompCode, New Object() {Me.LoginInfo.EntryId}).Copy
            tbCompCode.TableName = "COMPCODE"
            tbCD068 = DAO.ExecQry(_DAL.QueryCD068).Copy
            tbCD068.TableName = "CD068"
            tbCD018 = DAO.ExecQry(_DAL.QueryCD018, New Object() {Me.LoginInfo.CompCode}).Copy
            tbCD018.TableName = "CD018"
            tbCD031 = DAO.ExecQry(_DAL.QueryCD031).Copy
            tbCD031.TableName = "CD031"
            tbCD031REFNO5 = DAO.ExecQry(_DAL.QueryCD031REFNO5).Copy
            tbCD031REFNO5.TableName = "CD031REFNO5"
            tbCD001 = DAO.ExecQry(_DAL.QueryCD001).Copy
            tbCD001.TableName = "CD001"
            tbCD002 = DAO.ExecQry(_DAL.QueryCD002).Copy
            tbCD002.TableName = "CD002"
            tbCD035 = DAO.ExecQry(_DAL.QueryCD035).Copy
            tbCD035.TableName = "CD035"
            tbCD004 = DAO.ExecQry(_DAL.QueryCD004).Copy
            tbCD004.TableName = "CD004"
            tbCM003 = DAO.ExecQry(_DAL.QueryCM003).Copy
            tbCM003.TableName = "CM003"
            tbCD013 = DAO.ExecQry(_DAL.QueryCD013).Copy
            tbCD013.TableName = "CD013"
            tbBillType = DAO.ExecQry(_DAL.QueryBillType).Copy
            tbBillType.TableName = "BILLTYPE"
            tbCD112 = DAO.ExecQry(_DAL.QueryCD112).Copy
            tbCD112.TableName = "CD112"
            tbSO041 = DAO.ExecQry(_DAL.QuerySO041).Copy
            tbSO041.TableName = "SO041"
            tbSO202 = DAO.ExecQry(_DAL.QuerySO202).Copy
            tbSO202.TableName = "SO202"
            tbSO017 = DAO.ExecQry(_DAL.QuerySO017).Copy
            tbSO017.TableName = "SO017"
            With dsReturn.Tables
                .Add(tbCompCode)
                .Add(tbCD068)
                .Add(tbCD018)
                .Add(tbCD031)
                .Add(tbCD001)
                .Add(tbCD002)
                .Add(tbCD035)
                .Add(tbCD004)
                .Add(tbCM003)
                .Add(tbCD013)
                .Add(tbBillType)
                .Add(tbCD112)
                .Add(tbSO041)
                .Add(tbSO202)
                .Add(tbSO017)
                .Add(tbOtherCondition)
                .Add(tbCD031REFNO5)
            End With
        Catch ex As Exception
            Throw
        Finally
            If tbOtherCondition IsNot Nothing Then
                tbOtherCondition.Dispose()
                tbOtherCondition = Nothing
            End If
            If tbCompCode IsNot Nothing Then
                tbCompCode.Dispose()
                tbCompCode = Nothing
            End If
            If tbCD068 IsNot Nothing Then
                tbCD068.Dispose()
                tbCD068 = Nothing
            End If
            If tbCD018 IsNot Nothing Then
                tbCD018.Dispose()
                tbCD018 = Nothing
            End If
            If tbCD031 IsNot Nothing Then
                tbCD031.Dispose()
                tbCD031 = Nothing
            End If
            If tbCD001 IsNot Nothing Then
                tbCD001.Dispose()
                tbCD001 = Nothing
            End If
            If tbCD002 IsNot Nothing Then
                tbCD002.Dispose()
                tbCD002 = Nothing
            End If
            If tbCD035 IsNot Nothing Then
                tbCD035.Dispose()
                tbCD035 = Nothing
            End If
            If tbCD004 IsNot Nothing Then
                tbCD004.Dispose()
                tbCD004 = Nothing
            End If
            If tbCM003 IsNot Nothing Then
                tbCM003.Dispose()
                tbCM003 = Nothing
            End If
            If tbCD013 IsNot Nothing Then
                tbCD013.Dispose()
                tbCD013 = Nothing
            End If
            If tbBillType IsNot Nothing Then
                tbBillType.Dispose()
                tbBillType = Nothing
            End If
            If tbCD112 IsNot Nothing Then
                tbCD112.Dispose()
                tbCD112 = Nothing
            End If
            If tbSO041 IsNot Nothing Then
                tbSO041.Dispose()
                tbSO041 = Nothing
            End If
            If tbSO202 IsNot Nothing Then
                tbSO202.Dispose()
                tbSO202 = Nothing
            End If
            If tbSO017 IsNot Nothing Then
                tbSO017.Dispose()
                tbSO017 = Nothing
            End If
            If tbCD031REFNO5 IsNot Nothing Then
                tbCD031REFNO5.Dispose()
                tbCD031REFNO5 = Nothing
            End If
            If dsSO108A IsNot Nothing Then
                dsSO108A.Dispose()
                dsSO108A = Nothing
            End If

        End Try

        Return dsReturn
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
