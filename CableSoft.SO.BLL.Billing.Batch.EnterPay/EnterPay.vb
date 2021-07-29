Imports CableSoft.BLL.Utility
Imports System.Web
Imports System.Xml
Imports System.Data.Common
Imports System.IO

Public Class EnterPay
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New EnterPayDALMultiDB(Me.LoginInfo.Provider)
    Private FNowDate As DateTime
    Private Const tbOKName As String = "OK"
    Private Const tbErrorName As String = "Error"
    Private Const tbImportErrName As String = "ImportError"
    Private Const ErrorCodeFieldName As String = "ErrorCode"
    Private Const ErrorNameFieldName As String = "ErrorName"
    Private Const WaringMsgFieldName As String = "WaringMsg"
    Private Const tbTotalInfoName As String = "TotalInfo"
    Private Const BillCount As String = "BillCount"
    Private Const AmtCount As String = "AmtCount"
    Private Const tbInfoName As String = "Info"
    Private Const tbTemp As String = "Temp"
    Private EntryNoCount As Integer = 0
    Private Language As New CableSoft.BLL.Language.SO61.EnterPayLanguage
    '匯入Excel方法
    'Dim CurrentDir As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory()
    '            RetData = CableSoft.Utility.Heterogeneous.xls.ToDataTable(String.Format("{0}\{1}", CurrentDir, FileName), , , , 1)

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
                'result.ErrorMessage = Language.NoPermission
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
    Public Function WriteMsgForDownload(ByVal Msg As String) As RIAResult
        Dim fileName As String = Now.ToString("yyyyMMddHHmmssff") & "-SO3311G.zip"
        Dim retFileName As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & "TXT\" & fileName
        Try

            Using zip As New Ionic.Zip.ZipFile(retFileName, System.Text.Encoding.GetEncoding(950))
                zip.AddEntry("SO3311Err.txt", Msg)
                zip.Save()
            End Using

            'Using txt As New XmlTextWriter(retFileName, System.Text.Encoding.GetEncoding(950))
            '    txt.WriteString(Msg)
            '    txt.Flush()
            '    txt.Close()
            'End Using
            Return New RIAResult With {.ResultBoolean = True, .ResultXML = "TXT\" & fileName, .DownloadFileName = "TXT\" & fileName}
        Catch ex As Exception
            Return New RIAResult With {.ResultBoolean = False, .ResultXML = ex.ToString, .ErrorMessage = ex.ToString}
        End Try
    End Function
    Public Function ImportExcel(ByVal FileName As String, ByVal CitemPara As DataTable, ByVal UCRefNo As Integer) As DataSet
        Dim tbExcel As DataTable = Nothing
        Dim dsReturn As DataSet = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim dtOK As DataTable = Nothing
        Dim dtError As DataTable = Nothing
        Dim dtImportErr As DataTable = Nothing
        Dim aErrMsg As String = Nothing
        Dim foldName As String = "xls"
        Dim extensionName As String = System.IO.Path.GetExtension("C:\" & FileName)
        If extensionName.ToUpper = ".xlsx".ToUpper Then
            foldName = "log"
        End If
        Dim CurrentDir As String = String.Format("{0}\{1}\", CableSoft.BLL.Utility.Utility.GetCurrentDirectory, foldName)
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

        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId)
        dsReturn = CreateResultDataset(Integer.Parse(CitemPara.Rows(0).Item("EnterType")))
        dtOK = dsReturn.Tables(tbOKName)
        dtError = dsReturn.Tables(tbErrorName)
        dtImportErr = New DataTable(tbImportErrName)
        dtImportErr.Columns.Add(New DataColumn(ErrorNameFieldName, GetType(String)))

        Try
            Try
                'tbExcel = CableSoft.Utility.Heterogeneous.xls.ToDataTable2(String.Format("{0}\{1}", CurrentDir, FileName), , , , 1)
                tbExcel = CableSoft.Utility.Heterogeneous.xls.ToDataTable(String.Format("{0}\{1}", CurrentDir, FileName), ,  , , 1)
            Catch ex As Exception
                'Throw New Exception("匯入失敗")
            End Try


            For Each rwExcel As DataRow In tbExcel.Rows

                Using dsResult As DataSet = ProcessBillData(Integer.Parse(CitemPara.Rows(0).Item("EnterType")),
                                                                      rwExcel.Item(Language.SNo), CitemPara, UCRefNo)
                    For Each rwOK As DataRow In dsResult.Tables(tbOKName).Rows
                        Dim rwNew As DataRow = dtOK.NewRow
                        rwNew.ItemArray = rwOK.ItemArray
                        dtOK.Rows.Add(rwNew)
                        dtOK.AcceptChanges()
                    Next
                    For Each rwError As DataRow In dsResult.Tables(tbErrorName).Rows
                        If String.IsNullOrEmpty(aErrMsg) Then
                            aErrMsg = String.Format("( {0} ) {1}",
                                                        rwExcel.Item(Language.SNo),
                                                        rwError.Item(ErrorNameFieldName))
                        Else
                            aErrMsg = aErrMsg & Environment.NewLine &
                                                    String.Format("( {0} ) {1}",
                                                                            rwExcel.Item(Language.SNo),
                                                                            rwError.Item(ErrorNameFieldName))
                        End If
                        Dim rwNew As DataRow = dtError.NewRow
                        rwNew.ItemArray = rwError.ItemArray
                        rwNew.Item(ErrorCodeFieldName) = rwError.Item(ErrorCodeFieldName)
                        rwNew.Item(ErrorNameFieldName) = rwError.Item(ErrorNameFieldName)
                        dtError.Rows.Add(rwNew)
                        dtError.AcceptChanges()
                    Next
                End Using
            Next
            If Not String.IsNullOrEmpty(aErrMsg) Then
                Dim rwImport As DataRow = dtImportErr.NewRow
                rwImport.Item(ErrorNameFieldName) = aErrMsg
                dtImportErr.Rows.Add(rwImport)
                dtImportErr.AcceptChanges()
            End If
            dsReturn.Tables.Add(dtImportErr)
            Using dsInfo As DataSet = GetTempAllInfo(Integer.Parse(CitemPara.Rows(0).Item("EnterType")))
                dsReturn.Tables.Add(dsInfo.Tables(tbInfoName).Copy)
                dsReturn.Tables.Add(dsInfo.Tables(tbTemp).Copy)
            End Using
            trans.Commit()
        Catch ex As Exception
            If trans IsNot Nothing Then
                trans.Rollback()
            End If
            Throw
        Finally
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If
            If tbExcel IsNot Nothing Then
                tbExcel.Dispose()
                tbExcel = Nothing
            End If

            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                    CSLog = Nothing
                End If
            End If
        End Try

        Return dsReturn
    End Function

    Public Function GetAllData() As DataSet
        Using dsReturn As New DataSet
            Dim dtCompCode As DataTable = GetCompCode.Copy
            dtCompCode.TableName = "CompCode"
            dsReturn.Tables.Add(dtCompCode)
            Dim dtCMCode As DataTable = GetCMCode.Copy
            dtCMCode.TableName = "CMCode"
            dsReturn.Tables.Add(dtCMCode)
            Dim dtPTCode As DataTable = GetPTCode.Copy
            dtPTCode.TableName = "PTCode"
            dsReturn.Tables.Add(dtPTCode)
            Dim dtClctEn As DataTable = GetClctEn.Copy
            dtClctEn.TableName = "ClctEn"
            dsReturn.Tables.Add(dtClctEn)
            Dim dtSTCode As DataTable = GetSTCode.Copy
            dtSTCode.TableName = "STCode"
            dsReturn.Tables.Add(dtSTCode)
            Dim dtPara As DataTable = GetParameters.Copy
            dtPara.TableName = "Para"
            dsReturn.Tables.Add(dtPara)

            Return dsReturn
        End Using
    End Function
    Public Function GetCompCode() As DataTable
        If Me.LoginInfo.GroupId = "0" AndAlso 1 = 0 Then
            Return DAO.ExecQry(_DAL.GetCompCode("0"))
        Else
            Return DAO.ExecQry(_DAL.GetCompCode(Me.LoginInfo.GroupId), New Object() {Me.LoginInfo.EntryId})
        End If

    End Function
    Public Function GetTempInfo(ByVal EntryType As Integer) As DataTable
        'Return DAO.ExecQry(_DAL.GetTempInfo(EntryType), New Object() {Me.LoginInfo.EntryName})
        Return DAO.ExecQry(_DAL.GetTempInfo(EntryType), New Object() {Me.LoginInfo.EntryId})
    End Function
    Public Function GetTempAllInfo(ByVal EntryType As Integer) As DataSet
        Dim dsReturn As New DataSet
        Try
            Dim dtInfo As DataTable = GetTempInfo(EntryType).Copy
            'Dim dtAllTemp As DataTable = DAO.ExecQry(_DAL.QueryEnterData(EntryType), New Object() {Me.LoginInfo.EntryName}).Copy
            Dim dtAllTemp As DataTable = DAO.ExecQry(_DAL.QueryEnterData(EntryType), New Object() {Me.LoginInfo.EntryId}).Copy
            dtInfo.TableName = tbInfoName
            dtAllTemp.TableName = tbTemp
            dsReturn.Tables.Add(dtInfo)
            dsReturn.Tables.Add(dtAllTemp)
        Catch ex As Exception
            Throw
        End Try
        Return dsReturn
    End Function
    Public Function GetCMCode() As DataTable
        Return DAO.ExecQry(_DAL.GetCMCode)
    End Function
    Public Function GetPTCode() As DataTable
        Return DAO.ExecQry(_DAL.GetPTCode)
    End Function
    Public Function GetClctEn() As DataTable
        Return DAO.ExecQry(_DAL.GetClctEn)
    End Function
    Public Function GetSTCode() As DataTable
        Return DAO.ExecQry(_DAL.GetSTCode)
    End Function
    Public Function QueryCancelReason() As DataSet
        Dim dsReturn As New DataSet
        Dim tbCancel As DataTable = DAO.ExecQry(_DAL.QueryCancelReason, New Object() {Me.LoginInfo.CompCode}).Copy
        tbCancel.TableName = "CancelReason"
        dsReturn.Tables.Add(tbCancel)
        Return dsReturn
    End Function
    Public Function GetParameters() As DataTable
        Return DAO.ExecQry(_DAL.GetParameters)
    End Function
    Private Function ChkManualNoOk(ByVal aManualNo As String,
                                    ByVal aBillNo As String,
                                    ByVal aServiceType As String,
                                    ByRef RetMsg As String) As Boolean
        'Dim BillUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim tbSystem As DataTable = Nothing
        Try
            tbSystem = DAO.ExecQry(_DAL.GetUseManual, New Object() {aServiceType})
            If tbSystem.Rows.Count > 0 Then
                If Integer.Parse("0" & tbSystem.Rows(0).Item("UseManual")) = 1 Then
                    Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetUseManualStatus, New Object() {aManualNo, Me.LoginInfo.CompCode})
                        If Not dr.HasRows Then
                            RetMsg = Language.NotFoundNo
                            Return False
                        Else
                            dr.Read()
                            If Integer.Parse(dr.Item(0)) = 0 Then
                                RetMsg = Language.HaveCancelNo
                                Return False
                            End If
                            If DAO.ExecSclr(_DAL.IsUseManual, New Object() {aManualNo, Me.LoginInfo.CompCode, aBillNo}) > 0 Then
                                RetMsg = Language.HaveBillNo
                                Return False
                            End If
                        End If
                    End Using
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            'BillUtility.Dispose()
            tbSystem.Dispose()
        End Try
        Return True
    End Function
    Public Function ChkCloseDate(ByVal CloseDate As String, ByVal ServiceType As String) As RIAResult
        Dim result As New RIAResult
        Dim DayCut As Integer = 0
        Dim Para6 As Integer = 0
        Dim TranDate As String = New Date(1911, 1, 1).ToString
        Try
            If (Not String.IsNullOrEmpty(CloseDate)) AndAlso (Not IsDate(CloseDate)) Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                '                result.ErrorMessage = "日期不合法！"
                result.ErrorMessage = Language.DateIsIllegal
                Return result
            End If
            If Date.Parse(CloseDate) > Date.Now Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                '                result.ErrorMessage = "此日期超過今天日期！"
                result.ErrorMessage = Language.OverToday
                Return result
            End If
            DayCut = Integer.Parse(DAO.ExecSclr(_DAL.GetDayCut, New Object() {LoginInfo.CompCode}))
            Using dtTranDate As DataTable = DAO.ExecQry(_DAL.GetTranDate(ServiceType), New Object() {LoginInfo.CompCode})
                If dtTranDate.Rows.Count > 0 Then
                    For Each rw As DataRow In dtTranDate.Rows
                        If IsDate(rw.Item(0)) Then
                            TranDate = rw.Item(0)
                            Exit For
                        End If
                    Next
                End If
                dtTranDate.Dispose()
            End Using
            Para6 = Integer.Parse(DAO.ExecSclr(_DAL.GetPara6(ServiceType), New Object() {LoginInfo.CompCode}))
            Dim diffday As TimeSpan = Date.Now.Subtract(Date.Parse(CloseDate))
            If diffday.Days > Para6 Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                'result.ErrorMessage = "此日期已超過系統設定的安全期限！"
                result.ErrorMessage = Language.OverSafeDay
                Return result
            End If
            If Date.Parse(CloseDate) <= TranDate Then
                If (DayCut = 1) AndAlso (Date.Parse(CloseDate) = TranDate) Then
                Else
                    result.ResultBoolean = False
                    result.ErrorCode = -1
                    'result.ErrorMessage = "之前已做過日結，不可使用之前日期入帳"
                    result.ErrorMessage = Language.HadClosed
                    Return result
                End If
            End If
            result.ResultBoolean = True
            result.ErrorMessage = Nothing
            result.ErrorCode = 0
        Catch ex As Exception
            result.ErrorCode = -1
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
        End Try
        Return result
    End Function
    Public Function ChkDataOK(ByVal BillNo As String) As Boolean
        'Dim dtSO033 As DataTable = Nothing
        Try
            Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetSO033Data(BillNo.Length),
                                                     New Object() {Me.LoginInfo.CompCode, BillNo})


            End Using
        Catch ex As Exception
            Throw ex
        Finally

        End Try
        Return True
    End Function
    Public Function CancelTempData(ByVal EntryType As Integer,
                                   ByVal BillNo As String,
                                   ByVal Item As String,
                                   ByVal CancelDate As String,
                                   ByVal CancelCode As Integer, ByVal CancelName As String) As Boolean
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim dsReturn As New DataSet
        Dim dtOK As DataTable = Nothing
        Dim dtError As DataTable = Nothing
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
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId)
        Try

            DAO.ExecNqry(_DAL.CancelTempData(EntryType), New Object() {
                         DateTime.Parse(CancelDate),
                         CancelCode,
                         CancelName,
                         BillNo, Integer.Parse(Item)})
            trans.Commit()
        Catch ex As Exception
            If trans IsNot Nothing Then
                trans.Rollback()
            End If
            Throw
        Finally

            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                    CSLog = Nothing
                End If
            End If
        End Try
        Return True
    End Function
    Public Function ChkCanCancel(ByVal CitemCode As String) As String
        If Integer.Parse(DAO.ExecSclr(_DAL.GetSecondDiscount, New Object() {Me.LoginInfo.CompCode})) = 0 Then
            Return Nothing
        End If
        Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.ChkHaveDiscount, New Object() {CitemCode, CitemCode})
            Do While dr.Read
                If Integer.Parse(dr.Item("SecendDiscount")) > 0 Then
                    Return Language.HaveDiscount
                End If
                If Integer.Parse(dr.Item("RefNo")) = 19 Then
                    Return Language.HaveDiscount
                End If
            Loop
        End Using
        Return Nothing
    End Function

    Public Function Execute(ByVal EntryType As Integer, ByVal CitemPara As DataTable) As DataSet
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim dsReturn As New DataSet
        Dim dtOK As DataTable = Nothing
        Dim dtError As DataTable = Nothing

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
        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans

        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId, Language.ClientInfoString)
        Dim objBillUtility As New CableSoft.SO.BLL.Billing.Utility.Utility(Me.LoginInfo, DAO)
        Try
            FNowDate = DateTime.Now
            'Using dtTempData As DataTable = DAO.ExecQry(_DAL.GetTempData(EntryType), New Object() {Me.LoginInfo.EntryName})
            Using dtTempData As DataTable = DAO.ExecQry(_DAL.GetTempData(EntryType), New Object() {Me.LoginInfo.EntryId})
                dsReturn = CreateResultDataset(EntryType)
                dtOK = dsReturn.Tables(tbOKName)
                dtError = dsReturn.Tables(tbErrorName)
                'dtOK = dtTempData.Clone
                'dtOK.Clear()
                'dtOK.TableName = tbOKName
                'dtError = dtOK.Clone
                'dtError.TableName = tbErrorName
                'dtError.Columns.Add(New DataColumn("ErrMsg", GetType(String)))
                For Each rwTemp As DataRow In dtTempData.Rows
                    Using tbSO033 As DataTable = DAO.ExecQry(_DAL.GetChargeData(), New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})
                        For Each rwSO033 As DataRow In tbSO033.Rows
                            If (DBNull.Value.Equals(rwSO033.Item("UCCode"))) OrElse (String.IsNullOrEmpty(rwSO033.Item("UCCode").ToString)) Then
                                Dim rwErr As DataRow = dtError.NewRow
                                For Each col As DataColumn In dtTempData.Columns
                                    rwErr.Item(col.ColumnName) = rwTemp.Item(col.ColumnName)
                                Next
                                rwErr.Item(ErrorCodeFieldName) = "-1"
                                rwErr.Item(ErrorNameFieldName) = Language.HasBillNo
                                dtError.Rows.Add(rwErr)
                                Continue For
                            End If
                            If (Not DBNull.Value.Equals(rwSO033.Item("CancelFlag"))) AndAlso
                                    (Integer.Parse(rwSO033.Item("CancelFlag").ToString) = 1) Then
                                Dim rwErr As DataRow = dtError.NewRow
                                For Each col As DataColumn In dtTempData.Columns
                                    rwErr.Item(col.ColumnName) = rwTemp.Item(col.ColumnName)
                                Next
                                rwErr.Item(ErrorCodeFieldName) = "-1"
                                rwErr.Item(ErrorNameFieldName) = Language.HasCancel
                                dtError.Rows.Add(rwErr)
                                Continue For
                            End If

                            DAO.ExecNqry(_DAL.UpdateChargeData, GetUpdateValuePara(rwSO033, rwTemp, EntryType, CitemPara))
                            'fill out fields of so127 as the field of manualno is not empty by kin 2017/05/12
                            If Not DBNull.Value.Equals(rwTemp("ManualNo")) Then
                                Dim Tel1 As Object = DAO.ExecSclr(_DAL.getTel1, New Object() {rwTemp("CustId")})
                                DAO.ExecNqry(_DAL.UpdateSO127, New Object() {rwTemp("CustId"),
                                                                             rwTemp("CustName"), Tel1, rwTemp("BillNo"),
                                                                             rwTemp("RealDate"), CableSoft.BLL.Utility.DateTimeUtility.GetDTString(Now),
                                                                             rwTemp("ManualNo")})
                            End If
                            DAO.ExecNqry(_DAL.DelTempData(EntryType), New Object() {rwTemp.Item("BillNo"),
                                                                                    rwTemp.Item("Item")})
                            Dim tbUpdSO033 As DataTable = DAO.ExecQry(_DAL.GetChargeData(), New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})

                            If Not objBillUtility.UpdatePeriodCycle(tbUpdSO033.Rows(0), False) Then
                                Dim rwErr As DataRow = dtError.NewRow
                                For Each col As DataColumn In dtTempData.Columns
                                    rwErr.Item(col.ColumnName) = rwTemp.Item(col.ColumnName)
                                Next
                                rwErr.Item(ErrorCodeFieldName) = "-1"
                                rwErr.Item(ErrorNameFieldName) = Language.UpdBillError
                                dtError.Rows.Add(rwErr)
                                Continue For
                            End If
                            Dim rwOK As DataRow = dtOK.NewRow
                            rwOK.ItemArray = rwTemp.ItemArray
                            dtOK.Rows.Add(rwOK)




                        Next
                    End Using
                Next
            End Using
            trans.Commit()

            'If dtOK IsNot Nothing Then
            '    dsReturn.Tables.Add(dtOK)
            'End If
            'If dtError IsNot Nothing Then
            '    dsReturn.Tables.Add(dtError)
            'End If
            Return dsReturn
        Catch ex As Exception
            If trans IsNot Nothing Then
                trans.Rollback()
            End If
            Throw ex
        Finally
            If cmd IsNot Nothing Then
                cmd.Dispose()
            End If
            If objBillUtility IsNot Nothing Then
                objBillUtility.Dispose()
                objBillUtility = Nothing
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

    End Function
    Private Function CreateResultDataset(ByVal EntryType As Integer) As DataSet
        Dim dsReturn As New DataSet
        Dim dtOK As DataTable = Nothing
        Dim dtError As DataTable = Nothing
        Try
            dtOK = DAO.ExecQry(_DAL.GetTempOK(EntryType), New Object() {"X", 1}).Clone
            dtOK.TableName = tbOKName
            dtError = dtOK.Clone
            dtError.TableName = tbErrorName
            dtOK.Columns.Add(New DataColumn(WaringMsgFieldName, GetType(String)))
            dtError.Columns.Add(New DataColumn(ErrorCodeFieldName, GetType(String)))
            dtError.Columns.Add(New DataColumn(ErrorNameFieldName, GetType(String)))
            dsReturn.Tables.Add(dtError)
            dsReturn.Tables.Add(dtOK)
        Catch ex As Exception
            Throw
        Finally

        End Try
        Return dsReturn
    End Function
    Private Function ProcessBillData(ByVal EntryType As Integer,
                                   ByVal BillNo As String, ByVal CitemPara As DataTable,
                                   ByVal UCRefNo As Integer) As DataSet
        Dim dtSO033 As DataTable = Nothing
        Dim dtOK As DataTable = Nothing
        Dim dtError As DataTable = Nothing
        Dim dsReturn As DataSet
        Dim RealAmt As Integer = 0
        Dim objBillUtility As New CableSoft.SO.BLL.Billing.Utility.Utility(Me.LoginInfo, DAO)
        Dim cloneCitemPara As DataTable = CitemPara.Copy()
        Try
            dsReturn = CreateResultDataset(EntryType)
            dtOK = dsReturn.Tables(tbOKName)
            dtError = dsReturn.Tables(tbErrorName)
            'EntryNoCount = Integer.Parse(DAO.ExecSclr(_DAL.GetEntryNoCount(EntryType), New Object() {Me.LoginInfo.EntryName}))
            EntryNoCount = Integer.Parse(DAO.ExecSclr(_DAL.GetEntryNoCount(EntryType), New Object() {Me.LoginInfo.EntryId}))
            'dtOK = DAO.ExecQry(_DAL.GetTempOK(EntryType), New Object() {"X", 1}).Copy
            'dtOK.TableName = "OK"
            'dtError = dtOK.Copy
            'dtError.TableName = "Error"
            'dtError.Columns.Add(New DataColumn("ErrorCode", GetType(String)))
            'dtError.Columns.Add(New DataColumn("ErrorName", GetType(String)))
            'dsReturn.Tables.Add(dtError)
            'dsReturn.Tables.Add(dtOK)
            If Integer.Parse(DAO.ExecSclr(_DAL.ChkDupData(EntryType, BillNo.Length), New Object() {BillNo})) > 0 Then
                Dim rw As DataRow = dtError.NewRow
                rw.Item(ErrorCodeFieldName) = "-1"
                rw.Item(ErrorNameFieldName) = Language.HasInput
                dtError.Rows.Add(rw)
                dtError.AcceptChanges()
                Return dsReturn
            End If
            If Integer.Parse(DAO.ExecSclr(_DAL.chkREF3(BillNo.Length), New Object() {BillNo, Me.LoginInfo.CompCode})) > 0 Then
                Dim rw As DataRow = dtError.NewRow
                rw.Item(ErrorCodeFieldName) = "-1"
                rw.Item(ErrorNameFieldName) = Language.PayForCounter
                dtError.Rows.Add(rw)
                dtError.AcceptChanges()
                Return dsReturn
            End If
            If BillNo.Length = 11 OrElse BillNo.Length = 12 OrElse BillNo.Length = 15 Then
                dtSO033 = DAO.ExecQry(_DAL.GetSO033Data(BillNo.Length), New Object() {Me.LoginInfo.CompCode, BillNo})
                dtSO033.TableName = "Charge"
                'dtSO033.Columns.Add("WaringMsg", GetType(String))
                'dtSO033.Columns.Add("BillCount", GetType(Integer))
                If dtSO033.Rows.Count = 0 Then
                    Dim rw As DataRow = dtError.NewRow
                    rw.Item(ErrorCodeFieldName) = "-1"
                    rw.Item(ErrorNameFieldName) = Language.NotFoundBillNo
                    dtError.Rows.Add(rw)
                    dtError.AcceptChanges()
                    Return dsReturn

                End If
               
                For Each rw033 As DataRow In dtSO033.Rows

                    If (Not DBNull.Value.Equals(cloneCitemPara.Rows(0).Item("RealAmt"))) AndAlso
                                (Not String.IsNullOrEmpty(cloneCitemPara.Rows(0).Item("RealAmt"))) Then
                        RealAmt = cloneCitemPara.Rows(0).Item("RealAmt")
                        '畫面條件大於0而且SO033筆數大於1則以SO033的應收金額為主 For Debby by Kin 2017/05/11
                        If dtSO033.Rows.Count > 1 AndAlso RealAmt > 0 Then
                            RealAmt = rw033("ShouldAmt")
                            cloneCitemPara.Rows(0).Item("RealAmt") = DBNull.Value
                        End If
                    Else
                        RealAmt = rw033("ShouldAmt")
                    End If
                    If (rw033("ShouldAmt") <> RealAmt) Then
                        If (DBNull.Value.Equals(cloneCitemPara.Rows(0).Item("STCode"))) OrElse
                            (String.IsNullOrEmpty(cloneCitemPara.Rows(0).Item("STCode").ToString)) Then
                            Dim rw As DataRow = dtError.NewRow
                            rw.Item(ErrorCodeFieldName) = "-1"
                            rw.Item(ErrorNameFieldName) = Language.MustSTReason
                            dtError.Rows.Add(rw)
                            dtError.AcceptChanges()
                            Return dsReturn
                        End If
                    End If
                Next

                Dim ErrMsg As String = Nothing
                Dim UCCode As String = Nothing
                Dim UCName As String = Nothing
                Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefaultUCCode, New Object() {UCRefNo})
                    While dr.Read
                        UCCode = dr.Item("CodeNo").ToString
                        UCName = dr.Item("Description").ToString
                    End While
                End Using



                For Each rw033 As DataRow In dtSO033.Rows
                    If Not DBNull.Value.Equals(cloneCitemPara.Rows(0).Item("ManualNo")) Then

                        If Not ChkManualNoOk(cloneCitemPara.Rows(0).Item("ManualNo"), rw033.Item("BillNo"),
                                         rw033.Item("ServiceType"), ErrMsg) Then
                            Dim rwErr As DataRow = dtError.NewRow
                            rwErr.Item(ErrorCodeFieldName) = "-1"
                            rwErr.Item(ErrorNameFieldName) = ErrMsg
                            dtError.Rows.Add(rwErr)
                            dtError.AcceptChanges()
                            Return dsReturn
                        End If
                    End If

                    ErrMsg = Nothing
                    Dim aryValue() As Object = GetValuePara(rw033, cloneCitemPara, EntryType, ErrMsg)


                    If Not String.IsNullOrEmpty(ErrMsg) Then
                        Dim rwErr As DataRow = dtError.NewRow
                        rwErr.Item(ErrorCodeFieldName) = "-1"
                        rwErr.Item(ErrorNameFieldName) = ErrMsg
                        dtError.Rows.Add(rwErr)
                        dtError.AcceptChanges()
                        Return dsReturn
                    End If
                    DAO.ExecNqry(_DAL.InsertTmpData(EntryType), aryValue)
                    DAO.ExecNqry(_DAL.UpdDefUCCodeCharge, New Object() {Integer.Parse(UCCode), UCName, rw033.Item("BillNo"), rw033.Item("Item")})
                    Dim tbUpdSO033 As DataTable = DAO.ExecQry(_DAL.GetChargeData(), New Object() {rw033.Item("BillNo"), rw033.Item("Item")})
                    objBillUtility.UpdateVODPoint(rw033)
                    Dim tbTempOK As DataTable = DAO.ExecQry(_DAL.GetTempOK(EntryType), _
                                                            New Object() {rw033.Item("BillNo"), rw033.Item("Item")})
                    If tbTempOK.Rows.Count > 0 Then
                        Dim rwOK As DataRow = dtOK.NewRow
                        rwOK.ItemArray = tbTempOK.Rows(0).ItemArray
                        If Integer.Parse(rw033.Item("CustStatusCode")) <> 1 Then
                            rwOK.Item(WaringMsgFieldName) = Language.CustidNotNormal
                        End If
                        dtOK.Rows.Add(rwOK)
                        dtOK.AcceptChanges()
                        EntryNoCount += 1
                    End If
                Next
            Else
                If dtSO033 Is Nothing Then
                    Dim rwErr As DataRow = dtError.NewRow
                    rwErr.Item(ErrorCodeFieldName) = "-1"
                    rwErr.Item(ErrorNameFieldName) = Language.SNoLenError
                    dtError.Rows.Add(rwErr)
                    dtError.AcceptChanges()
                    Return dsReturn
                End If
            End If
            Return dsReturn
        Catch ex As Exception
            Throw
        Finally
            If cloneCitemPara IsNot Nothing Then
                cloneCitemPara.Dispose()
                cloneCitemPara = Nothing
            End If
            objBillUtility.Dispose()
            objBillUtility = Nothing
        End Try
    End Function
    Public Function EntryBillData(ByVal EntryType As Integer,
                                  ByVal BillNo As String, ByVal CitemPara As DataTable,
                                  ByVal UCRefNo As Integer) As DataSet
        'Dim dtSO033 As DataTable = Nothing
        'Dim dtOK As DataTable = Nothing
        'Dim dtError As DataTable = Nothing
        Dim dsReturn As DataSet = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        'Dim objBillUtility As New CableSoft.SO.BLL.Billing.Utility.Utility(Me.LoginInfo, DAO)

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

        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId)
        Try
            dsReturn = ProcessBillData(EntryType, BillNo, CitemPara, UCRefNo)
            Using dsInfo As DataSet = GetTempAllInfo(EntryType)
                dsReturn.Tables.Add(dsInfo.Tables("Info").Copy)
                dsReturn.Tables.Add(dsInfo.Tables("Temp").Copy)
            End Using
            'If BillNo.Length = 11 OrElse BillNo.Length = 12 OrElse BillNo.Length = 15 Then
            '    dtSO033 = DAO.ExecQry(_DAL.GetSO033Data(BillNo.Length), New Object() {Me.LoginInfo.CompCode, BillNo})
            '    dtSO033.TableName = "Charge"
            '    dtSO033.Columns.Add("WaringMsg", GetType(String))
            '    dtSO033.Columns.Add("BillCount", GetType(Integer))
            '    If dtSO033.Rows.Count = 0 Then
            '        Dim rw As DataRow = dtError.NewRow
            '        rw.Item("ErrorCode") = "-1"
            '        rw.Item("ErrorName") = "無此單據編號或此單據已收款!!"
            '        dtError.Rows.Add(rw)
            '        dtError.AcceptChanges()
            '        Return dsReturn

            '    End If
            '    If Integer.Parse(DAO.ExecSclr(_DAL.ChkDupData(EntryType, BillNo.Length), New Object() {BillNo})) > 0 Then
            '        Dim rw As DataRow = dtError.NewRow
            '        rw.Item("ErrorCode") = "-1"
            '        rw.Item("ErrorName") = "此單據已登錄過!"
            '        dtError.Rows.Add(rw)
            '        dtError.AcceptChanges()
            '        Return dsReturn

            '    End If
            '    Dim ErrMsg As String = Nothing
            '    Dim UCCode As String = Nothing
            '    Dim UCName As String = Nothing
            '    Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefaultUCCode, New Object() {UCRefNo})
            '        While dr.Read
            '            UCCode = dr.Item("CodeNo").ToString
            '            UCName = dr.Item("Description").ToString
            '        End While
            '    End Using
            '    For Each rw As DataRow In dtSO033.Rows
            '        If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("ManualNo")) Then

            '            If Not ChkManualNoOk(CitemPara.Rows(0).Item("ManualNo"), rw.Item("BillNo"),
            '                             rw.Item("ServiceType"), ErrMsg) Then
            '                Dim rwErr As DataRow = dtError.NewRow
            '                rwErr.Item("ErrorCode") = "-1"
            '                rwErr.Item("ErrorName") = ErrMsg
            '                dtError.Rows.Add(rwErr)
            '                dtError.AcceptChanges()
            '                Return dsReturn
            '            End If
            '        End If
            '        ErrMsg = Nothing
            '        Dim aryValue() As Object = GetValuePara(rw, CitemPara, EntryType, ErrMsg)

            '        If Not String.IsNullOrEmpty(ErrMsg) Then
            '            Dim rwErr As DataRow = dtError.NewRow
            '            rwErr.Item("ErrorCode") = "-1"
            '            rwErr.Item("ErrorName") = ErrMsg
            '            dtError.Rows.Add(rwErr)
            '            dtError.AcceptChanges()
            '            Return dsReturn
            '        End If
            '        DAO.ExecNqry(_DAL.InsertTmpData(EntryType), aryValue)
            '        DAO.ExecNqry(_DAL.UpdDefUCCodeCharge, New Object() {UCCode, UCName, rw.Item("BillNo"), rw.Item("Item")})
            '        objBillUtility.UpdateVODPoint(rw)
            '        Dim tbTempOK As DataTable = DAO.ExecQry(_DAL.GetTempOK(EntryType), _
            '                                                New Object() {rw.Item("BillNo"), rw.Item("Item")})
            '        If tbTempOK.Rows.Count > 0 Then
            '            Dim rwOK As DataRow = dtOK.NewRow
            '            rwOK.ItemArray = tbTempOK.Rows(0).ItemArray
            '            dtOK.Rows.Add(rwOK)
            '            dtOK.AcceptChanges()
            '        End If
            '    Next


            'Else
            '    If dtSO033 Is Nothing Then
            '        Dim rwErr As DataRow = dtError.NewRow
            '        rwErr.Item("ErrorCode") = "-1"
            '        rwErr.Item("ErrorName") = "單據長度錯誤!!"
            '        dtError.Rows.Add(rwErr)
            '        dtError.AcceptChanges()
            '        Return dsReturn
            '    End If
            'End If

            'trans.Commit()
            'Return dsReturn
            trans.Commit()
        Catch ex As Exception
            Throw ex
        Finally

            'objBillUtility.Dispose()
            'dtError.Dispose()
            'dsReturn.Dispose()
            If cmd IsNot Nothing Then
                cmd.Dispose()
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
        Return dsReturn
    End Function
    Private Function GetUpdateValuePara(ByVal rw033 As DataRow, ByVal rwTemp As DataRow, ByVal EntryType As Integer, ByVal CitemPara As DataTable) As Object()
        Dim lstPara As New List(Of Object)
        Dim aFieldName As String = "CitemCode,CitemName,SHOULDDATE, " & _
                                           " REALDATE,SHOULDAMT,REALAMT, " & _
                        "REALPERIOD,REALSTARTDATE,REALSTOPDATE,CLCTEN," & _
                        "CLCTNAME,PTCODE,PTNAME,UPDTIME,NEWUPDTIME,UPDEN,CMCODE,CMNAME," & _
                       "MANUALNO,UCCode,UCName,STCODE,STNAME,Note,ServiceType," & _
                       "CancelFlag,CancelCode,CancelName," & _
                       "BankCode,BankName,AccountNo,AuthorizeNo,AdjustFlag" & _
                       "NextPeriod,NextAmt,InvSeqNo,FirstTime,ClctYM"
        lstPara.Add(rwTemp.Item("CitemCode"))
        lstPara.Add(rwTemp.Item("CitemName"))
        lstPara.Add(rwTemp.Item("SHOULDDATE"))
        lstPara.Add(rwTemp.Item("REALDATE"))
        lstPara.Add(rwTemp.Item("SHOULDAMT"))
        lstPara.Add(rwTemp.Item("REALAMT"))
        lstPara.Add(rwTemp.Item("REALPERIOD"))
        lstPara.Add(rwTemp.Item("REALSTARTDATE"))
        lstPara.Add(rwTemp.Item("REALSTOPDATE"))
        lstPara.Add(rwTemp.Item("CLCTEN"))
        lstPara.Add(rwTemp.Item("CLCTNAME"))
        lstPara.Add(rwTemp.Item("PTCODE"))
        lstPara.Add(rwTemp.Item("PTNAME"))
        lstPara.Add(CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate))
        lstPara.Add(FNowDate)
        'lstPara.Add(rwTemp.Item("EntryEn"))
        lstPara.Add(LoginInfo.EntryName)
        lstPara.Add(rwTemp.Item("CMCODE"))
        lstPara.Add(rwTemp.Item("CMNAME"))
        lstPara.Add(rwTemp.Item("MANUALNO"))
        lstPara.Add(DBNull.Value)
        lstPara.Add(DBNull.Value)
        lstPara.Add(rwTemp.Item("STCODE"))
        lstPara.Add(rwTemp.Item("STNAME"))
        lstPara.Add(rwTemp.Item("Note"))
        lstPara.Add(rwTemp.Item("ServiceType"))
        lstPara.Add(rwTemp.Item("CancelFlag"))
        lstPara.Add(rwTemp.Item("CancelCode"))
        lstPara.Add(rwTemp.Item("CancelName"))
        If EntryType = 1 Then
            lstPara.Add(rw033.Item("BankCode"))
            lstPara.Add(rw033.Item("BankName"))
        Else
            lstPara.Add(rwTemp.Item("BankCode"))
            lstPara.Add(rwTemp.Item("BankName"))
        End If
        If EntryType = 1 Then
            lstPara.Add(rw033.Item("AccountNo"))
        Else
            lstPara.Add(rwTemp.Item("AccountNo"))
        End If
        If EntryType = 1 Then
            lstPara.Add(rw033.Item("AuthorizeNo"))
        Else
            lstPara.Add(rwTemp.Item("AuthorizeNo"))
        End If
        If EntryType = 1 Then
            lstPara.Add(rw033.Item("AdjustFlag"))
        Else
            lstPara.Add(rwTemp.Item("AdjustFlag"))
        End If
        If EntryType = 1 Then
            lstPara.Add(rw033.Item("NextPeriod"))
        Else
            lstPara.Add(rwTemp.Item("NextPeriod"))
        End If
        If EntryType = 1 Then
            lstPara.Add(rw033.Item("NextAmt"))
        Else
            lstPara.Add(rwTemp.Item("NextAmt"))
        End If
        If EntryType = 1 Then
            lstPara.Add(rw033.Item("InvSeqNo"))
        Else
            lstPara.Add(rwTemp.Item("InvSeqNo"))
        End If
        If DBNull.Value.Equals(rw033("RealDate")) AndAlso (Not DBNull.Value.Equals(rwTemp.Item("RealDate"))) AndAlso
                (DBNull.Value.Equals(rw033.Item("FirstTime"))) Then
            lstPara.Add(CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate))
        Else
            lstPara.Add(rw033.Item("FirstTime"))
        End If
        If CitemPara IsNot Nothing AndAlso CitemPara.Rows.Count > 0 Then
            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("ClctYM")) Then
                lstPara.Add(CitemPara.Rows(0).Item("ClctYM"))
            Else
                lstPara.Add(DBNull.Value)
            End If
        Else
            lstPara.Add(DBNull.Value)
        End If
        lstPara.Add(rwTemp.Item("BillNo"))
        lstPara.Add(rwTemp.Item("Item"))

        Return lstPara.ToArray
    End Function

    Private Function GetValuePara(ByVal rw As DataRow,
                                   ByVal CitemPara As DataTable,
                                   ByVal EnterType As Integer,
                                  ByRef ErrMsg As String) As Object()
        Dim lstPara As New List(Of Object)
        Dim RealAmt As Integer = 0
        'aFieldName = "BillNo,Item,Custid,CustName,CitemCode,CitemName, " & _
        '                   "MediaBillNo,PrtSNo,ShouldAmt,ShouldDate," & _
        '                   "RealAmt,ManualNo,RealDate,RealPeriod,RealStartDate,RealStopDate," & _
        '                   "EntryEn,Note,CMCode,CMName,ClctEn,ClctName," & _
        '                   "PTCode,PTName,RcdRowId,EntryNO,StCode,STName," & _
        '                   "ServiceType,CompCode,ManualNo,CancelFlag,CancelCode," & _
        '                   "CancelName,BankCode,BankName,AccountNo,AuthorizeNo," & _
        '                   "AdjustFlag,NextPeriod,NextAmt,FaciSeqNo,InvSeqNo"
        Try
            lstPara.Add(rw.Item("BillNo"))
            lstPara.Add(rw.Item("Item"))
            lstPara.Add(rw.Item("Custid"))
            lstPara.Add(rw.Item("CustName"))
            lstPara.Add(rw.Item("CitemCode"))
            lstPara.Add(rw.Item("CitemName"))
            lstPara.Add(rw.Item("MediaBillNo"))
            lstPara.Add(rw.Item("PrtSNo"))
            lstPara.Add(rw.Item("ShouldAmt"))
            lstPara.Add(rw.Item("ShouldDate"))
            If (Not DBNull.Value.Equals(CitemPara.Rows(0).Item("RealAmt"))) AndAlso
                (Not String.IsNullOrEmpty(CitemPara.Rows(0).Item("RealAmt"))) Then
                lstPara.Add(CitemPara.Rows(0).Item("RealAmt"))
                RealAmt = Integer.Parse(CitemPara.Rows(0).Item("RealAmt"))
            Else
                lstPara.Add(rw.Item("ShouldAmt"))
                RealAmt = Integer.Parse(rw.Item("ShouldAmt"))
            End If

            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("ManualNo")) Then
                lstPara.Add(CitemPara.Rows(0).Item("ManualNo"))
            Else
                lstPara.Add(rw.Item("ManualNo"))
            End If
            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("RealDate")) Then
                lstPara.Add(CitemPara.Rows(0).Item("RealDate"))
            Else
                lstPara.Add(New Date(Now.Year, Now.Month, Now.Day))
            End If
            Dim Msg As String = New String(" ", 2000)
            Dim intRealAmount As Integer = 0
            Dim intRealPeriod As Integer = 0
            Dim strRealStopDate As String = New String(" ", 2000)
            Dim RetNum As Integer
            Dim period As Integer = 0
            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("Period")) AndAlso Not String.IsNullOrEmpty(CitemPara.Rows(0).Item("Period")) Then
                period = CitemPara.Rows(0).Item("Period")
            Else
                period = Integer.Parse(rw.Item("RealPeriod"))
            End If

            If RealAmt = Integer.Parse(rw.Item("ShouldAmt")) Then
                period = Integer.Parse(rw.Item("RealPeriod"))
            End If
            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("Period")) AndAlso Integer.Parse("0" & CitemPara.Rows(0).Item("Period")) > 0 Then
                If RealAmt = Integer.Parse(rw.Item("ShouldAmt")) Then
                    lstPara.Add(rw.Item("RealPeriod"))
                Else
                    rw("RealPeriod") = CitemPara.Rows(0).Item("Period")
                    lstPara.Add(rw.Item("RealPeriod"))
                End If


                RetNum = SFGetAmount(False, 2, rw("CustId"), rw("CitemCode"),
                                     period, Date.Parse(rw.Item("RealStartDate")).ToString("yyyyMMdd"),
                                     rw("ServiceType"), rw("CompCode"), strRealStopDate, intRealAmount, Msg, intRealPeriod,
                                      New String(" ", 2000), 0, New String(" ", 2000),
                                      New String(" ", 2000), 0, 0, 0,
                                      New String(" ", 2000), New String(" ", 2000), 0)
                If RetNum < 0 Then
                    ErrMsg = Msg
                    Return Nothing
                    'MsgBox(Msg, vbExclamation, "訊息")
                Else
                    If String.IsNullOrEmpty(strRealStopDate) Then
                        RetNum = SFGetAmount(True, 2, rw("CustId"), rw("CitemCode"),
                                     CitemPara.Rows(0).Item("Period"), Date.Parse(rw.Item("RealStartDate")).ToString("yyyyMMdd"),
                                     rw("ServiceType"), rw("CompCode"), strRealStopDate, intRealAmount, Msg, intRealPeriod,
                                      New String(" ", 2000), 0, New String(" ", 2000),
                                      New String(" ", 2000), 0, 0, 0,
                                      New String(" ", 2000), New String(" ", 2000), 0)
                        If RetNum < 0 Then
                            ErrMsg = Msg
                            Return Nothing
                        End If
                    End If
                End If
            Else
                lstPara.Add(rw.Item("RealPeriod"))
            End If
            lstPara.Add(rw.Item("RealStartDate"))
            If Not String.IsNullOrEmpty(strRealStopDate.Replace(" ", "")) Then
                lstPara.Add(New Date(strRealStopDate.Substring(0, 4),
                                     strRealStopDate.Substring(4, 2),
                                     strRealStopDate.Substring(6, 2)))
            Else
                lstPara.Add(rw.Item("RealStopDate"))
                'lstPara.Add(New Date(1911, 1, 1))
            End If

            '            lstPara.Add(Me.LoginInfo.EntryName)
            lstPara.Add(Me.LoginInfo.EntryId)
            lstPara.Add(rw.Item("Note") & CitemPara.Rows(0).Item("Note"))
            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("CMCode")) Then
                lstPara.Add(CitemPara.Rows(0).Item("CMCode"))
                lstPara.Add(CitemPara.Rows(0).Item("CMName"))
            Else
                lstPara.Add(rw.Item("CMCode"))
                lstPara.Add(rw.Item("CMName"))
            End If
            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("ClctEn")) Then
                lstPara.Add(CitemPara.Rows(0).Item("ClctEn"))
                lstPara.Add(CitemPara.Rows(0).Item("ClctName"))
            Else
                If Not DBNull.Value.Equals(rw.Item("ClctEn")) Then
                    lstPara.Add(rw.Item("ClctEn"))
                    lstPara.Add(rw.Item("ClctName"))
                Else
                    lstPara.Add(rw.Item("OldClctEN"))
                    lstPara.Add(rw.Item("OldClctName"))
                End If
            End If
            If Not DBNull.Value.Equals(CitemPara.Rows(0).Item("PTCode")) Then
                lstPara.Add(CitemPara.Rows(0).Item("PTCode"))
                lstPara.Add(CitemPara.Rows(0).Item("PTName"))
            Else
                lstPara.Add(rw.Item("PTCode"))
                lstPara.Add(rw.Item("PTName"))
            End If
            'lstPara.Add(rw.Item("RowId"))
            lstPara.Add(rw.Item("CTID"))
            'lstPara.Add(Integer.Parse("0" & CitemPara.Rows(0).Item("BillCount")) + 1)

            lstPara.Add(EntryNoCount + 1)

            If Integer.Parse(rw.Item("ShouldAmt")) <> RealAmt Then
                lstPara.Add(CitemPara.Rows(0).Item("STCode"))
                lstPara.Add(CitemPara.Rows(0).Item("STName"))
            Else
                lstPara.Add(rw.Item("STCode"))
                lstPara.Add(rw.Item("STName"))

            End If


            lstPara.Add(rw.Item("ServiceType"))
            lstPara.Add(rw.Item("CompCode"))
            lstPara.Add(rw.Item("CancelFlag"))
            lstPara.Add(rw.Item("CancelCode"))
            lstPara.Add(rw.Item("CancelName"))
            lstPara.Add(rw.Item("BankCode"))
            lstPara.Add(rw.Item("BankName"))
            lstPara.Add(rw.Item("AccountNo"))
            lstPara.Add(rw.Item("AuthorizeNo"))
            lstPara.Add(rw.Item("AdjustFlag"))
            lstPara.Add(rw.Item("NextPeriod"))
            lstPara.Add(rw.Item("NextAmt"))
            lstPara.Add(rw.Item("FaciSeqNo"))
            lstPara.Add(rw.Item("InvSeqNo"))
            lstPara.Add(rw.Item("UCCode"))
            lstPara.Add(rw.Item("UCName"))
        Catch ex As Exception
            Throw
        End Try
        Return lstPara.ToArray.Clone
    End Function
    Public Function QueryEnterData(ByVal EntryType As Integer) As DataSet
        Dim dsReturn As New DataSet
        'Dim dtTemp As DataTable = DAO.ExecQry(_DAL.QueryEnterData(EntryType), New Object() {Me.LoginInfo.EntryName}).Copy
        Dim dtTemp As DataTable = DAO.ExecQry(_DAL.QueryEnterData(EntryType), New Object() {Me.LoginInfo.EntryId}).Copy
        dtTemp.TableName = "Temp"
        dsReturn.Tables.Add(dtTemp)
        Return dsReturn
    End Function
    Private Function SFGetAmount(ByVal blnChoose As Boolean, ByVal P_OPTION As Integer, _
                                                            ByVal P_CUSTID As Integer, _
                                                            ByVal P_CITEMCODE As Integer, _
                                                            ByVal P_PERIOD As Integer, _
                                                            ByVal P_REALSTARTDATE As String, _
                                                            ByVal p_SERVICETYPE As String, _
                                                            ByVal p_CompCode As Integer, _
                                                            ByRef P_REALSTOPDATE As String, _
                                                            ByRef P_SHOULDAMT As Integer, _
                                                            ByRef P_RETMSG As String, _
                                                            ByRef p_RealPeriod As Integer, _
                                                            Optional ByVal P_BPCode As String = "", _
                                                            Optional ByRef intPFlag1 As Integer = 0, _
                                                            Optional ByVal P_FaciSeqNo As String = "", _
                                                            Optional ByVal p_PackageNo As String = "", _
                                                            Optional ByVal p_PackageStepNo As Integer = 0, _
                                                            Optional ByVal p_AmountType As Integer = 0, _
                                                             Optional ByVal p_StepNo As Integer = 0, _
                                                            Optional ByRef p_CrossStepDate As String = "", _
                                                            Optional ByRef p_OrderNo As String = "", _
                                                            Optional ByRef p_Expiretype As Integer = 0) As Integer

        If blnChoose Then
            Return SF_GetAmount2(P_OPTION, P_CUSTID, P_CITEMCODE,
                                 P_PERIOD, P_REALSTARTDATE, P_REALSTOPDATE,
                                 p_SERVICETYPE, p_CompCode, P_SHOULDAMT, P_RETMSG)
        Else
            Return SF_GetAmount(P_OPTION, P_CUSTID, P_CITEMCODE, P_PERIOD, P_REALSTARTDATE,
                                       p_SERVICETYPE, p_CompCode, P_REALSTOPDATE, P_SHOULDAMT,
                                       P_RETMSG, p_RealPeriod, P_BPCode,
                                       intPFlag1, P_FaciSeqNo, p_PackageNo, p_PackageStepNo,
                                       p_AmountType, p_StepNo, p_CrossStepDate, p_OrderNo,
                                       p_Expiretype, New String(" ", 2000), New String(" ", 2000), 0, New String(" ", 2000))
        End If
    End Function
    Private Function SF_GetAmount2(ByVal P_OPTION As Integer, _
                ByVal P_CUSTID As Integer, _
                ByVal P_CITEMCODE As Integer, _
                ByVal P_PERIOD As Integer, _
                ByVal P_REALSTARTDATE As String, _
                ByVal P_REALSTOPDATE As String, _
                ByVal p_SERVICETYPE As String, _
                ByVal p_CompCode As Integer, _
                ByRef P_SHOULDAMT As Integer, _
                ByRef P_RETMSG As String) As Integer

        Dim InPut As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim OutPut As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim RetNum As Integer = 0
        InPut.Add("P_OPTION", P_OPTION)
        InPut.Add("P_CUSTID", P_CUSTID)
        InPut.Add("P_CITEMCODE", P_CITEMCODE)
        InPut.Add("P_PERIOD", P_PERIOD)
        InPut.Add("P_REALSTARTDATE", P_REALSTARTDATE)
        InPut.Add("P_REALSTOPDATE", P_REALSTOPDATE)
        InPut.Add("P_SERVICETYPE", p_SERVICETYPE)
        InPut.Add("P_COMPCODE", p_CompCode)
        OutPut.Add("P_SHOULDAMT", P_SHOULDAMT)

        OutPut.Add("P_RETMSG", P_RETMSG)
        DAO.ExecSF("SF_GETAMOUNT2", InPut, OutPut, RetNum)
        If (OutPut("P_RETMSG") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("P_RETMSG"))) Then
            P_RETMSG = OutPut("P_RETMSG").ToString
        End If
        If (OutPut("P_SHOULDAMT") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("P_SHOULDAMT"))) Then
            P_SHOULDAMT = OutPut("P_SHOULDAMT")
        End If

        Return RetNum
    End Function
    Private Function SF_GetAmount(ByVal P_OPTION As Integer, _
                ByVal P_CUSTID As Integer, ByVal P_CITEMCODE As Integer, _
                ByVal P_PERIOD As Integer, ByVal P_REALSTARTDATE As String, _
                ByVal p_SERVICETYPE As String, ByVal p_CompCode As Integer, _
                ByRef P_REALSTOPDATE As String, ByRef P_SHOULDAMT As Integer, _
                ByRef P_RETMSG As String, ByRef p_RealPeriod As Integer, _
                Optional ByVal P_BPCode As String = "", _
                Optional ByRef P_PFlag1 As Integer = 0, _
                Optional ByVal P_FaciSeqNo As String = "",
                Optional ByVal p_PackageNo As String = "", _
                Optional ByVal p_PackageStepNo As Integer = 0, Optional ByVal p_AmountType As Integer = 0, _
                Optional ByVal p_StepNo As Integer = 0, _
                Optional ByRef p_CrossStepDate As String = "", _
                Optional ByRef p_OrderNo As String = "", _
                Optional ByRef p_Expiretype As Integer = 0, _
                Optional ByRef p_BPCode1 As String = "", Optional ByRef p_BPName As String = "", _
                Optional ByRef p_PromCode As Integer = 0, Optional ByRef p_PromName As String = "") As Integer

        Dim InPut As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim OutPut As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        Dim RetNum As Integer = 0
        InPut.Add("P_OPTION", P_OPTION)
        InPut.Add("P_CUSTID", P_CUSTID)
        InPut.Add("P_CITEMCODE", P_CITEMCODE)
        InPut.Add("P_PERIOD", P_PERIOD)
        InPut.Add("P_REALSTARTDATE", P_REALSTARTDATE)
        InPut.Add("P_SERVICETYPE", p_SERVICETYPE)
        InPut.Add("P_COMPCODE", p_CompCode)
        InPut.Add("P_BPCODE", P_BPCode)
        InPut.Add("P_FACISEQNO", P_FaciSeqNo)
        InPut.Add("P_PACKAGENO", p_PackageNo)
        InPut.Add("P_PACKAGESTEPNO", p_PackageStepNo)
        InPut.Add("P_AMOUNTTYPE", p_AmountType)
        InPut.Add("P_STEPNO", p_StepNo)
        If P_REALSTOPDATE.Length = 0 Then
            P_REALSTOPDATE = New String(" ", 2000)
        End If
        OutPut.Add("P_REALSTOPDATE", P_REALSTOPDATE)
        OutPut.Add("P_SHOULDAMT", P_SHOULDAMT)
        OutPut.Add("P_REALPERIOD", p_RealPeriod)
        OutPut.Add("P_PUNISH", 0)
        OutPut.Add("P_DISCOUNTDATE1", New String(" ", 2000))
        OutPut.Add("P_DISCOUNTDATE2", New String(" ", 2000))
        OutPut.Add("P_PFLAG1", P_PFlag1)
        If p_CrossStepDate.Length = 0 Then
            p_CrossStepDate = New String(" ", 2000)
        End If
        OutPut.Add("p_CrossStepDate", p_CrossStepDate)
        OutPut.Add("p_OrderNo", p_OrderNo)
        OutPut.Add("p_Expiretype", p_Expiretype)
        OutPut.Add("p_BPCode1", p_BPCode1)
        OutPut.Add("p_BPName", p_BPName)
        OutPut.Add("p_PromCode", p_PromCode)
        OutPut.Add("p_PromName", p_PromName)
        OutPut.Add("P_RETMSG", P_RETMSG)
        DAO.ExecSF("SF_GETAMOUNT", InPut, OutPut, RetNum)
        If (OutPut("P_RETMSG") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("P_RETMSG").ToString)) Then
            P_RETMSG = OutPut("P_RETMSG").ToString
        End If
        If (OutPut("p_CrossStepDate") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("p_CrossStepDate").ToString)) Then
            p_CrossStepDate = OutPut("p_CrossStepDate")
        End If
        If (OutPut("P_REALSTOPDATE") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("P_REALSTOPDATE").ToString)) Then
            P_REALSTOPDATE = OutPut("P_REALSTOPDATE")
        End If
        If (OutPut("P_SHOULDAMT") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("P_SHOULDAMT").ToString)) Then
            P_SHOULDAMT = OutPut("P_SHOULDAMT")
        End If
        If (OutPut("P_REALPERIOD") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("P_REALPERIOD").ToString)) Then
            p_RealPeriod = OutPut("P_REALPERIOD")
        End If
        If (OutPut("P_PFlag1") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("P_PFlag1").ToString)) Then
            P_PFlag1 = OutPut("P_PFlag1")
        End If
        If (OutPut("p_OrderNo") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("p_OrderNo").ToString)) Then
            p_OrderNo = OutPut("p_OrderNo")
        End If
        If (OutPut("p_Expiretype") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("p_Expiretype").ToString)) Then
            p_Expiretype = OutPut("p_Expiretype")
        End If
        If (OutPut("p_BPCode1") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("p_BPCode1").ToString)) Then
            p_BPCode1 = OutPut("p_BPCode1")
        End If
        If (OutPut("p_BPName") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("p_BPName").ToString)) Then
            p_BPName = OutPut("p_BPName")
        End If
        If (OutPut("p_PromCode") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("p_PromCode").ToString)) Then
            p_PromCode = OutPut("p_PromCode")
        End If
        If (OutPut("p_PromName") IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(OutPut("p_PromName").ToString)) Then
            p_PromName = OutPut("p_PromName")
        End If
        Return RetNum
    End Function
    Private Function CancelBillNo(ByVal EntryType As Integer, ByVal BillNo As String, ByVal Item As Integer) As DataSet
        Dim dsReturn As DataSet = Nothing
        Dim dtError As DataTable = Nothing
        Dim dtOK As DataTable = Nothing
        Dim objBillUtility As New CableSoft.SO.BLL.Billing.Utility.Utility(Me.LoginInfo, DAO)
        Dim dtTemp As DataTable = DAO.ExecQry(_DAL.GetTempData(EntryType, BillNo, Item))
        dsReturn = CreateResultDataset(EntryType)
        dtOK = dsReturn.Tables(tbOKName)
        dtError = dsReturn.Tables(tbErrorName)
        Dim CMCode As Object = DBNull.Value
        Dim CMName As Object = DBNull.Value
        Dim PTCode As Object = DBNull.Value
        Dim PTName As Object = DBNull.Value
        Try
            For Each rwTemp As DataRow In dtTemp.Rows
                CMCode = DBNull.Value
                CMName = DBNull.Value
                PTCode = DBNull.Value
                PTName = DBNull.Value
                If Integer.Parse(DAO.ExecSclr(_DAL.IsPayOKOrCancel,
                                              New Object() {rwTemp.Item("BillNo")})) = 0 Then
                    Dim rwErr As DataRow = dtError.NewRow
                    For Each col As DataColumn In dtTemp.Columns
                        rwErr.Item(col.ColumnName) = rwTemp.Item(col.ColumnName)
                    Next
                    rwErr.Item(ErrorCodeFieldName) = "-1"
                    rwErr.Item(ErrorNameFieldName) = Language.CanNotCancel
                    dtError.Rows.Add(rwErr)
                    dtError.AcceptChanges()
                Else
                    Using dt1 As DataTable = DAO.ExecQry(_DAL.GetCMAndPTData,
                                                            New Object() {rwTemp.Item("CustId"),
                                                                          rwTemp.Item("FaciSeqNo"), rwTemp.Item("CitemCode")})
                        If dt1.Rows.Count > 0 Then
                            CMCode = dt1.Rows(0).Item("CMCode")
                            CMName = dt1.Rows(0).Item("CMName")
                            PTCode = dt1.Rows(0).Item("PTCode")
                            PTName = dt1.Rows(0).Item("PTName")
                        Else
                            Using drCMCode As DataTable = DAO.ExecQry(_DAL.GetDefaultCMCode)
                                CMCode = drCMCode.Rows(0).Item("CodeNo").ToString
                                CMName = drCMCode.Rows(0).Item("Description").ToString
                            End Using
                            Using drPTCode As DataTable = DAO.ExecQry(_DAL.GetDefaultPTCode)

                                PTCode = drPTCode.Rows(0).Item("CodeNo").ToString
                                PTName = drPTCode.Rows(0).Item("Description").ToString
                            End Using
                        End If
                        DAO.ExecNqry(_DAL.CancelCharge, New Object() {rwTemp.Item("SUCCode"),
                                                                      rwTemp("SUCName"),
                                                                      CMCode, CMName, PTCode, PTName,
                                                                      rwTemp("BillNo"), rwTemp("Item")})
                        Using tbSO033 As DataTable = DAO.ExecQry(_DAL.GetChargeData,
                                                                  New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})
                            For Each rwSO033 As DataRow In tbSO033.Rows
                                objBillUtility.UpdateVODPoint(rwSO033)
                            Next
                        End Using
                        DAO.ExecNqry(_DAL.DelTempData(EntryType), New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})
                        Dim rwOK As DataRow = dtOK.NewRow
                        rwOK.ItemArray = rwTemp.ItemArray
                        dtOK.Rows.Add(rwOK)



                    End Using
                    'Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetCMAndPTData,
                    '                                         New Object() {rwTemp.Item("CustId"),
                    '                                                       rwTemp.Item("FaciSeqNo"), rwTemp.Item("CitemCode")})
                    '    If dr.HasRows Then
                    '        While dr.Read
                    '            CMCode = dr.Item("CMCode")
                    '            CMName = dr.Item("CMName")
                    '            PTCode = dr.Item("PTCode")
                    '            PTName = dr.Item("PTName")
                    '        End While
                    '    Else
                    '        Using drCMCode As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefaultCMCode)
                    '            While drCMCode.Read
                    '                CMCode = drCMCode.Item("CodeNo").ToString
                    '                CMName = drCMCode.Item("Description").ToString
                    '            End While
                    '        End Using
                    '        Using drPTCode As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefaultPTCode)
                    '            While drPTCode.Read
                    '                PTCode = drPTCode.Item("CodeNo").ToString
                    '                PTName = drPTCode.Item("Description").ToString
                    '            End While
                    '        End Using
                    '    End If
                    '    DAO.ExecNqry(_DAL.CancelCharge, New Object() {rwTemp.Item("SUCCode"),
                    '                                                  rwTemp("SUCName"),
                    '                                                  CMCode, CMName, PTCode, PTName,
                    '                                                  rwTemp("BillNo"), rwTemp("Item")})
                    '    Using tbSO033 As DataTable = DAO.ExecQry(_DAL.GetChargeData,
                    '                                              New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})
                    '        For Each rwSO033 As DataRow In tbSO033.Rows
                    '            objBillUtility.UpdateVODPoint(rwSO033)
                    '        Next
                    '    End Using
                    '    DAO.ExecNqry(_DAL.DelTempData(EntryType), New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})
                    '    Dim rwOK As DataRow = dtOK.NewRow
                    '    rwOK.ItemArray = rwTemp.ItemArray
                    '    dtOK.Rows.Add(rwOK)
                    'End Using
                End If
            Next


        Catch ex As Exception
            Throw
        Finally
            objBillUtility.Dispose()
            objBillUtility = Nothing
        End Try
        Return dsReturn
    End Function
    Public Function CancelBillData(ByVal EntryType As Integer, ByVal BillNo As String, ByVal Item As Integer) As DataSet
        Dim dsReturn As DataSet = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing

        Dim blnAutoClose As Boolean = False
        'Dim objBillUtility As New CableSoft.SO.BLL.Billing.Utility.Utility(Me.LoginInfo, DAO)
        Dim dtError As DataTable = Nothing
        Dim dtOK As DataTable = Nothing
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
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId)
        '        Dim dtTemp As DataTable = DAO.ExecQry(_DAL.GetTempData(EntryType), New Object() {LoginInfo.EntryName})
        Dim dtTemp As DataTable = DAO.ExecQry(_DAL.GetTempData(EntryType), New Object() {LoginInfo.EntryId})
        dsReturn = CreateResultDataset(EntryType)
        dtOK = dsReturn.Tables(tbOKName)
        dtError = dsReturn.Tables(tbErrorName)
        Try
            Using dsCancelData As DataSet = CancelBillNo(EntryType,
                                                             BillNo, _
                                                             Item)
                If dsCancelData.Tables(tbOKName).Rows.Count > 0 Then
                    For Each rwOK As DataRow In dsCancelData.Tables(tbOKName).Rows
                        Dim rwNew As DataRow = dtOK.NewRow
                        rwNew.ItemArray = rwOK.ItemArray
                        dtOK.Rows.Add(rwNew)
                    Next
                    dtOK.AcceptChanges()
                End If
                If dsCancelData.Tables(tbErrorName).Rows.Count > 0 Then
                    For Each rwError As DataRow In dsCancelData.Tables(tbErrorName).Rows
                        Dim rwNew As DataRow = dtError.NewRow
                        rwNew.ItemArray = rwError.ItemArray
                        dtError.Rows.Add(rwNew)
                    Next
                    dtError.AcceptChanges()
                End If
            End Using
            Using dsInfo As DataSet = GetTempAllInfo(EntryType)
                dsReturn.Tables.Add(dsInfo.Tables(tbInfoName).Copy)
                dsReturn.Tables.Add(dsInfo.Tables(tbTemp).Copy)
            End Using
            trans.Commit()
        Catch ex As Exception
            If trans IsNot Nothing Then
                trans.Rollback()
            End If
            Throw
        Finally
            'If objBillUtility IsNot Nothing Then
            '    objBillUtility.Dispose()
            'End If

            dtError.Dispose()
            dsReturn.Dispose()
            'If cmd IsNot Nothing Then
            '    cmd.Dispose()

            'End If
            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If

            End If
        End Try
        Return dsReturn
    End Function
    Public Function CancelAllBillData(ByVal EntryType As Integer) As DataSet

        Dim dsReturn As DataSet = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing

        Dim blnAutoClose As Boolean = False
        'Dim objBillUtility As New CableSoft.SO.BLL.Billing.Utility.Utility(Me.LoginInfo, DAO)
        Dim dtError As DataTable = Nothing
        Dim dtOK As DataTable = Nothing
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

        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId)
        '        Dim dtTemp As DataTable = DAO.ExecQry(_DAL.GetTempData(EntryType), New Object() {LoginInfo.EntryName})
        Dim dtTemp As DataTable = DAO.ExecQry(_DAL.GetTempData(EntryType), New Object() {LoginInfo.EntryId})
        dsReturn = CreateResultDataset(EntryType)
        dtOK = dsReturn.Tables(tbOKName)
        dtError = dsReturn.Tables(tbErrorName)
        'Dim CMCode As Object = DBNull.Value
        'Dim CMName As Object = DBNull.Value
        'Dim PTCode As Object = DBNull.Value
        'Dim PTName As Object = DBNull.Value
        Try
            For Each rwTemp As DataRow In dtTemp.Rows
                Using dsCancelData As DataSet = CancelBillNo(EntryType,
                                                             rwTemp.Item("BillNo").ToString, _
                                                             Integer.Parse(rwTemp.Item("Item")))
                    If dsCancelData.Tables(tbOKName).Rows.Count > 0 Then
                        For Each rwOK As DataRow In dsCancelData.Tables(tbOKName).Rows
                            Dim rwNew As DataRow = dtOK.NewRow
                            rwNew.ItemArray = rwOK.ItemArray
                            dtOK.Rows.Add(rwNew)
                        Next
                        dtOK.AcceptChanges()
                    End If
                    If dsCancelData.Tables(tbErrorName).Rows.Count > 0 Then
                        For Each rwError As DataRow In dsCancelData.Tables(tbErrorName).Rows
                            Dim rwNew As DataRow = dtError.NewRow
                            rwNew.ItemArray = rwError.ItemArray
                            dtError.Rows.Add(rwNew)
                        Next
                        dtError.AcceptChanges()
                    End If
                End Using

                'CMCode = DBNull.Value
                'CMName = DBNull.Value
                'PTCode = DBNull.Value
                'PTName = DBNull.Value
                'If Integer.Parse(DAO.ExecSclr(_DAL.IsPayOKOrCancel, _
                '                              New Object() {rwTemp.Item("BillNo")})) = 0 Then
                '    Dim rwErr As DataRow = dtError.NewRow
                '    For Each col As DataColumn In dtTemp.Columns
                '        rwErr.Item(col.ColumnName) = rwTemp.Item(col.ColumnName)
                '    Next
                '    rwErr.Item(ErrorCodeFieldName) = "-1"
                '    rwErr.Item(ErrorNameFieldName) = "該單據編號有資料已收或作廢, 不允許取消登錄!!"
                '    dtError.Rows.Add(rwErr)
                '    dtError.AcceptChanges()
                'Else
                '    Using dr As DbDataReader = DAO.ExecDtRdr(_DAL.GetCMAndPTData,
                '                                             New Object() {rwTemp.Item("CustId"),
                '                                                           rwTemp.Item("FaciSeqNo"), rwTemp.Item("CitemCode")})
                '        If dr.HasRows Then
                '            While dr.Read
                '                If Not DBNull.Value.Equals(dr.Item("CMCode")) Then
                '                    CMCode = dr.Item("CMCode")
                '                    CMName = dr.Item("CMName")
                '                End If
                '                If Not DBNull.Value.Equals(dr.Item("PTCode")) Then
                '                    PTCode = dr.Item("PTCode")
                '                    PTName = dr.Item("PTName")
                '                End If                                
                '            End While
                '        Else
                '            Using drCMCode As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefaultCMCode)
                '                While drCMCode.Read
                '                    CMCode = drCMCode.Item("CodeNo").ToString
                '                    CMName = drCMCode.Item("Description").ToString
                '                End While
                '            End Using
                '            Using drPTCode As DbDataReader = DAO.ExecDtRdr(_DAL.GetDefaultPTCode)
                '                While drPTCode.Read
                '                    PTCode = drPTCode.Item("CodeNo").ToString
                '                    PTName = drPTCode.Item("Description").ToString
                '                End While
                '            End Using
                '        End If
                '        DAO.ExecNqry(_DAL.CancelCharge, New Object() {rwTemp.Item("SUCCode"),
                '                                                      rwTemp("SUCName"),
                '                                                      CMCode, CMName, PTCode, PTName,                                                                     
                '                                                      rwTemp("BillNo"), rwTemp("Item")})
                '        Using tbSO033 As DataTable = DAO.ExecQry(_DAL.GetChargeData,
                '                                                  New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})
                '            For Each rwSO033 As DataRow In tbSO033.Rows
                '                objBillUtility.UpdateVODPoint(rwSO033)
                '            Next
                '        End Using
                '        DAO.ExecNqry(_DAL.DelTempData(EntryType), New Object() {rwTemp.Item("BillNo"), rwTemp.Item("Item")})
                '        Dim rwOK As DataRow = dtOK.NewRow
                '        rwOK.ItemArray = rwTemp.ItemArray
                '        dtOK.Rows.Add(rwOK)
                '    End Using
                'End If
            Next
            'dsReturn.Tables.Add(dtError)
            'dsReturn.Tables.Add(dtOK)
            trans.Commit()

            Return dsReturn
        Catch ex As Exception
            If trans IsNot Nothing Then
                trans.Rollback()
            End If
            Throw ex
        Finally
            'If objBillUtility IsNot Nothing Then
            '    objBillUtility.Dispose()

            'End If

            dtError.Dispose()
            dsReturn.Dispose()
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If
            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                'If CSLog IsNot Nothing Then
                '    CSLog.Dispose()
                'End If
            End If
        End Try

    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                End If
                _DAL.Dispose()
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
