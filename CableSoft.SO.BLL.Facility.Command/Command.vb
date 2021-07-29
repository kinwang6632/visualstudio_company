Imports CableSoft.SO.BLL.Facility
Imports CableSoft.BLL.Utility
Imports System.Threading
Imports System.Reflection
Imports System.Data.Common
'Imports Lang = CableSoft.SO.BLL.Facility.Command.CommandLanguage
Public Class Command
    Inherits BLLBasic
    Implements IDisposable
    Private WaitTimeOut As Int32 = 30
    Private OwnerName As String = Nothing
    Private TakeTableName As String = Nothing
    Private _DAL As New CommandDALMultiDB(Me.LoginInfo.Provider)
    Private tbMaster As DataTable = Nothing
    Private tbDetail As DataTable = Nothing
    Private fFieldsAndValues As New Dictionary(Of String, Object)
    Private fDetailFieldsAndValues As Dictionary(Of String, Object) = Nothing
    Private fSeqNo As Object = Nothing
    Private fSeqNoFieldName As String = Nothing

    Private evn As New AutoResetEvent(False)
    Private ti As Threading.RegisteredWaitHandle
    'Private CmdStatusField As String = Nothing
    'Private ErrorCodeField As String = Nothing
    'Private ErrorMsgField As String = Nothing
    'Private SuccessCodeStr As String = Nothing
    'Private ErrorCodeStr As String = Nothing
    Private CmdStatusField As Object = Nothing
    Private ErrorCodeField As Object = Nothing
    Private ErrorMsgField As Object = Nothing
    Private SuccessCodeStr As Object = Nothing
    Private ErrorCodeStr As Object = Nothing
    Private daoCmd As CableSoft.Utility.DataAccess.DAO = Nothing

    Private Const ReturnMasterName As String = "Master"
    Private Const ReturnDetailName As String = "Detail"
    Private Const QueryOKExit As String = "QueryOKExit"
    Private Const QueryRIAResult As String = "RIAResult"
    Private Const SourceDetail As String = "Detail"
    Private Const DefaultField As String = "DefaultValue"
    Private Const TimeOutCodeNo As Integer = -999
    Private IsQueryData As Boolean = False
    Private Lang As New CableSoft.BLL.Language.SO61.CommandLanguage
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub
    Public Function InsertCommand(ByVal TableName As String,
                             ByVal CMDID As String, ByVal WipData As DataSet,
                                ByVal NoWait As Boolean,
                                ByVal WriteTimeOut As Boolean,
                                ByVal DeleteData As Boolean) As RIAResult

        Dim aRet As New RIAResult()
        Dim thrResult As New ThreadLocal(Of Dictionary(Of String, Object))
        daoCmd = New CableSoft.Utility.DataAccess.DAO(Me.LoginInfo.Provider, Me.LoginInfo.ConnectionString)
        daoCmd.AutoCloseConn = False
        thrResult.Value = New Dictionary(Of String, Object)
        aRet.ResultBoolean = False
        Dim trans As DbTransaction = Nothing
        Dim cn As DbConnection = daoCmd.GetConn
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If
        trans = cn.BeginTransaction
        daoCmd.Transaction = trans        

        Try
            '讀取設定檔資料
            If Not TakeTable(TableName, CMDID) Then
                aRet.ErrorCode = -1
                aRet.ErrorMessage = Lang.TakeSetErr
                Return aRet
            End If


            ' DAO.AutoCloseConn = False
            '取出新增命令的語法

            Dim aSQL As String = TakeInsertSQL(tbMaster, tbDetail, WipData)
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, tbMaster.Rows(0).Item("Caption"))
            Try
                daoCmd.ExecNqry(aSQL, fFieldsAndValues.Values.ToArray)
                If (Not DBNull.Value.Equals(tbMaster.Rows(0).Item("DTableName"))) AndAlso
                    (WipData.Tables.Contains(SourceDetail)) AndAlso
                    (WipData.Tables(SourceDetail).Rows.Count > 0) Then
                    If chkDetailSchema(WipData.Tables(SourceDetail)) Then
                        For Each rw As DataRow In WipData.Tables(SourceDetail).Rows
                            aSQL = TakeInsertDetailSQL(tbMaster, rw)
                            If fDetailFieldsAndValues.Values.Count > 0 Then
                                daoCmd.ExecNqry(aSQL, fDetailFieldsAndValues.Values.ToArray)
                            End If

                        Next
                    End If


                End If
                trans.Commit()

                If NoWait Then
                    aRet.ResultBoolean = True
                    Dim ds As New DataSet
                    ds.Tables.Add(daoCmd.ExecQry(_DAL.QuertyStatus(TakeTableName, fSeqNoFieldName), New Object() {fSeqNo}).Copy)
                    ds.Tables(0).TableName = ReturnMasterName
                    aRet.ResultDataSet = ds
                    Return aRet
                End If
            Catch ex As Exception

                'trans.Rollback()

                aRet.ErrorCode = -111
                aRet.ErrorMessage = ex.Message
                Return aRet
            End Try


            thrResult.Value.Add(QueryOKExit, False)
            thrResult.Value.Add(QueryRIAResult, aRet)

            IsQueryData = False
            ti = ThreadPool.RegisterWaitForSingleObject( _
                            evn, New WaitOrTimerCallback(AddressOf WaitProc), _
                            thrResult.Value, 1000, False)

            evn.WaitOne(TimeSpan.FromSeconds(WaitTimeOut))
            'evn.Set()
            
            If Not thrResult.Value.Item(QueryOKExit) Then
                thrResult.Value.Item(QueryOKExit) = True

                If WriteTimeOut Then
                    trans = cn.BeginTransaction
                    WriteTimeOutError()
                    trans.Commit()
                End If

                aRet.ErrorCode = Int32.Parse(TimeOutCodeNo)
                aRet.ErrorMessage = Lang.CmdTimeOut                
                If (daoCmd IsNot Nothing) AndAlso (daoCmd.GetConn.State = ConnectionState.Open) Then
                    Dim ds As New DataSet
                    ds.Tables.Add(daoCmd.ExecQry(_DAL.QuertyStatus(TakeTableName, fSeqNoFieldName), New Object() {fSeqNo}).Copy)
                    ds.Tables(0).TableName = ReturnMasterName
                    'Dim dtDetail As DataTable = GetReturnDetailTable()
                    'If dtDetail IsNot Nothing Then
                    '    ds.Tables.Add(dtDetail)
                    'End If
                    aRet.ResultDataSet = ds
                End If
            End If
            If (Not NoWait) AndAlso (aRet.ResultBoolean) AndAlso (DeleteData) Then
                trans = cn.BeginTransaction
                DeleteCMDData(fSeqNo)
                trans.Commit()
            End If
            If ti IsNot Nothing Then
                ti.Unregister(Nothing)
            End If
            'Thread.Sleep(100)
            If (Not aRet.ResultBoolean) AndAlso (String.IsNullOrEmpty(aRet.ErrorMessage)) Then
                aRet.ErrorMessage = Lang.CmdOtherErr
            End If

        Catch ex As Exception

            aRet.ErrorCode = -999
            aRet.ErrorMessage = ex.Message
            Return aRet
        Finally
            CableSoft.BLL.Utility.Utility.ClearClientInfo(daoCmd)

            If trans IsNot Nothing Then
                trans.Dispose()
            End If
            If cn IsNot Nothing Then
                cn.Close()
                cn.Dispose()
            End If

            'DAO.AutoCloseConn = True

            If ti IsNot Nothing Then
                ti.Unregister(Nothing)
            End If
            If tbDetail IsNot Nothing Then
                tbDetail.Dispose()
            End If
            If tbMaster IsNot Nothing Then
                tbMaster.Dispose()
            End If

            fFieldsAndValues.Clear()
            fFieldsAndValues = Nothing
            If fDetailFieldsAndValues IsNot Nothing Then
                fDetailFieldsAndValues.Clear()
                fDetailFieldsAndValues = Nothing
            End If
            evn.Dispose()
            thrResult.Value.Clear()
            thrResult.Dispose()
            If daoCmd IsNot Nothing Then
                daoCmd.Dispose()
                daoCmd = Nothing
            End If
        End Try

        Return aRet

    End Function
    Private Sub WriteTimeOutError()
        Try
            If String.IsNullOrEmpty(OwnerName) Then
                TakeTableName = tbMaster.Rows(0).Item("TableName")
            Else
                OwnerName = OwnerName.Replace(".", "") & "."
                TakeTableName = OwnerName & tbMaster.Rows(0).Item("TableName")
            End If

            If ErrorCodeField <> ErrorMsgField Then
                daoCmd.ExecNqry(_DAL.WriteTimeOutError(TakeTableName,
                                                    CmdStatusField,
                                                    ErrorCodeField, ErrorMsgField, fSeqNoFieldName),
                                                New Object() {TimeOutCodeNo, Lang.CmdTimeOut, fSeqNo})
            Else
                daoCmd.ExecNqry(_DAL.WriteTimeOutError(TakeTableName,
                                                   CmdStatusField,
                                                   ErrorCodeField, ErrorMsgField, fSeqNoFieldName),
                                               New Object() {TimeOutCodeNo, fSeqNo})

            End If
            
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub DeleteCMDData(ByVal SeqNo As String)
        If String.IsNullOrEmpty(OwnerName) Then
            TakeTableName = tbMaster.Rows(0).Item("TableName")
        Else
            OwnerName = OwnerName.Replace(".", "") & "."
            TakeTableName = OwnerName & tbMaster.Rows(0).Item("TableName")
        End If
        daoCmd.ExecNqry(_DAL.DeleteCMDData(TakeTableName, fSeqNoFieldName), New Object() {fSeqNo})

    End Sub
    Private Sub WaitProc(state As Object, timedOut As Boolean)
        If IsQueryData Then
            Exit Sub
        End If
        SyncLock state
            IsQueryData = True
            If (state IsNot Nothing) AndAlso (TypeOf state Is Dictionary(Of String, Object)) Then
                If (CType(state, Dictionary(Of String, Object)).Keys.Contains(QueryOKExit)) AndAlso
                                (CType(state, Dictionary(Of String, Object)).Item(QueryOKExit)) Then
                    Debug.Print("Exit 1")
                    IsQueryData = False
                    Exit Sub
                End If
                If (Not CType(state, Dictionary(Of String, Object)).Keys.Contains(QueryOKExit)) Then
                    Debug.Print("Exit 2")
                    IsQueryData = False
                    Exit Sub
                End If
            End If


            Try
                Using tbResult As New ThreadLocal(Of DataTable)

                    tbResult.Value = daoCmd.ExecQry(_DAL.QuertyStatus(TakeTableName, fSeqNoFieldName), New Object() {fSeqNo})
                    Try
                        If SuccessCodeStr.Contains(tbResult.Value.Rows(0).Item(CmdStatusField).ToString.ToUpper) Then
                            CType(state, Dictionary(Of String, Object)).Item(QueryOKExit) = True
                            CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ResultBoolean = True


                            Exit Sub
                        End If
                        If ErrorCodeStr.Contains(tbResult.Value.Rows(0).Item(CmdStatusField).ToString.ToUpper) Then
                            CType(state, Dictionary(Of String, Object)).Item(QueryOKExit) = True
                            CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ResultBoolean = False
                            If CableSoft.Utility.DataAccess.IsNum(tbResult.Value.Rows(0).Item(ErrorCodeField)) Then
                                CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ErrorCode = tbResult.Value.Rows(0).Item(ErrorCodeField)
                            Else
                                CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ErrorCode = -999
                            End If
                            If String.IsNullOrEmpty(tbResult.Value.Rows(0).Item(ErrorMsgField).ToString) Then
                                CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ErrorMessage =
                                    CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ErrorCode
                            Else
                                CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ErrorMessage = tbResult.Value.Rows(0).Item(ErrorMsgField)
                            End If
                            'Dim dsReturn As New DataSet()
                            'dsReturn.Tables.Add(tbResult.Value.Copy)
                            'dsReturn.Tables(0).TableName = "Master"
                            'CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ResultDataSet = dsReturn

                            'ti.Unregister(Nothing)
                            'evn.Set()

                            Exit Sub
                        End If

                    Finally
                        If (state IsNot Nothing) AndAlso (TypeOf state Is Dictionary(Of String, Object)) Then
                            If (CType(state, Dictionary(Of String, Object)).Keys.Contains(QueryOKExit)) AndAlso
                                (CType(state, Dictionary(Of String, Object)).Item(QueryOKExit)) Then
                                Dim dsReturn As New DataSet()
                                dsReturn.Tables.Add(tbResult.Value.Copy)
                                dsReturn.Tables(0).TableName = ReturnMasterName
                                Dim dtDetail As DataTable = GetReturnDetailTable()
                                If dtDetail IsNot Nothing Then
                                    dsReturn.Tables.Add(dtDetail)
                                End If

                                CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ResultDataSet = dsReturn

                                ti.Unregister(Nothing)
                                evn.Set()

                            End If
                        End If

                        tbResult.Value.Dispose()
                    End Try


                End Using
            Catch ex As Exception
                If (state IsNot Nothing) AndAlso (TypeOf state Is Dictionary(Of String, Object)) Then

                    If (CType(state, Dictionary(Of String, Object)).Keys.Contains(QueryOKExit)) Then
                        CType(state, Dictionary(Of String, Object)).Item(QueryOKExit) = True
                        CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ResultBoolean = False
                        CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ErrorCode = -999
                        CType(CType(state, Dictionary(Of String, Object)).Item(QueryRIAResult), RIAResult).ErrorMessage = ex.Message & "--WaitProc"
                    End If

                    ti.Unregister(Nothing)
                    evn.Set()
                End If
            Finally
                IsQueryData = False
            End Try

        End SyncLock
    End Sub
    Private Function GetReturnDetailTable() As DataTable
        Dim dtReturn As DataTable = Nothing
        Try
            If DBNull.Value.Equals(tbMaster.Rows(0).Item("DTableName")) Then
                Return Nothing
            End If
            Dim DTableName As String = tbMaster.Rows(0).Item("DTableName")

            If Not String.IsNullOrEmpty(OwnerName) Then
                DTableName = OwnerName & DTableName
            End If
            'Dim aSQL As String = String.Format("SELECT * FROM {0} WHERE {1} = ", DTableName, fSeqNoFieldName)
            Dim aSQL As String = _DAL.GetReturnDetailTable(DTableName, fSeqNoFieldName)
            aSQL = String.Format(aSQL & "{0}0", _DAL.Sign)
            Dim dt As DataTable = daoCmd.ExecQry(aSQL, New Object() {fSeqNo})
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                dtReturn = dt.Copy
                dtReturn.TableName = ReturnDetailName
            End If
            
        Catch ex As Exception
            Throw New Exception(ex.Message & "--GetReturnDetailTable")
        End Try

        Return dtReturn
    End Function
    Private Sub StartTimer(thrResult As Object)
        Thread.Sleep(5 * 1000)
        'thrTimer = New Threading.Timer(AddressOf QueryStatus, thrResult, 0, 100)
        'evn.WaitOne(WaitTimeOut * 1000)
    End Sub
    Private Sub QueryStatus(Result As Object)
        'thrTimer.Change(Timeout.Infinite, Timeout.Infinite)
        'SyncLock Result
        '    Try
        '        If CType(Result, List(Of Object)).Item(0) = True Then
        '            thrTimer.Change(Timeout.Infinite, Timeout.Infinite)
        '            Exit Sub
        '        End If
        '        Using tbResult As New ThreadLocal(Of DataTable)
        '            tbResult.Value = DAO.ExecQry(_DAL.QuertyStatus(TakeTableName, fSeqNoFieldName), New Object() {fSeqNo})
        '            Select Case tbResult.Value.Rows(0).Item("CmdStatus").ToString.ToUpper
        '                Case "S"
        '                    CType(Result, List(Of Object)).Item(0) = True
        '                    CType(CType(Result, List(Of Object)).Item(1), RIAResult).ResultBoolean = True
        '                    CType(Result, List(Of Object)).Item(2) = True
        '                    evn.Set()
        '                    'thrTimer.Change(Timeout.Infinite, Timeout.Infinite)
        '                Case "E"
        '                    CType(Result, List(Of Object)).Item(0) = True
        '                    CType(CType(Result, List(Of Object)).Item(1), RIAResult).ResultBoolean = False
        '                    CType(CType(Result, List(Of Object)).Item(1), RIAResult).ErrorCode = -3
        '                    CType(CType(Result, List(Of Object)).Item(1), RIAResult).ErrorMessage = "1111"
        '                    evn.Set()
        '                Case Else
        '                    If Not CType(Result, List(Of Object)).Item(2) Then
        '                        thrTimer.Change(0, 100)
        '                    End If

        '                    'thrTimer.Change(Timeout.Infinite, Timeout.Infinite)
        '            End Select
        '        End Using
        '    Catch ex As Exception

        '        evn.Set()
        '        thrTimer.Change(Timeout.Infinite, Timeout.Infinite)
        '    Finally

        '    End Try
        'End SyncLock

    End Sub
    Private Function TakeInsertDetailSQL(ByVal tbMaster As DataTable,
                                   ByVal rw As DataRow) As String
        Dim aRet As String = Nothing
        Dim aFields As String = Nothing
        Dim aValues As String = Nothing
        fDetailFieldsAndValues = New Dictionary(Of String, Object)
        Dim DTableName As String = tbMaster.Rows(0).Item("DTableName")
        Try
            If Not String.IsNullOrEmpty(OwnerName) Then
                DTableName = OwnerName & DTableName
            End If
            For Each col As DataColumn In rw.Table.Columns
                If col.ColumnName.ToUpper = fSeqNoFieldName.ToUpper Then
                    fDetailFieldsAndValues.Add(col.ColumnName, fSeqNo)
                Else
                    If Not DBNull.Value.Equals(rw.Item(col.ColumnName)) Then
                        fDetailFieldsAndValues.Add(col.ColumnName, rw.Item(col.ColumnName))
                    End If

                End If
            Next
            If rw.Table.Columns.Cast(Of DataColumn).Where(Function(col As DataColumn)
                                                              Return col.ColumnName.ToUpper = fSeqNoFieldName.ToUpper
                                                          End Function).Count = 0 Then
                fDetailFieldsAndValues.Add(fSeqNoFieldName, fSeqNo)
            End If


            For i As Int32 = 0 To fDetailFieldsAndValues.Count - 1
                If String.IsNullOrEmpty(aFields) Then
                    aFields = fDetailFieldsAndValues.Keys(i)
                Else
                    aFields = aFields & "," & fDetailFieldsAndValues.Keys(i)
                End If
                If String.IsNullOrEmpty(aValues) Then
                    aValues = "{0}0"
                Else
                    aValues = aValues & ",{0}" & i
                End If
            Next
            'aRet = String.Format("INSERT INTO {0} ( {1} ) VALUES ({2})", DTableName, aFields, aValues)
            aRet = _DAL.GetTakeInsertDetailSQL(DTableName, aFields, aValues)
            aRet = String.Format(aRet, _DAL.Sign)
        Catch ex As Exception
            Throw New Exception(ex.Message & "--TakeInsertDetailSQL")
        End Try

        Return aRet
    End Function
    Private Function chkDetailSchema(ByVal tbCheck As DataTable) As Boolean
        Try
            Dim chkTableName As String = tbMaster.Rows(0).Item("DTableName")
            If Not String.IsNullOrEmpty(OwnerName) Then
                chkTableName = OwnerName & chkTableName
            End If
            '            Dim aSQL As String = "SELECT * FROM " & chkTableName & " WHERE 1=0 "
            Dim aSQL As String = _DAL.GetChkDetailSchema(chkTableName)
            Dim tbScheam As DataTable = daoCmd.ExecQry(aSQL)
            For Each col As DataColumn In tbCheck.Columns
                If tbScheam.Columns.Cast(Of DataColumn).Where(Function(colSchema As DataColumn)
                                                                  Return colSchema.ColumnName.ToUpper = col.ColumnName.ToUpper
                                                              End Function).Count = 0 Then
                    Throw New Exception(String.Format(Lang.FieldNotExistsTable, col.ColumnName, tbMaster.Rows(0).Item("DTableName")))
                End If
            Next
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message & "--chkDetailSchema")
        End Try

    End Function
    Public Function ConverDataType(ByVal FieldName As String, ByVal tbSchema As DataTable, ByVal Value As Object) As Object
        Dim aRet As Object = Nothing
        Try
            If Not tbSchema.Columns.Contains(FieldName) Then
                Throw New Exception(String.Format(Lang.NotFoundRealField, FieldName))
            End If
            If Value.GetType.Equals(tbSchema.Columns(FieldName).DataType) Then
                Return Value
            End If
            aRet = Convert.ChangeType(Value, tbSchema.Columns(FieldName).DataType)
        Catch ex As Exception
            Throw New Exception(String.Format(Lang.ConverDataTypeError,
                                              FieldName, Value.ToString,
                                              tbSchema.Columns(FieldName).DataType.Name))
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 取出新增Insert 的SQL語法
    ''' </summary>
    ''' <param name="tbMaster">SO1102A</param>
    ''' <param name="tbDetail">SO1102B</param>
    ''' <param name="WipData">Insert Values</param>
    ''' <returns>Insert SQL</returns>
    ''' <remarks></remarks>
    Private Function TakeInsertSQL(ByVal tbMaster As DataTable, ByVal tbDetail As DataTable,
                                   ByVal WipData As DataSet) As String
        If String.IsNullOrEmpty(OwnerName) Then
            TakeTableName = tbMaster.Rows(0).Item("TableName")
        Else
            OwnerName = OwnerName.Replace(".", "") & "."
            TakeTableName = OwnerName & tbMaster.Rows(0).Item("TableName")
        End If
        Dim aRet As String = Nothing
        Dim tbSchema As DataTable = daoCmd.ExecQry(_DAL.GetSchemaTable(TakeTableName))
        Const LoginInfo As String = "LoginInfo"
        Const SeqNo As String = "SeqNo"
        Dim aHaveDefault As Boolean = False
        Try
            Dim intType As Int32 = -1
            For Each rw As DataRow In tbDetail.Rows
                aHaveDefault = False
                    If DBNull.Value.Equals(rw.Item("FieldType")) Then
                        intType = 0
                    Else
                        intType = Int32.Parse(rw.Item("FieldType"))
                    End If
                    Select Case intType
                    Case 0, 4
                        If Not DBNull.Value.Equals(rw.Item(DefaultField)) Then
                            fFieldsAndValues.Add(rw.Item("FieldName"), ConverDataType(rw.Item("FieldName"),
                                                                                                          tbSchema, rw.Item(DefaultField)))
                            aHaveDefault = True
                        End If
                        Select Case rw.Item("SourceTable").ToString.ToUpper
                            Case LoginInfo.ToUpper
                                Dim aProperty As Object = GetLoginInfoValue(rw.Item("SourceField"))
                                If aProperty IsNot Nothing Then
                                    fFieldsAndValues.Add(rw.Item("FieldName"), ConverDataType(rw.Item("FieldName"),
                                                                                              tbSchema, aProperty))
                                End If
                            Case SeqNo.ToUpper
                                fSeqNoFieldName = rw.Item("FieldName")
                                If Not aHaveDefault Then
                                    fFieldsAndValues.Add(rw.Item("FieldName"), ConverDataType(rw.Item("FieldName"),
                                                                                          tbSchema,
                                                                                          GetSeqNo(rw.Item("SourceField"))))
                                End If

                            Case Else
                                If Not aHaveDefault Then
                                    If Not WipData.Tables.Contains(rw.Item("SourceTable")) Then
                                        Throw New Exception("WipData SourceTable Is Null！--AutoSerialNo = " & rw.Item("AutoSerialNo"))
                                    End If
                                    If WipData.Tables(rw.Item("SourceTable").ToString).Rows.Count > 0 Then
                                        If Not tbSchema.Columns.Contains(rw.Item("FieldName")) Then
                                            Throw New Exception(String.Format(Lang.NotFoundField, TakeTableName, rw.Item("FieldName")))
                                        End If
                                        If Not DBNull.Value.Equals(WipData.Tables(rw.Item("SourceTable").ToString).Rows(0).Item(rw.Item("SourceField").ToString)) Then
                                            fFieldsAndValues.Add(rw.Item("FieldName"), ConverDataType(rw.Item("FieldName"), tbSchema,
                                                                 WipData.Tables(rw.Item("SourceTable").ToString).Rows(0).Item(rw.Item("SourceField").ToString)))
                                        End If
                                    End If
                                End If
                        End Select
                    Case 1
                        CmdStatusField = rw.Item("FieldName")
                        If Not DBNull.Value.Equals(rw.Item(DefaultField)) Then
                            fFieldsAndValues.Add(rw.Item("FieldName"), ConverDataType(rw.Item("FieldName"),
                                                                                                          tbSchema, rw.Item(DefaultField)))
                        Else
                            If (DBNull.Value.Equals(rw.Item("SourceTable"))) AndAlso (DBNull.Value.Equals(rw.Item(DefaultField))) Then
                                fFieldsAndValues.Add(CmdStatusField, "W")
                            Else
                                If DBNull.Value.Equals(rw("SourceField")) Then
                                    Throw New Exception("SourceField Is Null！--AutoSerialNo = " & rw.Item("AutoSerialNo"))
                                End If
                                If Not WipData.Tables.Contains(rw.Item("SourceTable")) Then
                                    Throw New Exception("WipData SourceTable Is Null！--AutoSerialNo = " & rw.Item("AutoSerialNo"))
                                End If
                                fFieldsAndValues.Add(CmdStatusField,
                                                                  WipData.Tables(rw.Item("SourceTable").ToString).Rows(0).Item(rw.Item("SourceField").ToString))

                            End If
                        End If
                    Case 2
                        ErrorCodeField = rw.Item("FieldName")
                    Case 3
                        ErrorMsgField = rw.Item("FieldName")

                    Case Else
                        Throw New Exception(String.Format(Lang.SetFieldTypeError, rw.Item("AutoSerialNo")))

                End Select
                If intType = 4 Then
                    '2012/09/14 Jacky 加參考號4 為Query Key
                    fSeqNoFieldName = rw.Item("FieldName")
                    fSeqNo = fFieldsAndValues(rw.Item("FieldName"))
                End If
               
            Next
            Dim aFields As String = Nothing
            Dim aValues As String = Nothing
            If String.IsNullOrEmpty(ErrorMsgField) Then
                ErrorMsgField = ErrorCodeField
            End If
            For i As Int32 = 0 To fFieldsAndValues.Count - 1
                If String.IsNullOrEmpty(aFields) Then
                    aFields = fFieldsAndValues.Keys(i)
                Else
                    aFields = aFields & "," & fFieldsAndValues.Keys(i)
                End If
                If String.IsNullOrEmpty(aValues) Then
                    aValues = "{0}0"
                Else
                    aValues = aValues & ",{0}" & i
                End If
            Next
            'aRet = String.Format("INSERT INTO {0} ( {1} ) VALUES ({2})", TakeTableName, aFields, aValues)
            'aRet = String.Format(aRet, _DAL.Sign)
            aRet = _DAL.GetMasterInsertSQL(TakeTableName, aFields, aValues)
        Catch ex As Exception
            Throw New Exception(ex.Message & "--TakeInsertSQL")
        Finally
            tbSchema.Dispose()
        End Try
        Return aRet
    End Function
    Private Function GetSeqNo(ByVal SourceField As String) As Object
        'Dim aSQL As String = "SELECT " & OwnerName & SourceField & ".NEXTVAL FROM DUAL"
        Dim aSQL As String = _DAL.GetSeqNo(OwnerName, SourceField)
        fSeqNo = daoCmd.ExecSclr(aSQL)
        Return fSeqNo
    End Function
    Private Function GetLoginInfoValue(ByVal PropertyName As String) As Object
        'Dim li As New CableSoft.BLL.Utility.LoginInfo()
        'Dim tp As Type = li.GetType()

        'Dim propInfo As PropertyInfo

        For Each pi As PropertyInfo In Me.LoginInfo.GetType().GetProperties
            If pi.Name.ToUpper = PropertyName.ToUpper Then
                Return pi.GetValue(Me.LoginInfo, Nothing)
            End If
        Next

        Return Nothing
    End Function
    Private Function TakeTable(ByVal TableName As String, ByVal CMDID As String) As Boolean
        Dim aSeqNo As Int32 = -1
        If String.IsNullOrEmpty(TableName) Then
            Throw New Exception(Lang.TableNameIsNull)
            Return False
        End If
        If String.IsNullOrEmpty(CMDID) Then
            Throw New Exception(Lang.CmdIdIsNull)
            Return False
        End If
        tbMaster = daoCmd.ExecQry(_DAL.TakeMasterTable, New Object() {TableName.ToUpper, CMDID.ToUpper})
        If tbMaster.Rows.Count = 0 Then
            Throw New Exception(String.Format(Lang.NotFoundMasterCMD, CMDID))
            Return False
        Else
            If tbMaster.Rows.Count > 1 Then
                Throw New Exception(String.Format(Lang.MasterCmdDouble, CMDID))
                Return False
            Else
                If DBNull.Value.Equals(tbMaster.Rows(0).Item("OwnerName")) Then
                    Throw New Exception(String.Format(Lang.MasterNoOwner, CMDID))
                    Return False
                Else
                    OwnerName = tbMaster.Rows(0).Item("OwnerName")
                End If
                If DBNull.Value.Equals(tbMaster.Rows(0).Item("SuccessCodeStr")) Then
                    Throw New Exception(String.Format(Lang.MasterCmdNoSucessCode, CMDID))
                    Return False
                End If
                If DBNull.Value.Equals(tbMaster.Rows(0).Item("ErrorCodeStr")) Then
                    Throw New Exception(String.Format(Lang.MasterCmdNoErrCode, CMDID))
                    Return False
                End If
                If Not DBNull.Value.Equals(tbMaster.Rows(0).Item("TimeOut")) Then
                    WaitTimeOut = Int32.Parse(tbMaster.Rows(0).Item("TimeOut"))
                End If
                SuccessCodeStr = tbMaster.Rows(0).Item("SuccessCodeStr")
                ErrorCodeStr = tbMaster.Rows(0).Item("ErrorCodeStr")
                aSeqNo = Int32.Parse(tbMaster.Rows(0).Item("SEQNO"))
            End If
        End If
        tbDetail = daoCmd.ExecQry(_DAL.TakeDetailTable, New Object() {aSeqNo})
        If tbDetail.Rows.Count = 0 Then
            Throw New Exception(Lang.DetailNoData)
            Return False
        End If

        Dim lstRw As List(Of DataRow) = tbDetail.AsEnumerable.Where(Function(rw As DataRow)
                                                                        If rw.IsNull("FieldName") Then
                                                                            Return True
                                                                        End If
                                                                        '如果DefaultValue 有值，直接填值，不用檢查
                                                                        If Not DBNull.Value.Equals(rw.Item(DefaultField)) Then
                                                                            Return False
                                                                        End If
                                                                        If (rw.IsNull("SourceTable")) OrElse (rw.IsNull("SourceField")) Then
                                                                            If (DBNull.Value.Equals(rw.Item("FieldType"))) OrElse
                                                                                (Int32.Parse(rw.Item("FieldType")) = 0) OrElse
                                                                                (Int32.Parse(rw.Item("FieldType")) = 4) Then
                                                                                Return True
                                                                            Else
                                                                                Return False
                                                                            End If
                                                                        End If
                                                                        Return False
                                                                    End Function).ToList

        If (lstRw IsNot Nothing) AndAlso (lstRw.Count > 0) Then
            Throw New Exception(String.Format(Lang.DetailSetError, lstRw.Item(0).Item("AutoSerialNo")))
            Return False
        End If
        For i As Int32 = 1 To 2
            If tbDetail.AsEnumerable.Count(Function(rw As DataRow)
                                               If (Not DBNull.Value.Equals(rw.Item("FieldType"))) AndAlso (Int32.Parse(rw.Item("FieldType")) = i) Then
                                                   Return True
                                               End If
                                               Return False
                                           End Function) = 0 Then
                Throw New Exception(String.Format(Lang.DetailNotSetType, i.ToString))
                Return False
            End If

        Next




        Return True
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                
                If tbDetail IsNot Nothing Then
                    tbDetail.Dispose()
                End If
                If tbMaster IsNot Nothing Then
                    tbMaster.Dispose()
                End If
                If fFieldsAndValues IsNot Nothing Then
                    fFieldsAndValues.Clear()
                End If
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                End If
                If daoCmd IsNot Nothing Then
                    daoCmd.Dispose()
                End If
                If Lang IsNot Nothing Then
                    Lang.Dispose()
                    Lang = Nothing
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
