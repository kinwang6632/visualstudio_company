Imports CableSoft.BLL.Utility
Imports System.Data.Common
Imports System.Reflection

Public Class DynUpdateGrid2
    Inherits BLLBasic
    Implements IDisposable
    Private dynaCdt As CableSoft.BLL.Dynamic.Condition.DynamicCondition = Nothing
    Private _DAL As New DynUpdateGrid2DAL(Me.LoginInfo.Provider)
    Private lang As New CableSoft.BLL.Language.SO61.DynUpdateGrid2Language()
    Private tbMaster As DataTable = Nothing
    Private tbDetail As DataTable = Nothing
    Private Const tbMasterName As String = "Master"
    Private Const tbDetailName As String = "Detail"
    Private Const LoginInfoString As String = "LoginInfo"
    Private Const SeqNoString As String = "SEQNO"
    Private tbSechema As DataTable = Nothing
    Private fFieldsAndValues As Dictionary(Of String, Object)
    Private fWhereFieldsAndValues As Dictionary(Of String, Object)
    Private fUKWhereFieldAndValues As Dictionary(Of String, Object)
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
    Private Sub FillSourceField()
        Try
            tbDetail.Columns.Add(New DataColumn("GetDesc", GetType(Boolean)))
            For Each rw As DataRow In tbDetail.Rows
                rw.BeginEdit()
                rw.Item("GetDesc") = False
                If Not DBNull.Value.Equals(rw.Item("SourceField")) Then
                    If rw.Item("SourceField").ToString.Length > "_DESC".Length Then
                        If rw.Item("SourceField").ToString.Substring(rw.Item("SourceField").ToString.Length - "_DESC".Length, "_DESC".Length).ToUpper =
                      "_DESC".ToUpper Then
                            rw.Item("GetDesc") = True
                            rw.Item("SourceField") =
                                rw.Item("SourceField").ToString.Substring(0, (rw.Item("SourceField").ToString.Length - "_DESC".Length))
                        End If
                    End If

                    Dim strSourceField As String = rw.Item("SourceField").ToString.Substring(rw.Item("SourceField").ToString.Length - 2, 2)

                    If (strSourceField <> "_0") AndAlso (strSourceField <> "_1") AndAlso (strSourceField <> "_2") Then
                        If (rw.Item("SourceTable").ToString.ToUpper <> LoginInfoString.ToUpper) AndAlso
                            (rw.Item("SourceTable").ToString.ToUpper <> SeqNoString.ToUpper) Then
                            'rw.Item("SourceField") = rw.Item("SourceField").ToString & "_0"
                        End If
                    End If
                End If
                rw.EndEdit()
            Next
            tbDetail.AcceptChanges()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function GetSechema() As DataTable
        Return DAO.ExecQry(_DAL.QuerySchema(tbMaster.Rows(0).Item("TableName"))).Copy
    End Function
    Public Function ConverDataType(ByVal FieldName As String, ByVal tbSchema As DataTable, ByVal Value As Object) As Object
        Dim aRet As Object = Nothing
        Try
            If DBNull.Value.Equals(Value) Then
                Return DBNull.Value
            End If
            If Value.GetType.Equals(tbSchema.Columns(FieldName).DataType) Then
                Return Value
            End If
            If DBNull.Value.Equals(Value) OrElse String.IsNullOrEmpty(Value.ToString) Then
                aRet = DBNull.Value
            Else
                aRet = Convert.ChangeType(Value, tbSchema.Columns(FieldName).DataType)
            End If

        Catch ex As Exception
            Throw ex
        End Try
        Return aRet
    End Function
    Private Function ReplaceLoginInfoWhere(ByVal UpdSQL As String, ByRef params() As Object) As String
        Dim RetSQL As String = String.Empty
        Try
            UpdSQL = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, UpdSQL, "LoginInfo.EntryName", Me.LoginInfo.EntryName, params)
            UpdSQL = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, UpdSQL, "LoginInfo.EntryId", Me.LoginInfo.EntryId, params)
            UpdSQL = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, UpdSQL, "LoginInfo.CompCode", Me.LoginInfo.CompCode, params)
            RetSQL = UpdSQL
        Catch ex As Exception
            Throw ex
        End Try
        Return RetSQL
    End Function
    Private Function GetFinalValue(ByVal strSource As String, ByVal dsSource As DataSet,
                                   ByRef ExeDao As CableSoft.Utility.DataAccess.DAO) As Object
        Dim dtDynReturn As DataTable = Nothing
        Try
            If (strSource.Length > 6) AndAlso (strSource.Trim.ToString.ToUpper.Substring(0, 6) = "SELECT".ToUpper) Then
                Dim params() As Object = Nothing
                If dynaCdt IsNot Nothing Then
                    dynaCdt.Dispose()
                    dynaCdt = Nothing

                End If
                dynaCdt = New CableSoft.BLL.Dynamic.Condition.DynamicCondition(Me.LoginInfo, ExeDao)
                dtDynReturn = dynaCdt.GetBuildConditionSQL(tbMaster.Rows(0).Item("SysProgramId"),
                                                           dsSource.Tables("Condition"), params)

                Dim aFieldName As String = Nothing

                For Each dr As DataRow In dtDynReturn.Rows
                    strSource = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                             strSource, dr("FieldName"), dr("ConditionSQL"), params)

                Next
                For i As Integer = 0 To dsSource.Tables("Condition").Rows.Count - 1
                    If Right(dsSource.Tables("Condition").Rows(i)("FieldName"), 2) = "_0" Then
                        aFieldName = dsSource.Tables("Condition").Rows(i)("FieldName").ToString.Substring(0,
                                                                                   dsSource.Tables("Condition").Rows(i)("FieldName").ToString.Length - 2)
                    Else
                        aFieldName = dsSource.Tables("Condition").Rows(i)("FieldName").ToString
                    End If
                    strSource = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                            strSource, aFieldName, dsSource.Tables("Condition").Rows(i)("FieldValue"), params)
                    strSource = ReplaceLoginInfoWhere(strSource, params)

                Next
                Try
                    If (params Is Nothing) OrElse (params.Count = 0) Then
                        Return ExeDao.ExecSclr(strSource)
                    Else
                        Return ExeDao.ExecSclr(strSource, params)
                    End If

                Catch ex As Exception
                    Return strSource
                End Try
            Else
                Return strSource
            End If
        Finally
            If dtDynReturn IsNot Nothing Then
                dtDynReturn.Dispose()
                dtDynReturn = Nothing
            End If
        End Try

    End Function
    Private Function GetLoginInfoValue(ByVal PropertyName As String, ByRef exeLoginInfo As LoginInfo) As Object

        For Each pi As PropertyInfo In exeLoginInfo.GetType().GetProperties
            If pi.Name.ToUpper = PropertyName.ToUpper Then
                Return pi.GetValue(exeLoginInfo, Nothing)
            End If
        Next

        Return Nothing
    End Function
    Private Function GetSeqNo(ByVal SourceField As String, ByRef exeDao As CableSoft.Utility.DataAccess.DAO) As Object
        'Dim aSQL As String = "SELECT " & SourceField & ".NEXTVAL FROM DUAL"
        Dim aSQL As String = _DAL.getSEQNo(SourceField)
        Try
            Return exeDao.ExecSclr(aSQL)
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Private Function GetDefaultToFieldsAndValues(ByVal EditMode As EditMode,
                                                 ByVal dsSource As DataSet,
                                                 ByRef exeDao As CableSoft.Utility.DataAccess.DAO,
                                                 ByRef exeLoginInfo As LoginInfo) As Boolean
        Try
            '取出預設值
            Dim lstRw As List(Of DataRow) = tbDetail.AsEnumerable.Where(Function(rw As DataRow)
                                                                            If (EditMode <> Utility.EditMode.Append) AndAlso rw.Item("FieldType") = 1 Then
                                                                                Return False
                                                                            End If

                                                                            If Not DBNull.Value.Equals(rw.Item("FinalValue")) Then
                                                                                Return True
                                                                            End If
                                                                            If rw.Item("SourceTable").ToString.ToUpper = LoginInfoString.ToUpper OrElse
                                                                                rw.Item("SourceTable").ToString.ToUpper = SeqNoString.ToUpper Then
                                                                                Return True
                                                                            End If

                                                                            Return False
                                                                        End Function).ToList

            If lstRw IsNot Nothing AndAlso lstRw.Count > 0 Then
                For Each rwDetail As DataRow In lstRw
                    If Not DBNull.Value.Equals(rwDetail.Item("FinalValue")) Then
                        If Not fFieldsAndValues.ContainsKey(rwDetail("FieldName").ToString) Then
                            fFieldsAndValues.Add(rwDetail("FieldName").ToString,
                                              ConverDataType(rwDetail("FieldName").ToString, tbSechema,
                                                             GetFinalValue(rwDetail.Item("FinalValue").ToString, dsSource, exeDao)))
                        End If

                    Else
                        Select Case rwDetail.Item("SourceTable").ToString.ToUpper
                            Case LoginInfoString.ToUpper
                                Dim aProperty As Object = Nothing
                                aProperty = GetLoginInfoValue(rwDetail.Item("SourceField"), exeLoginInfo)
                                If aProperty IsNot Nothing Then
                                    If Not fFieldsAndValues.ContainsKey(rwDetail("FieldName").ToString) Then
                                        fFieldsAndValues.Add(rwDetail("FieldName").ToString,
                                                        ConverDataType(rwDetail("FieldName").ToString, tbSechema, aProperty))
                                    End If
                                End If
                            Case SeqNoString.ToUpper
                                If Not fFieldsAndValues.ContainsKey(rwDetail("FieldName").ToString) Then
                                    fFieldsAndValues.Add(rwDetail("FieldName").ToString, ConverDataType(rwDetail("FieldName").ToString,
                                                                                                    tbSechema,
                                                                                                    GetSeqNo(rwDetail.Item("SourceField"), Me.DAO)))
                                End If

                            Case Else
                                If Not fFieldsAndValues.ContainsKey(rwDetail("FieldName").ToString) Then
                                    fFieldsAndValues.Add(rwDetail("FieldName").ToString,
                                                    ConverDataType(rwDetail("FieldName").ToString, tbSechema,
                                                                   GetFinalValue(rwDetail.Item("FinalValue").ToString, dsSource, exeDao)))
                                End If

                        End Select
                    End If

                Next

            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return True
    End Function
    Private Function GetInsertSQL(ByVal dsSource As DataSet,
                                  ByRef exeDao As CableSoft.Utility.DataAccess.DAO,
                                  ByRef exeLoginInfo As LoginInfo) As String
        Dim aRet As String = Nothing
        Dim aFieldValueName As String = "FieldValue"
        Try

            For Each rwSource As DataRow In dsSource.Tables("Condition").Rows
                For Each rwDetail As DataRow In tbDetail.Rows
                    aFieldValueName = "FieldValue"
                    If rwDetail.Item("GetDesc") Then
                        aFieldValueName = "FieldDesc"
                    End If

                    If (rwDetail.Item("SourceField").ToString.ToUpper = rwSource("FieldName").ToString.ToUpper) Then
                        If Not DBNull.Value.Equals(rwSource(aFieldValueName)) Then
                            If fFieldsAndValues Is Nothing Then fFieldsAndValues = New Dictionary(Of String, Object)
                            fFieldsAndValues.Add(rwDetail.Item("FieldName"),
                                             ConverDataType(rwDetail.Item("FieldName"), tbSechema, rwSource(aFieldValueName)))
                        End If
                    End If
                Next
            Next

            GetDefaultToFieldsAndValues(EditMode.Append, dsSource, exeDao, exeLoginInfo)
            aRet = _DAL.GetInsertSQL(tbMaster.Rows(0).Item("TableName"), fFieldsAndValues)
        Catch ex As Exception
            Throw ex
        End Try
        Return aRet
    End Function
    'Private Function TakeSQL(ByVal EditMode As EditMode, ByVal dsSource As DataSet) As String
    '    Dim aRet As String = Nothing

    '    Dim lstPkRw As List(Of DataRow) = Nothing
    '    lstPkRw = GetPKRow(dsSource)

    '    If lstPkRw Is Nothing OrElse lstPkRw.Count = 0 Then
    '        Throw New Exception(Language.NotFoundPKValue)
    '    End If
    '    If fFieldsAndValues Is Nothing Then
    '        fFieldsAndValues = New Dictionary(Of String, Object)
    '    End If

    '    If fWhereFieldsAndValues Is Nothing Then
    '        fWhereFieldsAndValues = New Dictionary(Of String, Object)
    '    End If
    '    fWhereFieldsAndValues.Clear()
    '    fFieldsAndValues.Clear()
    '    GetWhereList(EditMode, dsSource, lstPkRw)
    '    If EditMode <> Utility.EditMode.Append Then
    '        If fWhereFieldsAndValues Is Nothing OrElse fWhereFieldsAndValues.Count = 0 Then
    '            Throw New Exception(String.Format(Language.GetNoWhere, "TakeSQL"))
    '        End If
    '    End If
    '    Select Case EditMode
    '        Case Utility.EditMode.Append
    '            aRet = GetInsertSQL(dsSource)
    '        Case Utility.EditMode.Edit
    '            aRet = GetUpdateSQL(dsSource)
    '        Case Utility.EditMode.Delete
    '            aRet = GetDelSQL(dsSource)
    '        Case Else
    '            aRet = GetUpdateSQL(dsSource)
    '    End Select
    '    'If EditMode <> Utility.EditMode.Append Then
    '    '    aRet = GetUpdateSQL(dsSource)
    '    'Else
    '    '    aRet = GetInsertSQL(dsSource)
    '    'End If

    '    Return aRet
    'End Function
    Public Function Execute(ByVal EditMode As EditMode, ByVal SysProgramId As String,
                            ByVal dsSource As DataSet) As RIAResult
        Dim result As New RIAResult
        Dim dsReturn As New DataSet()
        Dim ErrMsg As String = Nothing


        Dim UpdateSQL As New List(Of String)
        Dim UpdateParams As New List(Of Object())
        Dim BeforeSQL As New List(Of String)
        Dim BeforeParams As New List(Of Object())
        Dim InserChildSQL As New List(Of String)
        Dim dtDynReturn As DataTable = Nothing

        Dim tbOriginal As DataTable = Nothing
        Dim tbUpdate As DataTable = Nothing
        Dim lstChildValue As New Dictionary(Of Integer, List(Of Object))
        'Dim aInsertSQL As String = String.Empty
        'Dim aChildInsertSQL As New List(Of Array)

        Dim PKErrMsg As String = Nothing

        Dim resultTable As DataTable = Nothing


        dynaCdt = New CableSoft.BLL.Dynamic.Condition.DynamicCondition(Me.LoginInfo, Me.DAO)

        result.ResultBoolean = False
        Try
            tbMaster = DAO.ExecQry(_DAL.QuerySO1109A, New Object() {SysProgramId})
            tbMaster.TableName = tbMasterName
            tbDetail = DAO.ExecQry(_DAL.QurerySO1109B, New Object() {tbMaster.Rows(0).Item("ProgramId")})
            tbDetail.TableName = tbDetailName

            Dim params() As Object = Nothing
            dtDynReturn = dynaCdt.GetBuildConditionSQL(tbMaster.Rows(0).Item("SysProgramId"),
                                                       dsSource.Tables("Condition"), params)

            Dim aFieldName As String = Nothing


            ErrMsg = chkSchema()
            Dim indexOrder As Integer = 0
            For Each rw As DataRow In dsSource.Tables("Condition").Rows
                rw.Item("FieldName") = Replace(rw("FieldName"), "_1", "")
                rw.Item("FieldName") = Replace(rw("FieldName"), "_0", "")
            Next
            FillSourceField()
            If Not String.IsNullOrEmpty(ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            tbSechema = GetSechema()
            'Dim aSQL As String = TakeSQL(EditMode, dsSource)
            Dim aSQL As String = GetInsertSQL(dsSource, DAO, LoginInfo)
            resultTable = DAO.ExecQry(aSQL)
            Dim rwNew As DataRow = resultTable.NewRow
            For i As Integer = 0 To fFieldsAndValues.Count - 1
                rwNew(fFieldsAndValues.Keys(i)) = fFieldsAndValues.Values(i)
            Next
            resultTable.Rows.Add(rwNew)
            dsReturn.Tables.Add(resultTable.Copy)
            result.ResultBoolean = True
            result.ErrorCode = 0
            result.ErrorMessage = Nothing
            result.ResultDataSet = dsReturn


        Catch ex As Exception
            result.ResultBoolean = False
            result.ErrorMessage = ex.ToString
            result.ErrorCode = -3
            'Throw ex
        Finally
            If resultTable IsNot Nothing Then
                resultTable.Dispose()
                resultTable = Nothing
            End If
            If dsReturn IsNot Nothing Then
                dsReturn.Dispose()
                dsReturn = Nothing
            End If

            If fFieldsAndValues IsNot Nothing Then
                fFieldsAndValues.Clear()
                fFieldsAndValues = Nothing
            End If
            If fWhereFieldsAndValues IsNot Nothing Then
                fWhereFieldsAndValues.Clear()
                fWhereFieldsAndValues = Nothing
            End If
            If fUKWhereFieldAndValues IsNot Nothing Then
                fUKWhereFieldAndValues.Clear()
                fUKWhereFieldAndValues = Nothing
            End If

            If tbMaster IsNot Nothing Then
                tbMaster.Dispose()
                tbMaster = Nothing
            End If
            If tbDetail IsNot Nothing Then
                tbDetail.Dispose()
                tbDetail = Nothing
            End If
            If tbOriginal IsNot Nothing Then
                tbOriginal.Dispose()
                tbOriginal = Nothing
            End If
            If tbSechema IsNot Nothing Then
                tbSechema.Dispose()
                tbSechema = Nothing
            End If

            If tbUpdate IsNot Nothing Then
                tbUpdate.Dispose()
                tbUpdate = Nothing
            End If
            If dsSource IsNot Nothing Then
                dsSource.Dispose()
                dsSource = Nothing
            End If

            If dynaCdt IsNot Nothing Then
                dynaCdt.Dispose()
                dynaCdt = Nothing
            End If
            If dtDynReturn IsNot Nothing Then
                dtDynReturn.Dispose()
                dtDynReturn = Nothing
            End If
            If UpdateSQL IsNot Nothing Then
                UpdateSQL.Clear()
                UpdateSQL = Nothing
            End If
            If tbOriginal IsNot Nothing Then
                tbOriginal.Dispose()
                tbOriginal = Nothing
            End If

        End Try
        Return result
    End Function
    Private Function chkSchema() As String
        Using tbSchema As DataTable = DAO.ExecQry(_DAL.QuerySchema(tbMaster.Rows(0).Item("TableName")))
            If tbSchema Is Nothing Then
                Return String.Format(lang.NotFoundTable, tbMaster.Rows(0).Item("TableName"))
            End If
            For Each rw As DataRow In tbDetail.Rows
                If Not tbSchema.Columns.Contains(rw.Item("FieldName")) Then
                    Return String.Format(lang.NotFoundField,
                                         tbMaster.Rows(0).Item("TableName"),
                                         rw.Item("FieldName"), rw.Item("AutoSerialNo"))
                End If
            Next
        End Using
        Return Nothing
    End Function
    Public Function ChkAuthority(ByVal SysProgramId As String) As RIAResult
        Dim result As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        Dim tbMaster As DataTable = DAO.ExecQry(_DAL.QueryMaster, New Object() {SysProgramId})
        Try
            If tbMaster.Rows.Count = 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -3
                result.ErrorMessage = lang.NoFundMaster
                Return result
            End If
            If Me.LoginInfo.GroupId = "0" AndAlso 1 = 0 Then
                Return result
            Else
                Using obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
                    result = obj.ChkPriv(LoginInfo.EntryId, tbMaster.Rows(0).Item("MID"))
                    obj.Dispose()
                End Using
                If Not result.ResultBoolean Then
                    result.ErrorCode = -1
                    result.ErrorMessage = lang.NoPermission
                End If
                'If Integer.Parse(DAO.ExecSclr(_DAL.chkAuthority(Me.LoginInfo.GroupId), New Object() {tbMaster.Rows(0).Item("MID")})) = 0 Then
                '    result.ResultBoolean = False
                '    result.ErrorCode = -1
                '    result.ErrorMessage = lang.NoPermission
                '    Return result
                'End If
            End If

        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -2
        Finally
            If tbMaster IsNot Nothing Then
                tbMaster.Dispose()
                tbMaster = Nothing
            End If
        End Try
        Return result

    End Function
    Public Function QueryDynUpdateGrid(ByVal SysProgramId As String) As DataSet
        Dim ds As New DataSet
        Dim tbMaster As DataTable = DAO.ExecQry(_DAL.QueryMaster, New Object() {SysProgramId})
        Dim aProgramId As String = "X"
        If (tbMaster IsNot Nothing) AndAlso (tbMaster.Rows.Count > 0) Then
            aProgramId = tbMaster.Rows(0).Item("ProgramId")
        End If
        Dim tbDetail As DataTable = DAO.ExecQry(_DAL.QueryDetail, New Object() {aProgramId})
        Dim tbSO1109B As DataTable = DAO.ExecQry(_DAL.QurerySO1109B, New Object() {aProgramId})
        Try
            tbMaster.TableName = "Master"
            tbDetail.TableName = "Detail"
            tbSO1109B.TableName = "SO1109B"
            tbMaster.Columns.Add("CanAppend", GetType(Boolean))
            tbMaster.Columns.Add("CanEdit", GetType(Boolean))
            tbMaster.Columns.Add("CanDelete", GetType(Boolean))
            tbMaster.Columns.Add("CanCopyOtherDB", GetType(Boolean))
            For Each rw As DataRow In tbMaster.Rows
                rw.BeginEdit()
                If DBNull.Value.Equals(rw.Item("AppendMID")) Then
                    rw.Item("CanAppend") = False
                Else
                    rw.Item("CanAppend") = ChkPriv(rw("AppendMID")).ResultBoolean
                End If
                If DBNull.Value.Equals(rw.Item("EditMID")) Then
                    rw.Item("CanEdit") = False
                Else
                    rw.Item("CanEdit") = ChkPriv(rw.Item("EditMID")).ResultBoolean
                End If

                If DBNull.Value.Equals(rw.Item("DeleteMID")) Then
                    rw.Item("CanDelete") = False
                Else
                    rw.Item("CanDelete") = ChkPriv(rw.Item("DeleteMID")).ResultBoolean
                End If
                rw.EndEdit()
            Next
            tbMaster.AcceptChanges()
            ds.Tables.Add(tbMaster.Copy)
            ds.Tables.Add(tbDetail.Copy)
            ds.Tables.Add(tbSO1109B.Copy)
        Catch ex As Exception
            Throw ex
        End Try
        Return ds
    End Function
    Public Function ChkPriv(ByVal PrivMid As String) As RIAResult

        Using objPriv As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
            Return objPriv.ChkPriv(Me.LoginInfo.EntryId, PrivMid)
        End Using
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, PrivMid)
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
