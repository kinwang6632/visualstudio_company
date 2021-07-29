Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports System.Reflection
Imports System.Windows.Forms
Imports CableSoft.SO.BLL.DataLog

Public Class DynamicUpdate
    Inherits BLLBasic
    Implements IDisposable


    Private _DAL As New DynamicUpdateDALMultiDB(Me.LoginInfo.Provider)

    Private tbMaster As DataTable = Nothing
    Private tbDetail As DataTable = Nothing
    Private tbChildSO1109B As DataTable = Nothing
    Private tbChildSO1109A As DataTable = Nothing
    Private Const DefaultField As String = "FinalValue"
    Private Const tbMasterName As String = "Master"
    Private Const tbDetailName As String = "Detail"
    Private Const LoginInfoString As String = "LoginInfo"
    Private Const SeqNoString As String = "SEQNO"
    Private fFieldsAndValues As Dictionary(Of String, Object)
    Private fWhereFieldsAndValues As Dictionary(Of String, Object)
    Private fUKWhereFieldAndValues As Dictionary(Of String, Object)
    Private tbSechema As DataTable = Nothing
    Private Language As New CableSoft.BLL.Language.SO61.DynamicUpdateLanguage
    Private dynaCdt As CableSoft.BLL.Dynamic.Condition.DynamicCondition = Nothing
    Private tbChildSchema As DataTable = Nothing
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
    Public Function Save(ByVal EditMode As EditMode, ByVal SysProgramId As String, ByVal dsSource As DataSet) As RIAResult
        Return Execute(EditMode, SysProgramId, dsSource)
    End Function
    Private Function GetSechema(ByVal aTableName As String) As DataTable
        Return DAO.ExecQry(_DAL.QuerySchema(aTableName)).Copy
    End Function
    Private Function GetSechema() As DataTable
        Return DAO.ExecQry(_DAL.QuerySchema(tbMaster.Rows(0).Item("TableName"))).Copy
    End Function
    '取得可選公司別
    Public Function GetCompCode() As DataTable
        Try
            If Me.LoginInfo.GroupId = "0" Then
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
    Public Function CopyToOtherDB(ByVal sysProgramId As String, ByVal IsCover As Boolean, ByVal dsSource As DataSet, ByVal dsCopyComps As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim ErrMsg As String = Nothing
        Dim result As New RIAResult
        Dim strRetMsg As New System.Text.StringBuilder()
        Dim dtDynReturn As DataTable = Nothing
        Dim BeforeSQL As New List(Of String)
        'Dim aChildInsertSQL As String = String.Empty
        Dim aChildInsertSQL As New ChildSQL

        'Dim lstChildValue As New Dictionary(Of Integer, List(Of Object))
        Dim dynaCdt = New CableSoft.BLL.Dynamic.Condition.DynamicCondition(Me.LoginInfo, Me.DAO)
        result.ErrorCode = 0
        result.ResultBoolean = True
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

        Try
            tbMaster = DAO.ExecQry(_DAL.QuerySO1109A, New Object() {sysProgramId})
            tbMaster.TableName = tbMasterName
            tbDetail = DAO.ExecQry(_DAL.QuerySO1109B, New Object() {tbMaster.Rows(0).Item("ProgramId")})
            tbDetail.TableName = tbDetailName
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryName, _
                                                            String.Format(Language.CopyClientInfo, tbMaster.Rows(0).Item("Caption")))
            End If
            If Not DBNull.Value.Equals(tbMaster.Rows(0).Item("BeforeSQL")) Then
                BeforeSQL = tbMaster.Rows(0).Item("BeforeSQL").ToString.Split(";").ToList
            End If
            Dim params() As Object = Nothing
            dtDynReturn = dynaCdt.GetBuildConditionSQL(tbMaster.Rows(0).Item("SysProgramId"),
                                                       dsSource.Tables("Condition"), params)
            '取出updateSQL By Kin 2015/06/23
            For Each dr As DataRow In dtDynReturn.Rows               
                For index As Integer = 0 To BeforeSQL.Count - 1
                    BeforeSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         BeforeSQL.Item(index), dr("FieldName"), dr("ConditionSQL"), params)
                Next
            Next
            ErrMsg = chkSchema()
            FillSourceField()
            If Not String.IsNullOrEmpty(ErrMsg) Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = ErrMsg
                Return result
            End If
            tbSechema = GetSechema()
            Dim lstPkRw As List(Of DataRow) = Nothing
            

            'For Each tb As DataTable In dsSource.Tables
            '    For Each rw As DataRow In tb.Rows
            '        rw.Item("FieldName") = Replace(rw("FieldName"), "_1", "")
            '        rw.Item("FieldName") = Replace(rw("FieldName"), "_0", "")
            '    Next
            'Next

            'lstPkRw = GetCopyPKRow() --如果SO1111A有設定複制的PK欄位 要使用這個Function
            If fWhereFieldsAndValues Is Nothing Then
                fWhereFieldsAndValues = New Dictionary(Of String, Object)
            End If




            For Each rw As DataRow In dsCopyComps.Tables(0).Rows
                'lstChildValue.Clear()                
                If fFieldsAndValues Is Nothing Then
                    fFieldsAndValues = New Dictionary(Of String, Object)
                End If
                fFieldsAndValues.Clear()
                Dim otherLoginInfo As LoginInfo = CableSoft.BLL.Utility.Utility.GetRealLoginInfo(LoginInfo, rw.Item("CompId"))
                Dim otherDao As New CableSoft.Utility.DataAccess.DAO(otherLoginInfo.Provider, otherLoginInfo.ConnectionString)               
                For Each rwSource As DataRow In dsSource.Tables("Condition").Rows
                    If Integer.Parse(rwSource("OBJECTTYPE") & "") <> 30 Then
                        rwSource.Item("FieldName") = Replace(rwSource("FieldName"), "_1", "")
                        rwSource.Item("FieldName") = Replace(rwSource("FieldName"), "_0", "")
                    Else
                        rwSource.Item("FieldName") = Replace(rwSource("FieldName"), "_1", "")
                        rwSource.Item("FieldName") = Replace(rwSource("FieldName"), "_0", "")
                        tbChildSO1109B = DAO.ExecQry(_DAL.QuerySO1109B, New Object() {rwSource.Item("FieldName")})

                        Dim aChildSQL As String = getChildInserSQL(rwSource.Item("FIELDVALUE").ToString, tbChildSO1109B)
                        'aChildInsertSQL = getChildInserSQL(rwSource.Item("FIELDVALUE").ToString, tbChildSO1109B)

                        tbChildSchema = GetSechema(rwSource.Item("FIELDVALUE").ToString)

                        Try
                            aChildInsertSQL.Clear()

                            For index As Integer = 0 To dsSource.Tables(rwSource.Item("FIELDVALUE").ToString).Rows.Count - 1
                                'lstChildValue.Add(index, GetChildValue(tbChildSO1109B,
                                '                                       dsSource.Tables(rwSource.Item("FIELDVALUE").ToString).Rows(index),
                                '                                       dsSource, otherDao, otherLoginInfo))
                                aChildInsertSQL.setSQL(aChildSQL, GetChildValue(tbChildSO1109B,
                                                                 dsSource.Tables(rwSource.Item("FIELDVALUE").ToString).Rows(index),
                                                                 dsSource, otherDao, otherLoginInfo))
                            Next

                        Finally

                        End Try
                    End If
                Next
                lstPkRw = GetPKRow(dsSource)
                fWhereFieldsAndValues.Clear()
                GetWhereList(EditMode.View, dsSource, otherLoginInfo, lstPkRw)
                Dim aFindSQL As String = GetFindSQL()
                Dim aDelSQL As String = GetDelSQL(dsSource)
                Dim aInsAllDataSQL As String = GetInsertSQL(dsSource, otherDao, otherLoginInfo)
                Try
                    If Not IsCover Then
                        If Integer.Parse(otherDao.ExecSclr(aFindSQL, fWhereFieldsAndValues.Values.ToArray)) = 0 Then
                            otherDao.ExecNqry(aInsAllDataSQL, fFieldsAndValues.Values.ToArray)
                            For Each befSQL As String In BeforeSQL
                                otherDao.ExecNqry(befSQL)
                            Next
                            For index As Integer = 0 To aChildInsertSQL.TotalCount - 1
                                otherDao.ExecNqry(aChildInsertSQL.readInsSQL(index), aChildInsertSQL.readValues(index))
                            Next
                            'If Not String.IsNullOrEmpty(aChildInsertSQL) Then
                            '    For index As Integer = 0 To lstChildValue.Count - 1
                            '        otherDao.ExecNqry(aChildInsertSQL, lstChildValue.Item(index).ToArray)
                            '    Next
                            'End If
                            strRetMsg.AppendLine(String.Format(Language.CopyOK, otherLoginInfo.CompCode))
                        Else
                            strRetMsg.AppendLine(String.Format(Language.CodeExists, otherLoginInfo.CompCode))
                        End If
                    Else
                        otherDao.ExecNqry(aDelSQL, fWhereFieldsAndValues.Values.ToArray)
                        otherDao.ExecNqry(aInsAllDataSQL, fFieldsAndValues.Values.ToArray)
                        For Each befSQL As String In BeforeSQL
                            otherDao.ExecNqry(befSQL)
                        Next
                        For index As Integer = 0 To aChildInsertSQL.TotalCount - 1
                            otherDao.ExecNqry(aChildInsertSQL.readInsSQL(index), aChildInsertSQL.readValues(index))
                        Next
                        'If Not String.IsNullOrEmpty(aChildInsertSQL) Then
                        '    For index As Integer = 0 To lstChildValue.Count - 1
                        '        otherDao.ExecNqry(aChildInsertSQL, lstChildValue.Item(index).ToArray)
                        '    Next
                        'End If
                        strRetMsg.AppendLine(String.Format(Language.CopyOK, otherLoginInfo.CompCode))
                    End If
                Catch ex As Exception
                    strRetMsg.AppendLine(String.Format(Language.CopyErr, otherLoginInfo.CompCode, ex.ToString))
                Finally
                    otherLoginInfo = Nothing
                    If otherDao IsNot Nothing Then
                        otherDao.Dispose()
                        otherDao = Nothing
                    End If
                    

                End Try
            Next
            result.ResultBoolean = True
            result.ResultXML = strRetMsg.ToString
            Return result
        Catch ex As Exception
            Throw ex
        Finally
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
                End If
            End If
            If dynaCdt IsNot Nothing Then
                dynaCdt.Dispose()
                dynaCdt = Nothing
            End If
            If dtDynReturn IsNot Nothing Then
                dtDynReturn.Dispose()
                dtDynReturn = Nothing
            End If
            If BeforeSQL IsNot Nothing Then
                BeforeSQL.Clear()
                BeforeSQL = Nothing
            End If
            If aChildInsertSQL IsNot Nothing Then
                aChildInsertSQL.Dispose()
                aChildInsertSQL = Nothing
            End If
            'If lstChildValue IsNot Nothing Then
            '    lstChildValue.Clear()
            '    lstChildValue = Nothing
            'End If
        End Try


    End Function

    Private Function getChildInserSQL(ByVal instTableName As String, ByVal tbChild As DataTable) As String
        Dim result As String = "Insert Into " & instTableName
        Dim aFields As String = String.Empty
        Dim aValues As String = String.Empty
        Dim i As Integer = 0
        For Each rw As DataRow In tbChild.Rows
            If String.IsNullOrEmpty(aFields) Then
                aFields = rw("FIELDNAME")
            Else
                aFields = aFields & "," & rw("FIELDNAME")
            End If
            If String.IsNullOrEmpty(aValues) Then
                aValues = ":0"
            Else
                aValues = aValues & ",:" & i.ToString
            End If
            i += 1
        Next
        result = result & " ( " & aFields & " ) Values ( " & aValues & " )"
        Return result
    End Function
    Private Function getConditionTable(ByVal dsSource As DataSet) As DataTable
        Dim dtCondition As DataTable = dsSource.Tables("condition").Clone

        Dim lstrow As List(Of DataRow) = dsSource.Tables("condition").AsEnumerable.Where(Function(ByVal rw As DataRow)
                                                                                             Return Integer.Parse(rw.Item("OBJECTTYPE")) = 30
                                                                                         End Function).ToList()
        If lstrow IsNot Nothing AndAlso lstrow.Count > 0 Then
            Dim aFieldNamd As String = lstrow.Item(0).Item("FieldName").ToString.Replace("_0", "").ToString
            aFieldNamd = aFieldNamd.Replace("_1", "")

            Dim lsfind As List(Of DataRow) = dsSource.Tables("condition").AsEnumerable.Where(Function(ByVal rw As DataRow)
                                                                                                 Return rw.Item("FieldName").ToString.IndexOf(aFieldNamd) AndAlso Integer.Parse(rw.Item("ObjectType")) <> 30

                                                                                             End Function).ToList

            For Each rwfind As DataRow In lsfind
                Dim rwnew As DataRow = dtCondition.NewRow
                rwnew.ItemArray = rwfind.ItemArray
                rwnew.Item("FieldName") = rwnew.Item("FieldName").ToString.Replace(aFieldNamd, "")
                dtCondition.Rows.Add(rwnew)
            Next
        End If
        Return dtCondition.Copy
    End Function

    Public Function Execute(ByVal EditMode As EditMode, ByVal SysProgramId As String,
                            ByVal dsSource As DataSet) As RIAResult
        Dim result As New RIAResult
        Dim dsReturn As New DataSet()
        Dim ErrMsg As String = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim cmd As DbCommand = Nothing
        Dim trans As DbTransaction = Nothing

        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim UpdateSQL As New List(Of String)
        Dim UpdateParams As New List(Of Object())
        Dim BeforeParams As New List(Of Object())
        Dim BeforeSQL As New List(Of String)
        Dim childBeforeSQL As New List(Of String)
        Dim childUpdateSQL As New List(Of String)
        Dim aBeforeAndFinallChildSQL As String = Nothing

        Dim InserChildSQL As New List(Of String)
        Dim dtDynReturn As DataTable = Nothing

        Dim aChildTableName As String = Nothing
        Dim tbOriginal As DataTable = Nothing
        Dim tbUpdate As DataTable = Nothing
        Dim lstChildValue As New Dictionary(Of Integer, List(Of Object))
        'Dim aInsertSQL As String = String.Empty
        'Dim aChildInsertSQL As New List(Of Array)
        Dim aChildInsertSQL As New ChildSQL()
        Dim PKErrMsg As String = Nothing
        Dim blnBeginTrancation As Boolean = False



        dynaCdt = New CableSoft.BLL.Dynamic.Condition.DynamicCondition(Me.LoginInfo, Me.DAO)

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
        CSLog = New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo, Me.DAO)
        cmd = DAO.GetConn.CreateCommand
        result.ResultBoolean = False
        blnBeginTrancation = True
        Try
            tbMaster = DAO.ExecQry(_DAL.QuerySO1109A, New Object() {SysProgramId})
            tbMaster.TableName = tbMasterName
            tbDetail = DAO.ExecQry(_DAL.QuerySO1109B, New Object() {tbMaster.Rows(0).Item("ProgramId")})
            tbDetail.TableName = tbDetailName
            If blnAutoClose Then
                Dim aAction As String = Nothing
                Select Case EditMode
                    Case Utility.EditMode.Append
                        aAction = String.Format(Language.AddClientInfo, tbMaster.Rows(0).Item("Caption"))
                    Case Utility.EditMode.Edit
                        aAction = String.Format(Language.EditClientInfo, tbMaster.Rows(0).Item("Caption"))
                    Case Utility.EditMode.Delete
                        aAction = String.Format(Language.DelClientInfo, tbMaster.Rows(0).Item("Caption"))
                    Case Else
                        aAction = String.Format(Language.EditClientInfo, tbMaster.Rows(0).Item("Caption"))
                End Select
                CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, aAction)
            End If
            If Not DBNull.Value.Equals(tbMaster.Rows(0).Item("FinalSQL")) Then
                UpdateSQL = tbMaster.Rows(0).Item("FinalSQL").ToString.Split(";").ToList
                'Array.Resize(UpdateParams.ToArray, UpdateSQL.Count)

            End If
            If Not DBNull.Value.Equals(tbMaster.Rows(0).Item("BeforeSQL")) Then
                BeforeSQL = tbMaster.Rows(0).Item("BeforeSQL").ToString.Split(";").ToList

            End If
            Dim params() As Object = Nothing
            dtDynReturn = dynaCdt.GetBuildConditionSQL(tbMaster.Rows(0).Item("SysProgramId"),
                                                       dsSource.Tables("Condition"), params)


            Dim aFieldName As String = Nothing

            For index As Integer = 0 To UpdateSQL.Count - 1
                Array.Resize(params, 0)
                For Each dr As DataRow In dtDynReturn.Rows
                    UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                   UpdateSQL.Item(index), dr("FieldName"), dr("ConditionSQL"), params)

                Next
                UpdateParams.Add(params)
            Next
            For index As Integer = 0 To BeforeSQL.Count - 1
                Array.Resize(params, 0)
                For Each dr As DataRow In dtDynReturn.Rows
                    BeforeSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                     BeforeSQL.Item(index), dr("FieldName"), dr("ConditionSQL"), params)
                Next
                BeforeParams.Add(params)

            Next
            For index As Integer = 0 To UpdateSQL.Count - 1
                For i As Integer = 0 To dsSource.Tables("Condition").Rows.Count - 1
                    If Right(dsSource.Tables("Condition").Rows(i)("FieldName"), 2) = "_0" OrElse Right(dsSource.Tables("Condition").Rows(i)("FieldName"), 2) = "_1" Then
                        aFieldName = dsSource.Tables("Condition").Rows(i)("FieldName").ToString.Substring(0,
                                                                                   dsSource.Tables("Condition").Rows(i)("FieldName").ToString.Length - 2)
                    Else
                        aFieldName = dsSource.Tables("Condition").Rows(i)("FieldName").ToString
                    End If
                    UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                    UpdateSQL.Item(index), aFieldName, dsSource.Tables("Condition").Rows(i)("FieldValue"), params)
                    UpdateSQL.Item(index) = ReplaceLoginInfoWhere(UpdateSQL.Item(index), params)
                Next
                'UpdateParams.Add(params)

            Next
            For index As Integer = 0 To BeforeSQL.Count - 1
                For i As Integer = 0 To dsSource.Tables("Condition").Rows.Count - 1
                    If Right(dsSource.Tables("Condition").Rows(i)("FieldName"), 2) = "_0" OrElse Right(dsSource.Tables("Condition").Rows(i)("FieldName"), 2) = "_1" Then
                        aFieldName = dsSource.Tables("Condition").Rows(i)("FieldName").ToString.Substring(0,
                                                                                   dsSource.Tables("Condition").Rows(i)("FieldName").ToString.Length - 2)
                    Else
                        aFieldName = dsSource.Tables("Condition").Rows(i)("FieldName").ToString
                    End If
                    BeforeSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                     BeforeSQL.Item(index), aFieldName, dsSource.Tables("Condition").Rows(i)("FieldValue"), params)
                    BeforeSQL.Item(index) = ReplaceLoginInfoWhere(BeforeSQL.Item(index), params)

                Next
                'BeforeParams.Add(params)
            Next




            ErrMsg = chkSchema()
            Dim indexOrder As Integer = 0
            For Each rw As DataRow In dsSource.Tables("Condition").Rows
                If DBNull.Value.Equals(rw("OBJECTTYPE")) OrElse Integer.Parse(rw("OBJECTTYPE") & "") <> 30 Then
                    rw.Item("FieldName") = Replace(rw("FieldName"), "_1", "")
                    rw.Item("FieldName") = Replace(rw("FieldName"), "_0", "")
                Else
                    rw.Item("FieldName") = Replace(rw("FieldName"), "_1", "")
                    rw.Item("FieldName") = Replace(rw("FieldName"), "_0", "")
                    aChildTableName = rw.Item("FieldValue")
                    tbChildSO1109B = DAO.ExecQry(_DAL.QuerySO1109B, New Object() {rw.Item("FieldName")})
                    tbChildSO1109A = DAO.ExecQry(_DAL.QuerySO1109A, New Object() {rw.Item("FieldName")})

                    'aInsertSQL = getChildInserSQL(rw.Item("FIELDVALUE").ToString, tbChildSO1109B)
                    Dim aChildSQL As String = getChildInserSQL(rw.Item("FIELDVALUE").ToString, tbChildSO1109B)
                    If Not DBNull.Value.Equals(tbChildSO1109A.Rows(0).Item("BeforeSQL")) Then
                        childBeforeSQL = tbChildSO1109A.Rows(0).Item("BeforeSQL").ToString.Split(";").ToList
                    End If
                    If Not DBNull.Value.Equals(tbChildSO1109A.Rows(0).Item("FinalSQL")) Then
                        childUpdateSQL = tbChildSO1109A.Rows(0).Item("FinalSQL").ToString.Split(";").ToList
                    End If
                    tbChildSchema = GetSechema(rw.Item("FIELDVALUE").ToString)

                    Try
                        For index As Integer = 0 To dsSource.Tables(rw.Item("FIELDVALUE").ToString).Rows.Count - 1
                            aChildInsertSQL.setSQL(aChildSQL, GetChildValue(tbChildSO1109B,
                                                                   dsSource.Tables(rw.Item("FIELDVALUE").ToString).Rows(index),
                                                                   dsSource, Me.DAO, Me.LoginInfo))


                        Next

                    Finally

                    End Try
                End If
            Next

            'For Each tb As DataTable In dsSource.Tables
            '    For Each rw As DataRow In tb.Rows
            '        rw.Item("FieldName") = Replace(rw("FieldName"), "_1", "")
            '        rw.Item("FieldName") = Replace(rw("FieldName"), "_0", "")
            '    Next
            'Next
            FillSourceField()
            If Not String.IsNullOrEmpty(ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            tbSechema = GetSechema()
            Dim aSQL As String = TakeSQL(EditMode, dsSource)
            '檢查PK
            If EditMode = Utility.EditMode.Append Then
                Dim aPKSQL As String = GetPKSQL()
                For i As Int32 = 0 To fWhereFieldsAndValues.Keys.Count - 1
                    For Each rw109b As DataRow In tbDetail.Rows
                        If rw109b.Item("FieldName").ToString.ToUpper = fWhereFieldsAndValues.Keys(i).ToUpper Then
                            If String.IsNullOrEmpty(PKErrMsg) Then
                                PKErrMsg = rw109b.Item("Caption")
                            Else
                                PKErrMsg = PKErrMsg & "+" & rw109b.Item("Caption")
                            End If
                        End If
                    Next
                Next
                If Not String.IsNullOrEmpty(PKErrMsg) Then
                    PKErrMsg = "[" & PKErrMsg & "] "
                End If
                If Integer.Parse(DAO.ExecSclr(aPKSQL, fWhereFieldsAndValues.Values.ToArray)) > 0 Then
                    result.ResultBoolean = True
                    result.ErrorCode = -1
                    result.ErrorMessage = String.Format(Language.DataExists, PKErrMsg)
                    trans.Rollback()
                    Return result
                End If
            End If
            '檢查UK
            fUKWhereFieldAndValues = New Dictionary(Of String, Object)
            Dim lstUKRw As List(Of DataRow) = GetUKRow(dsSource)
            If lstUKRw IsNot Nothing AndAlso lstUKRw.Count > 0 Then
                GetUKWhereList(dsSource, lstUKRw)
                Dim objValue As New List(Of Object)
                For i As Int32 = 0 To fUKWhereFieldAndValues.Keys.Count - 1
                    For Each rw109b As DataRow In tbDetail.Rows
                        If rw109b.Item("FieldName").ToString.ToUpper = fUKWhereFieldAndValues.Keys(i).ToUpper Then
                            Dim aUKSQL = GetUKSQL(EditMode, fUKWhereFieldAndValues.Keys(i))
                            objValue.Clear()
                            objValue.Add(fUKWhereFieldAndValues.Values(i))
                            If EditMode = Utility.EditMode.Edit Then
                                objValue.AddRange(fWhereFieldsAndValues.Values.ToArray)
                            End If

                            If DAO.ExecSclr(aUKSQL, objValue.ToArray) > 0 Then
                                PKErrMsg = "[" & rw109b.Item("Caption") & "] "
                                result.ResultBoolean = True
                                result.ErrorCode = -1
                                result.ErrorMessage = String.Format(Language.DataExists, PKErrMsg)
                                trans.Rollback()
                                Return result
                            End If
                        End If
                    Next
                Next

            End If

            Dim objValues As New List(Of Object)
            objValues.AddRange(fFieldsAndValues.Values)
            If fWhereFieldsAndValues IsNot Nothing AndAlso fWhereFieldsAndValues.Count > 0 Then
                If EditMode <> Utility.EditMode.Append Then
                    objValues.AddRange(fWhereFieldsAndValues.Values)
                End If
                If (EditMode = Utility.EditMode.Edit) OrElse (EditMode = Utility.EditMode.Delete) Then
                    tbOriginal = DAO.ExecQry(_DAL.QueryCurrectData(tbMaster.Rows(0).Item("TableName"), fWhereFieldsAndValues),
                                                    fWhereFieldsAndValues.Values.ToArray).Copy()
                End If
            End If
            If EditMode <> Utility.EditMode.Append Then
                For i As Integer = 0 To BeforeSQL.Count - 1
                    If BeforeParams.Item(i).Count > 0 Then
                        DAO.ExecNqry(BeforeSQL.Item(i), BeforeParams.Item(i))
                    Else
                        DAO.ExecNqry(BeforeSQL.Item(i))
                    End If
                Next
                'For Each befSQL As String In BeforeSQL

                'Next
            End If
            DAO.ExecNqry(aSQL, objValues.ToArray)
            If (EditMode = Utility.EditMode.Edit) Then
                tbUpdate = DAO.ExecQry(_DAL.QueryCurrectData(tbMaster.Rows(0).Item("TableName"), fWhereFieldsAndValues),
                                                fWhereFieldsAndValues.Values.ToArray).Copy()
                For i As Integer = 0 To tbOriginal.Rows.Count - 1
                    For Each col As DataColumn In tbOriginal.Columns
                        tbOriginal.Rows(i).Item(col.ColumnName) = tbUpdate.Rows(i).Item(col.ColumnName)
                    Next
                Next
            End If


            For index As Integer = 0 To aChildInsertSQL.TotalCount - 1
                'objectype =30 Details'before sql
                For i As Integer = 0 To childBeforeSQL.Count - 1
                    aBeforeAndFinallChildSQL = childBeforeSQL.Item(i)
                    Array.Resize(params, 0)
                    For Each col As DataColumn In dsSource.Tables(aChildTableName).Columns
                        aBeforeAndFinallChildSQL = ReplaceChildSQL(aBeforeAndFinallChildSQL, col.ColumnName, dsSource.Tables(aChildTableName).Rows(index).Item(col.ColumnName), params)
                    Next
                    aBeforeAndFinallChildSQL = ReplaceLoginInfoWhere(aBeforeAndFinallChildSQL, params)
                    DAO.ExecNqry(aBeforeAndFinallChildSQL)
                Next

                DAO.ExecNqry(aChildInsertSQL.readInsSQL(index), aChildInsertSQL.readValues(index))
                'objectype =30 Details'finall sql
                For i As Integer = 0 To childUpdateSQL.Count - 1
                    aBeforeAndFinallChildSQL = childUpdateSQL.Item(i)
                    Array.Resize(params, 0)
                    For Each col As DataColumn In dsSource.Tables(aChildTableName).Columns
                        aBeforeAndFinallChildSQL = ReplaceChildSQL(aBeforeAndFinallChildSQL, col.ColumnName, dsSource.Tables(aChildTableName).Rows(index).Item(col.ColumnName), params)
                    Next
                    aBeforeAndFinallChildSQL = ReplaceLoginInfoWhere(aBeforeAndFinallChildSQL, params)
                    DAO.ExecNqry(aBeforeAndFinallChildSQL)
                Next
            Next



            If UpdateSQL IsNot Nothing Then
                For i As Integer = 0 To UpdateSQL.Count - 1
                    If UpdateParams.Item(i).Count > 0 Then
                        DAO.ExecNqry(UpdateSQL.Item(i), UpdateParams.Item(i))
                    Else
                        DAO.ExecNqry(UpdateSQL.Item(i))
                    End If
                Next
                'For Each updSQL As String In UpdateSQL

                'Next
            End If

            Select Case EditMode
                Case Utility.EditMode.Append
                    dsReturn.Tables.Add(DAO.ExecQry(_DAL.QueryCurrectData(tbMaster.Rows(0).Item("TableName"),
                                                                          fWhereFieldsAndValues),
                                                    fWhereFieldsAndValues.Values.ToArray).Copy)
                Case Else

                    dsReturn.Tables.Add(
                        DAO.ExecQry(_DAL.QueryCurrectData(tbMaster.Rows(0).Item("TableName"), fWhereFieldsAndValues),
                                    fWhereFieldsAndValues.Values.ToArray).Copy)
            End Select
            If (EditMode = Utility.EditMode.Edit) OrElse (EditMode = Utility.EditMode.Delete) Then
                result = CSLog.SummaryExpansion(cmd, IIf(EditMode = Utility.EditMode.Delete, OpType.Delete, OpType.Update),
                                            tbMaster.Rows(0).Item("TableName"),
                                  tbOriginal,
                                  Integer.Parse(DateTime.Now.ToString("yyyyMMdd")))

                If Not result.ResultBoolean Then
                    Select Case result.ErrorCode
                        Case -5
                        Case -6
                            If blnAutoClose Then
                                trans.Rollback()
                                Return result
                            End If
                    End Select
                End If
            End If

            result.ResultBoolean = True
            result.ErrorCode = 0
            result.ErrorMessage = Nothing
            result.ResultDataSet = dsReturn
            If blnAutoClose Then
                trans.Commit()
            End If
        Catch exOracle As OracleClient.OracleException
            result.ResultBoolean = False
            result.ErrorMessage = exOracle.ToString
            result.ErrorCode = -2
            If blnAutoClose AndAlso blnBeginTrancation Then
                trans.Rollback()
            End If
            'Throw exOracle
        Catch ex As Exception

            result.ResultBoolean = False
            result.ErrorMessage = ex.ToString
            result.ErrorCode = -3
            If blnAutoClose AndAlso blnBeginTrancation Then
                trans.Rollback()
            End If
            'Throw ex
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                DAO.AutoCloseConn = True
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cmd IsNot Nothing Then
                    cmd.Dispose()
                    cmd = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                    cn = Nothing
                End If

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
            If aChildInsertSQL IsNot Nothing Then
                'aChildInsertSQL.Clear()
                aChildInsertSQL = Nothing
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
            If tbChildSO1109B IsNot Nothing Then
                tbChildSO1109B.Dispose()
                tbChildSO1109B = Nothing
            End If
            If tbUpdate IsNot Nothing Then
                tbUpdate.Dispose()
                tbUpdate = Nothing
            End If
            If dsSource IsNot Nothing Then
                dsSource.Dispose()
                dsSource = Nothing
            End If
            If CSLog IsNot Nothing Then
                CSLog.Dispose()
                CSLog = Nothing
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
            If aChildInsertSQL IsNot Nothing Then
                aChildInsertSQL.Dispose()
                aChildInsertSQL = Nothing
            End If
        End Try
        Return result
    End Function
    Private Function ReplaceChildSQL(ByVal aSQL As String, ByVal colnumName As String, ByVal columnValue As Object, ByRef params() As Object) As String

        Dim retSQL As String = Nothing
        Try
            retSQL = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, aSQL, colnumName, columnValue, params)

        Catch ex As Exception
            Throw ex
        End Try
        Return retSQL
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
    Private Function GetLoginInfoValue(ByVal PropertyName As String, ByRef exeLoginInfo As LoginInfo) As Object

        For Each pi As PropertyInfo In exeLoginInfo.GetType().GetProperties
            If pi.Name.ToUpper = PropertyName.ToUpper Then
                Return pi.GetValue(exeLoginInfo, Nothing)
            End If
        Next

        Return Nothing
    End Function

    Private Function GetLoginInfoValue(ByVal PropertyName As String) As Object
        Return GetLoginInfoValue(PropertyName, Me.LoginInfo)

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
            Throw New Exception(String.Format(Language.ConverDataTypeError,
                                              FieldName, Value.ToString,
                                              tbSchema.Columns(FieldName).DataType.Name))
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 如果SO1109B.SourceField後２碼沒有"_0"，則程式自動補上＂_0＂
    ''' </summary>
    ''' <remarks>設定檔可設可不設</remarks>
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
    Private Function GetCopyPKRow() As List(Of DataRow)
        Dim lstPKRw As New List(Of DataRow)
        Dim strArray As String = "CodeNo,Description"
        For Each Str As String In strArray.Split(",")
            For Each rw As DataRow In tbDetail.Rows
                If rw.Item("FieldName").ToString.ToUpper = Str.ToUpper Then
                    Dim rwNew As DataRow = tbDetail.NewRow
                    rwNew.ItemArray = rw.ItemArray
                    lstPKRw.Add(rwNew)
                End If
            Next
        Next
        Return lstPKRw
    End Function
    Private Function GetUKRow(dsSource As DataSet) As List(Of DataRow)
        Dim lstUKRw As List(Of DataRow) = Nothing

        'Dim tbName As String = tb.TableName
        Dim lstRw As List(Of DataRow) = Nothing
        For Each rw As DataRow In dsSource.Tables("Condition").Rows
            Dim i As Int32 = tbDetail.AsEnumerable.Where(Function(rwDetail As DataRow)
                                                             If (rwDetail.Item("FieldType") = 2) AndAlso (Not DBNull.Value.Equals(rw.Item("FieldValue"))) Then
                                                                 Dim aSourceField As String = rwDetail.Item("SourceField").ToString.ToUpper

                                                                 If (rwDetail.Item("SourceField").ToString.ToUpper = rw.Item("FieldName").ToString.ToUpper) Then
                                                                     Return True
                                                                 Else
                                                                     Return False
                                                                 End If
                                                             End If
                                                             Return False
                                                         End Function).Count


            lstRw = tbDetail.AsEnumerable.Where(Function(rwDetail As DataRow)
                                                    If (rwDetail.Item("SourceField").ToString.ToUpper = rw.Item("FieldName").ToString.ToUpper) AndAlso
                                                        (rwDetail.Item("FieldType") = 2) AndAlso (Not DBNull.Value.Equals(rw.Item("FieldValue"))) Then
                                                        Return True
                                                    End If
                                                    Return False
                                                End Function).ToList

            If lstRw IsNot Nothing AndAlso lstRw.Count > 0 Then
                If lstUKRw Is Nothing Then
                    lstUKRw = New List(Of DataRow)
                End If
                lstUKRw.AddRange(lstRw)
            End If
        Next



        For Each rw As DataRow In tbDetail.Rows
            If rw.Item("FieldType") = 1 Then
                Dim blnAdd As Boolean = False
                Select Case rw.Item("SourceTable").ToString.ToUpper
                    Case SeqNoString.ToUpper
                        blnAdd = True
                    Case LoginInfoString.ToUpper
                        blnAdd = True
                End Select
                If blnAdd Then
                    If lstUKRw Is Nothing Then
                        lstUKRw = New List(Of DataRow)
                    End If
                    lstUKRw.Add(rw)
                End If
            End If

        Next
        Return lstUKRw
    End Function
    ''' <summary>
    ''' 取得設定檔設定的PK
    ''' </summary>
    ''' <param name="dsSource"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetPKRow(ByVal dsSource As DataSet) As List(Of DataRow)
        Dim lstPkRw As List(Of DataRow) = Nothing

        'Dim tbName As String = tb.TableName
        Dim lstRw As List(Of DataRow) = Nothing
        For Each rw As DataRow In dsSource.Tables("Condition").Rows
            Dim i As Int32 = tbDetail.AsEnumerable.Where(Function(rwDetail As DataRow)
                                                             If (rwDetail.Item("FieldType") = 1) AndAlso (Not DBNull.Value.Equals(rw.Item("FieldValue"))) Then
                                                                 Dim aSourceField As String = rwDetail.Item("SourceField").ToString.ToUpper

                                                                 If (rwDetail.Item("SourceField").ToString.ToUpper = rw.Item("FieldName").ToString.ToUpper) Then
                                                                     Return True
                                                                 Else
                                                                     Return False
                                                                 End If
                                                             End If
                                                             Return False
                                                         End Function).Count


            lstRw = tbDetail.AsEnumerable.Where(Function(rwDetail As DataRow)
                                                    If (rwDetail.Item("SourceField").ToString.ToUpper = rw.Item("FieldName").ToString.ToUpper) AndAlso
                                                        (rwDetail.Item("FieldType") = 1) Then
                                                        Return True
                                                    End If
                                                    Return False
                                                End Function).ToList

            If lstRw IsNot Nothing AndAlso lstRw.Count > 0 Then
                If lstPkRw Is Nothing Then
                    lstPkRw = New List(Of DataRow)
                End If
                lstPkRw.AddRange(lstRw)
            End If
        Next



        For Each rw As DataRow In tbDetail.Rows
            If rw.Item("FieldType") = 1 Then
                Dim blnAdd As Boolean = False
                Select Case rw.Item("SourceTable").ToString.ToUpper
                    Case SeqNoString.ToUpper
                        blnAdd = True
                    Case LoginInfoString.ToUpper
                        blnAdd = True
                End Select
                If blnAdd Then
                    If lstPkRw Is Nothing Then
                        lstPkRw = New List(Of DataRow)
                    End If
                    lstPkRw.Add(rw)
                End If
            End If

        Next
        Return lstPkRw
    End Function
    Private Function GetUKWhereList(ByVal dsSource As DataSet, ByVal otherLoginInfo As LoginInfo,
                                  ByVal lstUKRw As List(Of DataRow)) As Boolean
        Try
            For Each rw As DataRow In lstUKRw
                If Integer.Parse("0" & rw.Item("FieldType")) = 2 Then
                    Select Case rw.Item("SourceTable").ToString.ToUpper
                        Case LoginInfoString.ToUpper
                            Dim aProperty As Object = GetLoginInfoValue(rw.Item("SourceField"), otherLoginInfo)
                            If aProperty IsNot Nothing Then
                                fUKWhereFieldAndValues.Add(rw.Item("FieldName"),
                                                          ConverDataType(rw.Item("FieldName"), tbSechema, aProperty))
                            End If
                        Case SeqNoString.ToUpper
                            Dim aSeqNoValue As Object = Nothing
                            For Each tb As DataTable In dsSource.Tables
                                For Each rwSource As DataRow In tb.Rows
                                    If rwSource.Item("FieldName").ToString.ToUpper = rw.Item("SourceField").ToString.ToUpper Then
                                        aSeqNoValue = rwSource.Item("FieldValue")
                                        Exit For
                                    End If
                                Next
                                If aSeqNoValue IsNot Nothing Then
                                    Exit For
                                End If
                            Next
                            fUKWhereFieldAndValues.Add(rw.Item("FieldName"),
                                                      ConverDataType(rw.Item("FieldName"), tbSechema, aSeqNoValue))
                        Case Else
                            Dim aWhereValue As Object = Nothing
                            Dim aFieldValueName As String = "FieldValue"
                            If rw.Item("GetDesc") Then
                                aFieldValueName = "FieldDesc"
                            End If

                            For Each rwSource As DataRow In dsSource.Tables("Condition").Rows
                                'If tb.TableName.ToUpper = rw.Item("SourceTable").ToString.ToUpper Then
                                If (rwSource("FieldName").ToString.ToUpper = rw.Item("SourceField").ToString.ToUpper) AndAlso
                                    (Not DBNull.Value.Equals(aFieldValueName)) Then
                                    fUKWhereFieldAndValues.Add(rw.Item("FieldName"),
                                                              ConverDataType(rw.Item("FieldName"), tbSechema, rwSource.Item(aFieldValueName)))
                                    Exit For
                                End If
                                'End If
                            Next

                    End Select
                End If
               
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return True
    End Function
    Private Function GetWhereList(ByVal EditMode As EditMode, ByVal dsSource As DataSet, ByVal otherLoginInfo As LoginInfo,
                                  ByVal lstPKRw As List(Of DataRow)) As Boolean
        Try
            For Each rw As DataRow In lstPKRw
                Select Case rw.Item("SourceTable").ToString.ToUpper
                    Case LoginInfoString.ToUpper
                        Dim aProperty As Object = GetLoginInfoValue(rw.Item("SourceField"), otherLoginInfo)
                        If aProperty IsNot Nothing Then
                            fWhereFieldsAndValues.Add(rw.Item("FieldName"),
                                                      ConverDataType(rw.Item("FieldName"), tbSechema, aProperty))
                        End If
                    Case SeqNoString.ToUpper
                        Dim aSeqNoValue As Object = Nothing
                        For Each tb As DataTable In dsSource.Tables
                            For Each rwSource As DataRow In tb.Rows
                                If rwSource.Item("FieldName").ToString.ToUpper = rw.Item("SourceField").ToString.ToUpper Then
                                    aSeqNoValue = rwSource.Item("FieldValue")
                                    Exit For
                                End If
                            Next
                            If aSeqNoValue IsNot Nothing Then
                                Exit For
                            End If
                        Next
                        fWhereFieldsAndValues.Add(rw.Item("FieldName"),
                                                  ConverDataType(rw.Item("FieldName"), tbSechema, aSeqNoValue))
                    Case Else
                        Dim aWhereValue As Object = Nothing
                        Dim aFieldValueName As String = "FieldValue"
                        If rw.Item("GetDesc") Then
                            aFieldValueName = "FieldDesc"
                        End If

                        For Each rwSource As DataRow In dsSource.Tables("Condition").Rows
                            'If tb.TableName.ToUpper = rw.Item("SourceTable").ToString.ToUpper Then
                            If (rwSource("FieldName").ToString.ToUpper = rw.Item("SourceField").ToString.ToUpper) AndAlso
                                (Not DBNull.Value.Equals(aFieldValueName)) Then
                                fWhereFieldsAndValues.Add(rw.Item("FieldName"),
                                                          ConverDataType(rw.Item("FieldName"), tbSechema, rwSource.Item(aFieldValueName)))
                                Exit For
                            End If
                            'End If
                        Next

                End Select
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return True
    End Function
    Private Function GetUKWhereList(ByVal dsSource As DataSet, ByVal lstUKRw As List(Of DataRow)) As Boolean
        Return GetUKWhereList(dsSource, Me.LoginInfo, lstUKRw)
    End Function
    Private Function GetWhereList(ByVal EditMode As EditMode, ByVal dsSource As DataSet,
                                  ByVal lstPKRw As List(Of DataRow)) As Boolean

        Return GetWhereList(EditMode, dsSource, Me.LoginInfo, lstPKRw)
    End Function
    Private Function GetInsertAllDataSQL(ByVal dtSechema As DataTable) As String
        Return _DAL.GetInsertAllDataSQL(tbMaster.Rows(0).Item("TableName"), dtSechema)
    End Function
    Private Function GetFindSQL() As String
        Return _DAL.GetFindSQL(tbMaster.Rows(0).Item("TableName"), fWhereFieldsAndValues)
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
    Private Function GetInsertSQL(ByVal dsSource As DataSet) As String
        Return GetInsertSQL(dsSource, Me.DAO, Me.LoginInfo)
    End Function
    Private Function GetChildValue(ByVal tbChildSO1109B As DataTable, ByVal rwDyn As DataRow,
                                   ByVal dsSource As DataSet,
                                   ByRef exeDao As CableSoft.Utility.DataAccess.DAO,
                                                 ByRef exeLoginInfo As LoginInfo) As List(Of Object)
        Dim result As New List(Of Object)
        Dim aValue As Object
        For Each rwSO1109B As DataRow In tbChildSO1109B.Rows
            aValue = DBNull.Value
            '如果動態條件的欄位在SO1109B有找到,則直接用動態條件的值帶入
            If rwDyn.Table.Columns.Contains(rwSO1109B("SOURCEFIELD")) Then
                aValue = rwDyn(rwSO1109B("SOURCEFIELD"))
                aValue = ConverDataType(rwSO1109B("FieldName"), tbChildSchema, aValue)
                result.Add(aValue)
            Else
                '動態條件的欄位在SO1109B沒找到
                If Not DBNull.Value.Equals(rwSO1109B.Item("FinalValue")) Then
                    result.Add(ConverDataType(rwSO1109B("FieldName").ToString, tbSechema,
                                                         GetFinalValue(rwSO1109B.Item("FinalValue").ToString, dsSource, exeDao)))


                    'If Not fFieldsAndValues.ContainsKey(rwSO1109B("FieldName").ToString) Then
                    '    fFieldsAndValues.Add(rwSO1109B("FieldName").ToString,
                    '                      ConverDataType(rwSO1109B("FieldName").ToString, tbSechema,
                    '                                     GetFinalValue(rwSO1109B.Item("FinalValue").ToString, dsSource, exeDao)))
                    'End If

                Else
                    Select Case rwSO1109B.Item("SourceTable").ToString.ToUpper
                        Case LoginInfoString.ToUpper
                            Dim aProperty As Object = Nothing
                            aProperty = GetLoginInfoValue(rwSO1109B.Item("SourceField"), exeLoginInfo)
                            If aProperty IsNot Nothing Then
                                result.Add(ConverDataType(rwSO1109B("FieldName").ToString, tbChildSchema, aProperty))
                            End If
                        Case SeqNoString.ToUpper
                            result.Add(ConverDataType(rwSO1109B("FieldName").ToString,
                                                                                                tbSechema,
                                                                                                GetSeqNo(rwSO1109B.Item("SourceField"), exeDao)))
                            'If Not fFieldsAndValues.ContainsKey(rwSO1109B("FieldName").ToString) Then
                            '    fFieldsAndValues.Add(rwSO1109B("FieldName").ToString, ConverDataType(rwSO1109B("FieldName").ToString,
                            '                                                                    tbSechema,
                            '                                                                    GetSeqNo(rwSO1109B.Item("SourceField"), exeDao)))
                            'End If

                            'Case Else
                            '    If Not fFieldsAndValues.ContainsKey(rwSO1109B("FieldName").ToString) Then
                            '        fFieldsAndValues.Add(rwDetail("FieldName").ToString,
                            '                        ConverDataType(rwDetail("FieldName").ToString, tbSechema,
                            '                                       GetFinalValue(rwDetail.Item("FinalValue").ToString, dsSource, exeDao)))
                            '    End If

                    End Select
                End If

            End If
        Next
        Return result
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
                                                                                                    GetSeqNo(rwDetail.Item("SourceField"), exeDao)))
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
    ''' <summary>
    ''' 取出預設值到FieldsAndValues集合裡
    ''' </summary>
    ''' <param name="EditMode">編輯模式</param>
    ''' <returns>True Or False</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultToFieldsAndValues(ByVal EditMode As EditMode, ByVal dsSource As DataSet) As Boolean
        Return GetDefaultToFieldsAndValues(EditMode, dsSource, Me.DAO, Me.LoginInfo)
    End Function
    Private Function GetSeqNo(ByVal SourceField As String, ByRef exeDao As CableSoft.Utility.DataAccess.DAO) As Object
        'Dim aSQL As String = "SELECT " & SourceField & ".NEXTVAL FROM DUAL"
        Dim aSQL As String = _DAL.getSEQNo(SourceField)
        Try
            Return exeDao.ExecSclr(aSQL)
        Catch ex As Exception
            Throw New Exception(String.Format(Language.GetSeqNoError, aSQL, "GetSeqNo"))
        End Try

    End Function
    Private Function GetSeqNo(ByVal SourceField As String) As Object
        Return GetSeqNo(SourceField, Me.DAO)
    End Function
    Private Function GetDelSQL(ByVal dsSource As DataSet) As String

        Try
            Return _DAL.GetDelSQL(tbMaster.Rows(0).Item("TableName"), fWhereFieldsAndValues)
        Catch ex As Exception
            Throw
        End Try

    End Function
    Private Function GetUKSQL(ByVal editMode As EditMode, ByVal FieldName As String) As String
        Dim aRet As String = Nothing
        Try
            aRet = _DAL.GetUKSQL(editMode, tbMaster.Rows(0).Item("TableName"), FieldName, fWhereFieldsAndValues)
        Catch ex As Exception
            Throw
        End Try
        Return aRet
    End Function
    Private Function GetPKSQL() As String
        Dim aRet As String = Nothing
        Try
            aRet = _DAL.GetPKSQL(tbMaster.Rows(0).Item("TableName"), fWhereFieldsAndValues)
        Catch ex As Exception
            Throw
        End Try
        Return aRet
    End Function

    Private Function GetUpdateSQL(ByVal dsSource As DataSet) As String
        Dim aRet As String = Nothing
        Try
            Dim aFieldValueName As String = "FieldValue"


            For Each rwSource As DataRow In dsSource.Tables("Condition").Rows
                For Each rwDetail As DataRow In tbDetail.Rows
                    aFieldValueName = "FieldValue"
                    If rwDetail.Item("GetDesc") Then
                        aFieldValueName = "FieldDesc"
                    End If
                    If rwDetail.Item("FieldType") <> 1 Then 'Update PK值要跳過因為要當Where條件
                        ' If tbSource.TableName.ToUpper = rwDetail.Item("SourceTable").ToString.ToUpper Then
                        If (rwDetail.Item("SourceField").ToString.ToUpper = rwSource("FieldName").ToString.ToUpper) Then
                            fFieldsAndValues.Add(rwDetail.Item("FieldName"),
                                                 ConverDataType(rwDetail.Item("FieldName"), tbSechema, rwSource(aFieldValueName)))
                        End If
                        'End If
                    End If
                Next
            Next


            '取出預設值
            GetDefaultToFieldsAndValues(EditMode.Edit, dsSource)

            'Dim lstRw As List(Of DataRow) = tbDetail.AsEnumerable.Where(Function(rw As DataRow)
            '                                                                If rw.Item("FieldType") <> 1 Then
            '                                                                    If Not DBNull.Value.Equals(rw.Item("FinalValue")) Then
            '                                                                        Return True
            '                                                                    End If
            '                                                                    If rw.Item("SourceTable").ToString.ToUpper = LoginInfoString.ToUpper OrElse
            '                                                                        rw.Item("SourceTable").ToString.ToUpper = SeqNoString.ToUpper Then
            '                                                                        Return True
            '                                                                    End If
            '                                                                End If
            '                                                                Return False
            '                                                            End Function).ToList

            'If lstRw IsNot Nothing AndAlso lstRw.Count > 0 Then
            '    For Each rwDetail As DataRow In lstRw
            '        Select Case rwDetail.Item("SourceTable").ToString.ToUpper
            '            Case LoginInfoString.ToUpper
            '                Dim aProperty As Object = Nothing
            '                aProperty = GetLoginInfoValue(rwDetail.Item("SourceField"))
            '                If aProperty IsNot Nothing Then
            '                    fFieldsAndValues.Add(rwDetail("FieldName").ToString,
            '                                         ConverDataType(rwDetail("FieldName").ToString, tbSechema, aProperty))
            '                End If
            '            Case SeqNoString.ToUpper
            '            Case Else
            '                fFieldsAndValues.Add(rwDetail("FieldName").ToString,
            '                                     ConverDataType(rwDetail("FieldName").ToString, tbSechema,
            '                                                    GetFinalValue(rwDetail.Item("FinalValue").ToString)))
            '        End Select
            '    Next

            'End If
            aRet = _DAL.GetUpdateSQL(tbMaster.Rows(0).Item("TableName"), fFieldsAndValues, fWhereFieldsAndValues)
        Catch ex As Exception
            Throw ex
        End Try
        Return aRet
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
                        'Return DAO.ExecSclr("select sysdate from dual")
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
        Catch ex As Exception
            Throw ex
        Finally
            If dtDynReturn IsNot Nothing Then
                dtDynReturn.Dispose()
                dtDynReturn = Nothing
            End If
        End Try

    End Function
    Private Function GetFinalValue(ByVal strSource As String, ByVal dsSource As DataSet) As Object
        Return GetFinalValue(strSource, dsSource, Me.DAO)
    End Function
    Private Function TakeSQL(ByVal EditMode As EditMode, ByVal dsSource As DataSet) As String
        Dim aRet As String = Nothing

        Dim lstPkRw As List(Of DataRow) = Nothing
        lstPkRw = GetPKRow(dsSource)

        If lstPkRw Is Nothing OrElse lstPkRw.Count = 0 Then
            Throw New Exception(Language.NotFoundPKValue)
        End If
        If fFieldsAndValues Is Nothing Then
            fFieldsAndValues = New Dictionary(Of String, Object)
        End If

        If fWhereFieldsAndValues Is Nothing Then
            fWhereFieldsAndValues = New Dictionary(Of String, Object)
        End If
        fWhereFieldsAndValues.Clear()
        fFieldsAndValues.Clear()
        GetWhereList(EditMode, dsSource, lstPkRw)
        If EditMode <> Utility.EditMode.Append Then
            If fWhereFieldsAndValues Is Nothing OrElse fWhereFieldsAndValues.Count = 0 Then
                Throw New Exception(String.Format(Language.GetNoWhere, "TakeSQL"))
            End If
        End If
        Select Case EditMode
            Case Utility.EditMode.Append
                aRet = GetInsertSQL(dsSource)
            Case Utility.EditMode.Edit
                aRet = GetUpdateSQL(dsSource)
            Case Utility.EditMode.Delete
                aRet = GetDelSQL(dsSource)
            Case Else
                aRet = GetUpdateSQL(dsSource)
        End Select
        'If EditMode <> Utility.EditMode.Append Then
        '    aRet = GetUpdateSQL(dsSource)
        'Else
        '    aRet = GetInsertSQL(dsSource)
        'End If

        Return aRet
    End Function

    Private Function chkQueryData(ByVal dsSource As DataSet) As Boolean
        If dsSource Is Nothing OrElse dsSource.Tables.Count = 0 Then
            Throw New Exception(Language.QueryDataError)
        End If
        Try

            For Each tb As DataTable In dsSource.Tables
                Dim tbName As String = tb.TableName
                For Each rw As DataRow In tb.Rows
                    If tbDetail.AsEnumerable.Where(Function(rwDetail As DataRow, bln As Boolean)
                                                       Select Case rwDetail.Item("SourceTable").ToString.ToUpper
                                                           Case LoginInfoString.ToUpper
                                                               Return True
                                                           Case SeqNoString.ToUpper
                                                               Return True
                                                           Case Else
                                                               If rwDetail.Item("SourceTable").ToString.ToUpper = tbName Then
                                                                   Return rwDetail("SourceField").ToString.ToUpper = rw.Item("FieldName").ToString.ToUpper
                                                               Else
                                                                   Return False
                                                               End If
                                                       End Select
                                                       Return True
                                                   End Function).Count = 0 Then
                        Throw New Exception(String.Format(Language.NotFoundSourceField,
                                                          tbName, rw.Item("FieldName"), "SourceField"))

                    End If
                Next
            Next
        Catch ex As Exception
            Throw ex
        End Try

        Return True
    End Function
    Public Function QueryEnvironment(ByVal SysProgramId As String) As DataSet
        Dim ds As New DataSet
        Dim ErrMsg As String = Nothing
        tbMaster = DAO.ExecQry(_DAL.QuerySO1109A, New Object() {SysProgramId})

        If tbMaster Is Nothing OrElse tbMaster.Rows.Count = 0 Then
            Throw New Exception(String.Format(Language.NoSO1109AData, SysProgramId))
        End If
        tbMaster.TableName = tbMasterName
        tbDetail = DAO.ExecQry(_DAL.QuerySO1109B, New Object() {tbMaster.Rows(0).Item("ProgramId")})
        If tbDetail Is Nothing OrElse tbDetail.Rows.Count = 0 Then
            Throw New Exception(String.Format(Language.NoSO1109BData, tbMaster.Rows(0).Item("ProgramId")))
        End If
        tbDetail.TableName = tbDetailName
        ErrMsg = chkSO1109A()
        If Not String.IsNullOrEmpty(ErrMsg) Then
            Throw New Exception(ErrMsg)
        End If
        ErrMsg = chkSO1109B()
        If Not String.IsNullOrEmpty(ErrMsg) Then
            Throw New Exception(ErrMsg)
        End If
        ds.Tables.Add(tbMaster.Copy)
        ds.Tables.Add(tbDetail.Copy)
        Return ds
    End Function
    Private Function chkSO1109A() As String
        If DBNull.Value.Equals(tbMaster.Rows(0).Item("TableName")) Then
            Return Language.NoTableName
        End If
        If DBNull.Value.Equals(tbMaster.Rows(0).Item("CondProgId")) Then
            Return Language.NoCondProgIdProp
        End If
        Return Nothing
    End Function
    Private Function chkSO1109B() As String
        For Each rw As DataRow In tbDetail.Rows
            If DBNull.Value.Equals(rw.Item("FieldName")) Then
                Return String.Format(Language.NoSetSO1109BField, rw.Item("AutoSerialNo"))
            End If
            If (DBNull.Value.Equals(rw.Item("SourceTable"))) OrElse
                (DBNull.Value.Equals(rw.Item("SourceField"))) Then
                If (Not String.IsNullOrEmpty(DefaultField)) AndAlso (tbDetail.Columns.Contains(DefaultField)) Then
                    If DBNull.Value.Equals(rw.Item(DefaultField)) Then
                        Return String.Format(Language.DetailFieldMustBe, rw.Item("AutoSerialNo"))
                    End If
                Else
                    Return String.Format(Language.DetailFieldMustBe, rw.Item("AutoSerialNo"))
                End If
            End If

        Next
        If tbDetail.AsEnumerable.Where(Function(rw As DataRow)
                                           Return rw.Item("FieldType") = 1
                                       End Function).Count = 0 Then
            Return Language.NoSetFieldTypeOne
        End If

        Return Nothing
    End Function
    ''' <summary>
    ''' 檢查設定檔是Schema是否符合SO1109A.Tablename的Schema
    ''' </summary>
    ''' <returns>ErrorMessage</returns>
    ''' <remarks></remarks>
    Private Function chkSchema() As String
        Using tbSchema As DataTable = DAO.ExecQry(_DAL.QuerySchema(tbMaster.Rows(0).Item("TableName")))
            If tbSchema Is Nothing Then
                Return String.Format(Language.NotFoundTable, tbMaster.Rows(0).Item("TableName"))
            End If
            For Each rw As DataRow In tbDetail.Rows
                If Not tbSchema.Columns.Contains(rw.Item("FieldName")) Then
                    Return String.Format(Language.NotFoundField,
                                         tbMaster.Rows(0).Item("TableName"),
                                         rw.Item("FieldName"), rw.Item("AutoSerialNo"))
                End If
            Next
        End Using
        Return Nothing
    End Function
    Private Class ChildSQL
        Implements IDisposable

        Dim insSQL As New Dictionary(Of Integer, String)
        Dim insValue As New Dictionary(Of Integer, Object())
        Property TotalCount As Integer = 0

        Public Sub New()

        End Sub
        Public Sub setSQL(ByVal sql As String, ByVal lstObj As List(Of Object))
            insSQL.Add(TotalCount, sql)
            insValue.Add(TotalCount, lstObj.ToArray)
            TotalCount += 1
        End Sub
        Public Function readInsSQL(ByVal index As Integer) As String
            Return insSQL(index)
        End Function
        Public Function readValues(ByVal index As Integer) As Object()
            Return insValue(index)
        End Function
        Public Sub Clear()
            insSQL.Clear()
            insValue.Clear()
            Me.TotalCount = 0
        End Sub
#Region "IDisposable Support"
        Private disposedValue As Boolean ' 偵測多餘的呼叫

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: 處置 Managed 狀態 (Managed 物件)。
                    TotalCount = 0
                    If insSQL IsNot Nothing Then
                        insSQL.Clear()
                        insSQL = Nothing
                    End If
                    If insValue IsNot Nothing Then
                        insValue.Clear()
                        insValue = Nothing
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
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If tbMaster IsNot Nothing Then
                    tbMaster.Dispose()
                    tbMaster = Nothing
                End If
                If tbDetail IsNot Nothing Then
                    tbDetail.Dispose()
                    tbDetail = Nothing
                End If
                If tbSechema IsNot Nothing Then
                    tbSechema.Dispose()
                    tbSechema = Nothing
                End If
                If tbChildSchema IsNot Nothing Then
                    tbChildSchema.Dispose()
                    tbChildSchema = Nothing
                End If
                If tbChildSO1109B IsNot Nothing Then
                    tbChildSO1109B.Dispose()
                    tbChildSO1109B = Nothing
                End If

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
