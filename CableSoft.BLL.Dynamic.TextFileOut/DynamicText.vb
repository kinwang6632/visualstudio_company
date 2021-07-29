Imports CableSoft.BLL.Utility
Imports System.Web
Imports System.Xml
Imports System.Data.Common

Public Class DynamicText
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New DynamicTextDALMultiDB(Me.LoginInfo.Provider)
    Private Const ErrorString As String = "Error"
    Private Const ErrorCaption As String = "ErrorCaption"
    Private Const ErrorLog As String = "ErrorLog"
    Private Const WriteDataName As String = "RetData"
    Private Const tbMasterName As String = "Master"
    Private Const tbDetailName As String = "Detail"
    Private Const TxtDirName As String = "TXT"
    Private tbMaster As DataTable = Nothing
    Private tbDetail As DataTable = Nothing
    Private FShouldAmt As Integer = 0
    Private FRecordCount As Integer = 0
    Private FSuccessCount As Integer = 0
    Private FFailCount As Integer = 0
    Private Language As New CableSoft.BLL.Language.SO61.DynamicTextLanguage
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
    '2017/01/09 Jacky 增加一個一次取回設定檔的method
    Public Function GetSettingData(ByVal SysProgramId As String) As RIAResult
        Using RetData As DataSet = New DataSet()
            Try
                Using CompCode As DataTable = GetCompCode()
                    CompCode.TableName = "CanChooseComp"
                    RetData.Merge(CompCode)
                End Using
            Catch ex As Exception
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Get CanChooseComp Error:" & ex.ToString()}
            End Try
            Try
                '取電子檔出帳設定檔
                Dim tData As DataSet = QueryDynTextOut(SysProgramId)
                RetData.Merge(tData)
            Catch ex As Exception
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Get TextFileOut Error:" & ex.ToString()}
            End Try
            Try
                '取動態條件
                Using bll As New CableSoft.BLL.Dynamic.Condition.DynamicCondition(LoginInfo, DAO)
                    Using CondSetting As DataSet = bll.GetConditionField(Nothing, RetData.Tables("OutMaster").Rows(0).Item("CondProgId"))
                        RetData.Merge(CondSetting)
                    End Using
                End Using
            Catch ex As Exception
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Get DynamicCondtion Error:" & ex.ToString()}
            End Try
            Return New RIAResult() With {.ResultBoolean = True, .ResultDataSet = RetData}
        End Using
    End Function
    Public Function QueryDynCondition(ByVal SysProgramId As String) As DataSet
        Dim objDynamicUpdate As New DynamicUpdate.DynamicUpdate(Me.LoginInfo, Me.DAO)
        Try
            'Return objDynamicUpdate.QueryEnvironment(SysProgramId)
            Dim ds As New DataSet
            Dim ErrMsg As String = Nothing
            tbMaster = DAO.ExecQry(_DAL.QuerySO1101A, New Object() {SysProgramId})

            If tbMaster Is Nothing OrElse tbMaster.Rows.Count = 0 Then
                Throw New Exception(String.Format(Language.NoSO1101AData, SysProgramId))
            End If
            tbMaster.TableName = tbMasterName
            tbDetail = DAO.ExecQry(_DAL.QuerySO1101B, New Object() {tbMaster.Rows(0).Item("ProgramId")})
            If tbDetail Is Nothing OrElse tbDetail.Rows.Count = 0 Then
                Throw New Exception(String.Format(Language.NoSO1101BData, tbMaster.Rows(0).Item("ProgramId")))
            End If
            tbDetail.TableName = tbDetailName
            'ErrMsg = chkSO1109A()
            If Not String.IsNullOrEmpty(ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            'ErrMsg = chkSO1109B()
            If Not String.IsNullOrEmpty(ErrMsg) Then
                Throw New Exception(ErrMsg)
            End If
            ds.Tables.Add(tbMaster.Copy)
            ds.Tables.Add(tbDetail.Copy)
            Return ds
        Catch ex As Exception
            Throw ex
        Finally
            objDynamicUpdate.Dispose()
        End Try
    End Function
    'Private Function chkSO1109A() As String
    '    If DBNull.Value.Equals(tbMaster.Rows(0).Item("TableName")) Then
    '        Return DynamicUpdateLanguage.NoTableName
    '    End If
    '    If DBNull.Value.Equals(tbMaster.Rows(0).Item("CondProgId")) Then
    '        Return DynamicUpdateLanguage.NoCondProgIdProp
    '    End If
    '    Return Nothing
    'End Function
    
    Public Function QueryDynTextOut(ByVal SysProgramId As String) As DataSet
        Dim ds As New DataSet
        Dim tbMaster As DataTable = DAO.ExecQry(_DAL.QueryMaster, New Object() {SysProgramId})
        Dim aProgramId As String = "X"
        If (tbMaster IsNot Nothing) AndAlso (tbMaster.Rows.Count > 0) Then
            aProgramId = tbMaster.Rows(0).Item("ProgramId")
        End If
        Dim tbDetail As DataTable = DAO.ExecQry(_DAL.QueryDetail, New Object() {aProgramId})
        

        Try
            tbMaster.TableName = "OutMaster"
            tbDetail.TableName = "OutDetail"
            ds.Tables.Add(tbMaster.Copy)
            ds.Tables.Add(tbDetail.Copy)
            Return ds
        Catch ex As Exception
            Throw ex
        Finally
            ds.Dispose()
        End Try

    End Function
    Private Function UpdateData(ByVal updQuery As String) As Boolean
        Dim trans As DbTransaction = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim blnAutoClose As Boolean = False
        Try

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

            'Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
            '    cmd.Connection = cn
            '    cmd.Transaction = trans
            '    cmd.CommandText = updQuery
            '    cmd.ExecuteNonQuery()
            'End Using
            DAO.ExecNqry(updQuery)
        Catch ex As Exception
            trans.Rollback()
            Throw ex
        Finally           
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
            End If
        End Try
        Return True
    End Function
    Private Sub BefUpdateData(ByVal lstUpd As List(Of String), ByRef cn As DbConnection, ByRef trans As DbTransaction)

        'Dim trans As DbTransaction = Nothing
        'Dim cn As DbConnection = DAO.GetConn()
        'Dim blnAutoClose As Boolean = False
        'If DAO.Transaction IsNot Nothing Then
        '    trans = DAO.Transaction
        'Else
        '    If cn.State = ConnectionState.Closed Then
        '        cn.ConnectionString = Me.LoginInfo.ConnectionString
        '        cn.Open()
        '    End If
        '    trans = cn.BeginTransaction
        '    DAO.Transaction = trans
        '    blnAutoClose = True
        'End If
        'DAO.AutoCloseConn = False
        If cn.State <> ConnectionState.Open Then
            cn.Open()
        End If
        Try
            If lstUpd Is Nothing OrElse lstUpd.Count = 0 Then
                Exit Sub
            End If
            For Each upd As String In lstUpd
                'Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                '    cmd.Connection = cn
                '    cmd.Transaction = trans
                '    cmd.CommandText = upd
                '    cmd.ExecuteNonQuery()
                'End Using
                DAO.ExecNqry(upd)
            Next
            'trans.Commit()
        Catch ex As Exception
            Throw
        Finally
            'If cn IsNot Nothing Then
            '    cn.Close()
            'End If
            'If blnAutoClose Then
            '    If trans IsNot Nothing Then
            '        trans.Dispose()
            '    End If
            '    If cn IsNot Nothing Then
            '        cn.Close()
            '        cn.Dispose()
            '    End If
            'If blnAutoClose Then
            '    DAO.AutoCloseConn = True
            'End If
            'End If
        End Try
    End Sub
    Private Function WriteText(ByVal tbDetail As DataTable, ByVal filePath As String, ByVal tbSource As DataTable, _
                               ByVal isGetway As Boolean, _
                               ByVal lstUpdSQL As List(Of String), _
                              ByRef cn As DbConnection, ByRef trans As DbTransaction) As String

        'Dim Path As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & TxtDirName
        If Right(filePath, 1) = "\" Then filePath = filePath.Substring(0, filePath.Length - 1)
        filePath = filePath & "\"
        Dim Path As String = filePath & TxtDirName
        Dim txtFileName As String = tbDetail.Rows(0).Item("DataText").ToString
        Dim errFileName As String = tbDetail.Rows(0).Item("ErrorText").ToString
        Dim aryTxtFile As New Dictionary(Of String, System.Text.StringBuilder)(StringComparer.OrdinalIgnoreCase)
        Dim retFileName As String = Nothing
        Dim colAry As New Dictionary(Of String, String)

        Dim sbdError As New System.Text.StringBuilder
        Dim aUpdateKeyField As String = Nothing
        Dim aBatchKeyField As String = Nothing
        Dim haveTextFileNameField As Boolean = False
        FShouldAmt = 0
        If Not System.IO.Directory.Exists(Path) Then
            System.IO.Directory.CreateDirectory(Path)
        End If
        If tbSource.Columns.Contains("TEXTFILENAME") Then
            haveTextFileNameField = True
            If (tbSource.Rows.Count > 0) AndAlso (Not DBNull.Value.Equals(tbSource.Rows(0).Item("TEXTFILENAME"))) Then
                'txtFileName = tbSource.Rows(0).Item("TEXTFILENAME")
                'Dim lstrw As List(Of DataRow) = tbSource.AsEnumerable.Distinct.Where(Function(rwDistinct As DataRow)
                '                                                                         Return Not DBNull.Value.Equals(rwDistinct.Item("TEXTFILENAME"))
                '                                                                     End Function)
            End If
        End If
        aryTxtFile.Add(txtFileName, New System.Text.StringBuilder)
        If Not DBNull.Value.Equals(tbDetail.Rows(0).Item("UpdateKeyField")) Then
            aUpdateKeyField = tbDetail.Rows(0).Item("UpdateKeyField").ToString
        End If
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If
        '將錯誤欄位選出來
        For Each col As DataColumn In tbSource.Columns
            If col.ColumnName.ToUpper <> "ErrorLog".ToUpper Then
                If (col.ColumnName.ToUpper.Substring(0, ErrorString.Length) = ErrorString.ToUpper) Then
                    If (col.ColumnName.Length > ErrorCaption.Length) Then
                        If col.ColumnName.Substring(0, ErrorCaption.Length).ToUpper <> ErrorCaption.ToUpper Then
                            colAry.Add(col.ColumnName, ErrorCaption & col.ColumnName.ToUpper.Substring(ErrorString.Length))
                        End If
                    Else
                        colAry.Add(col.ColumnName, ErrorCaption & col.ColumnName.ToUpper.Substring(ErrorString.Length))
                    End If
                End If
            End If
        Next
        Try
            Dim haveError As Boolean = False
            Dim haveWriteErrLog As Boolean = False
            FRecordCount = 0
            FFailCount = 0
            FShouldAmt = 0
            FSuccessCount = 0
            For Each rw As DataRow In tbSource.Rows
                If Int32.Parse("0" & rw.Item("DataRowType").ToString) = 0 Then
                    FRecordCount = FRecordCount + 1
                End If
                '先判斷是否有錯誤資料
                haveError = False
                haveWriteErrLog = False
                For i As Int32 = 0 To colAry.Keys.Count - 1
                    If (Not DBNull.Value.Equals(rw.Item(colAry.Keys(i)))) AndAlso
                        (Not DBNull.Value.Equals(rw.Item(colAry.Values(i)))) Then
                        If IsNumeric(rw.Item(colAry.Keys(i))) Then
                            If Integer.Parse(rw.Item(colAry.Keys(i))) > 0 Then
                                If Integer.Parse(rw.Item(colAry.Keys(i))) = 1 Then
                                    haveError = True
                                    sbdError.Append(rw.Item(colAry.Values(i)))
                                Else
                                    haveWriteErrLog = True
                                    sbdError.Append(rw.Item(colAry.Values(i)))
                                End If
                            End If
                        End If
                    End If
                Next
                If haveError Then
                    If Not DBNull.Value.Equals(rw.Item(ErrorLog)) Then
                        sbdError.Append(":" & rw.Item(ErrorLog))
                        If haveError AndAlso Integer.Parse("0" & rw.Item("DataRowType").ToString) = 0 Then
                            FFailCount = FFailCount + 1
                        End If
                    End If
                    If Not String.IsNullOrEmpty(sbdError.ToString) Then
                        sbdError.Append(Environment.NewLine)
                    End If
                    '將所有錯誤的資料依欄位挑出來,避免批次更新到錯誤的資料
                    If lstUpdSQL.Count > 0 Then
                        If Not String.IsNullOrEmpty(aUpdateKeyField) Then
                            If Not String.IsNullOrEmpty(aBatchKeyField) Then
                                aBatchKeyField = String.Format("{0},'{1}'", aBatchKeyField, rw.Item(aUpdateKeyField))
                            Else
                                aBatchKeyField = String.Format("'{0}'", rw.Item(aUpdateKeyField))
                            End If
                        End If
                    End If
                Else
                    If haveWriteErrLog Then
                        If Not DBNull.Value.Equals(rw.Item(ErrorLog)) Then
                            sbdError.Append(":" & rw.Item(ErrorLog))
                        End If
                        If Not String.IsNullOrEmpty(sbdError.ToString) Then
                            sbdError.Append(Environment.NewLine)
                        End If
                    End If

                    If haveTextFileNameField Then
                        If (Not DBNull.Value.Equals(rw.Item("TEXTFILENAME"))) Then
                            If Not aryTxtFile.ContainsKey((rw.Item("TEXTFILENAME").ToString)) Then
                                aryTxtFile.Add(rw.Item("TEXTFILENAME").ToString, New System.Text.StringBuilder)
                            End If
                            aryTxtFile.Item(rw.Item("TEXTFILENAME").ToString).AppendLine(rw.Item(WriteDataName))
                        Else
                            aryTxtFile.Item(txtFileName).AppendLine(rw.Item(WriteDataName))
                        End If
                    Else
                        aryTxtFile.Item(txtFileName).AppendLine(rw.Item(WriteDataName))
                    End If


                    If tbSource.Columns.Contains("ShouldAmt") Then
                        If (Not DBNull.Value.Equals(rw.Item("ShouldAmt"))) AndAlso
                            (Integer.Parse(rw.Item("ShouldAmt")) <> 0) Then
                            FShouldAmt = FShouldAmt + rw.Item("ShouldAmt")
                        End If
                    End If
                End If
                '單筆更新
                If (Integer.Parse("0" & tbDetail.Rows(0).Item("UpdateType")) = 0) AndAlso
                        (Not DBNull.Value.Equals(tbDetail.Rows(0).Item("UpdateSQL"))) Then
                    '沒錯誤才更新
                    If (Not haveError) AndAlso
                        (Integer.Parse("0" & rw.Item("DataRowType").ToString) = 0) Then
                        Dim params() As Object = Nothing
                        Try

                            For Each updQuery As String In lstUpdSQL
                                If params IsNot Nothing Then
                                    Array.Clear(params, 0, params.Length)
                                End If

                                For Each col As DataColumn In tbSource.Columns
                                    updQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, updQuery, col.ColumnName, rw.Item(col.ColumnName), params)
                                Next
                                If updQuery.ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                                    For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                                        Dim FieldName = PropertyInfo.Name
                                        If PropertyInfo IsNot Nothing Then
                                            updQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, updQuery, "LoginInfo." & FieldName, GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                                        End If
                                    Next
                                End If

                                updQuery = ReplaceLoginInfoWhere(updQuery, params)

                                'Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                                '    cmd.Connection = cn
                                '    cmd.Transaction = trans
                                '    cmd.CommandText = updQuery
                                '    cmd.ExecuteNonQuery()
                                'End Using
                                DAO.ExecNqry(updQuery)
                            Next

                        Catch ex As Exception
                            trans.Rollback()
                            Throw ex
                            Exit For
                        End Try
                    End If
                End If
            Next
            FSuccessCount = FRecordCount - FFailCount

            Dim arybtyTxt As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)
            Dim btyErr() As Byte = System.Text.Encoding.GetEncoding(950).GetBytes(sbdError.ToString)
            If Not isGetway Then
                retFileName = Me.LoginInfo.EntryId & "-" & Now.ToString("yyyyMMddHHmmssff") & ".zip"
            Else
                retFileName = "Getway-" & Me.LoginInfo.EntryId & "-" & Now.ToString("yyyyMMddHHmmssff") & ".zip"
            End If
            For i As Int32 = 0 To aryTxtFile.Keys.Count - 1
                If Not String.IsNullOrEmpty(aryTxtFile.Item(aryTxtFile.Keys(i)).ToString) Then
                    arybtyTxt.Add(aryTxtFile.Keys(i).ToString,
                                  System.Text.Encoding.GetEncoding(950).GetBytes(aryTxtFile.Item(aryTxtFile.Keys(i)).ToString))
                End If
            Next
            Using zip As New Ionic.Zip.ZipFile(Path & "\" & retFileName, System.Text.Encoding.GetEncoding(950))

                For i As Integer = 0 To arybtyTxt.Keys.Count - 1
                    zip.AddEntry(arybtyTxt.Keys(i).ToString,
                                 arybtyTxt.Item(arybtyTxt.Keys(i).ToString))
                Next
                If btyErr.Count > 0 Then
                    zip.AddEntry(errFileName, btyErr)
                End If
                zip.Save()
            End Using
            '批次更新
            If (lstUpdSQL.Count > 0) AndAlso (Integer.Parse("0" & tbDetail.Rows(0).Item("UpdateType")) = 1) Then
                Dim params() As Object = Nothing
                For Each updSQL As String In lstUpdSQL

                    If (Not String.IsNullOrEmpty(aBatchKeyField)) AndAlso (Not String.IsNullOrEmpty(aBatchKeyField)) Then
                        updSQL = String.Format("{0} AND {1} NOT IN ({2})", updSQL, aUpdateKeyField, aBatchKeyField)
                    End If
                    'Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                    '    cmd.Connection = cn
                    '    cmd.Transaction = trans
                    '    cmd.CommandText = updSQL
                    '    cmd.ExecuteNonQuery()
                    'End Using
                    DAO.ExecNqry(updSQL)
                Next

            End If
            'trans.Commit()

            Erase btyErr
        Catch ex As Exception
            Throw
        Finally
            sbdError.Clear()
            aryTxtFile.Clear()

        End Try
        '2017/01/09 Jacky 加傳目錄位置
        Return String.Format("{0}\{1}", TxtDirName, retFileName)
    End Function
    Private Function WriteText(ByVal tbDetail As DataTable, tbSource As DataTable, ByVal lstUpdSQL As List(Of String), ByVal isGetway As Boolean, _
                               ByRef cn As DbConnection, ByRef trans As DbTransaction) As String
        '2017/01/09 Jacky 改取絕對位置
        Return WriteText(tbDetail, CableSoft.BLL.Utility.Utility.GetCurrentDirectory(), tbSource, isGetway, lstUpdSQL, cn, trans)
        Dim Path As String = CableSoft.BLL.Utility.Utility.GetCurrentDirectory() & TxtDirName

        Dim txtFileName As String = tbDetail.Rows(0).Item("DataText").ToString
        Dim errFileName As String = tbDetail.Rows(0).Item("ErrorText").ToString
        Dim aryTxtFile As New Dictionary(Of String, System.Text.StringBuilder)(StringComparer.OrdinalIgnoreCase)
        Dim retFileName As String = Nothing
        Dim colAry As New Dictionary(Of String, String)

        Dim sbdError As New System.Text.StringBuilder
        Dim aUpdateKeyField As String = Nothing
        Dim aBatchKeyField As String = Nothing
        Dim haveTextFileNameField As Boolean = False
        FShouldAmt = 0
        If Not System.IO.Directory.Exists(Path) Then
            System.IO.Directory.CreateDirectory(Path)
        End If
        If tbSource.Columns.Contains("TEXTFILENAME") Then
            haveTextFileNameField = True
            If (tbSource.Rows.Count > 0) AndAlso (Not DBNull.Value.Equals(tbSource.Rows(0).Item("TEXTFILENAME"))) Then
                'txtFileName = tbSource.Rows(0).Item("TEXTFILENAME")
                'Dim lstrw As List(Of DataRow) = tbSource.AsEnumerable.Distinct.Where(Function(rwDistinct As DataRow)
                '                                                                         Return Not DBNull.Value.Equals(rwDistinct.Item("TEXTFILENAME"))
                '                                                                     End Function)
            End If
        End If
        aryTxtFile.Add(txtFileName, New System.Text.StringBuilder)
        If Not DBNull.Value.Equals(tbDetail.Rows(0).Item("UpdateKeyField")) Then
            aUpdateKeyField = tbDetail.Rows(0).Item("UpdateKeyField").ToString
        End If
        If cn.State = ConnectionState.Closed Then
            cn.Open()
        End If
        '將錯誤欄位選出來
        For Each col As DataColumn In tbSource.Columns
            If col.ColumnName.ToUpper <> "ErrorLog".ToUpper Then
                If (col.ColumnName.ToUpper.Substring(0, ErrorString.Length) = ErrorString.ToUpper) Then
                    If (col.ColumnName.Length > ErrorCaption.Length) Then
                        If col.ColumnName.Substring(0, ErrorCaption.Length).ToUpper <> ErrorCaption.ToUpper Then
                            colAry.Add(col.ColumnName, ErrorCaption & col.ColumnName.ToUpper.Substring(ErrorString.Length))
                        End If
                    Else
                        colAry.Add(col.ColumnName, ErrorCaption & col.ColumnName.ToUpper.Substring(ErrorString.Length))
                    End If
                End If
            End If
        Next
        Try
            Dim haveError As Boolean = False
            Dim haveWriteErrLog As Boolean = False
            FRecordCount = 0
            FFailCount = 0
            FShouldAmt = 0
            FSuccessCount = 0
            For Each rw As DataRow In tbSource.Rows
                If Int32.Parse("0" & rw.Item("DataRowType").ToString) = 0 Then
                    FRecordCount = FRecordCount + 1
                End If
                '先判斷是否有錯誤資料
                haveError = False
                haveWriteErrLog = False
                For i As Int32 = 0 To colAry.Keys.Count - 1
                    If (Not DBNull.Value.Equals(rw.Item(colAry.Keys(i)))) AndAlso
                        (Not DBNull.Value.Equals(rw.Item(colAry.Values(i)))) Then
                        If IsNumeric(rw.Item(colAry.Keys(i))) Then
                            If Integer.Parse(rw.Item(colAry.Keys(i))) > 0 Then
                                If Integer.Parse(rw.Item(colAry.Keys(i))) = 1 Then
                                    haveError = True
                                    sbdError.Append(rw.Item(colAry.Values(i)))
                                Else
                                    haveWriteErrLog = True
                                    sbdError.Append(rw.Item(colAry.Values(i)))
                                End If
                            End If
                        End If
                    End If
                Next
                If haveError Then
                    If Not DBNull.Value.Equals(rw.Item(ErrorLog)) Then
                        sbdError.Append(":" & rw.Item(ErrorLog))
                        If haveError Then
                            FFailCount = FFailCount + 1
                        End If
                    End If
                    If Not String.IsNullOrEmpty(sbdError.ToString) Then
                        sbdError.Append(Environment.NewLine)
                    End If
                    '將所有錯誤的資料依欄位挑出來,避免批次更新到錯誤的資料
                    If lstUpdSQL.Count > 0 Then
                        If Not String.IsNullOrEmpty(aUpdateKeyField) Then
                            If Not String.IsNullOrEmpty(aBatchKeyField) Then
                                aBatchKeyField = String.Format("{0},'{1}'", aBatchKeyField, rw.Item(aUpdateKeyField))
                            Else
                                aBatchKeyField = String.Format("'{0}'", rw.Item(aUpdateKeyField))
                            End If
                        End If
                    End If
                Else
                    If haveWriteErrLog Then
                        If Not DBNull.Value.Equals(rw.Item(ErrorLog)) Then
                            sbdError.Append(":" & rw.Item(ErrorLog))
                        End If
                        If Not String.IsNullOrEmpty(sbdError.ToString) Then
                            sbdError.Append(Environment.NewLine)
                        End If
                    End If

                    If haveTextFileNameField Then
                        If (Not DBNull.Value.Equals(rw.Item("TEXTFILENAME"))) Then
                            If Not aryTxtFile.ContainsKey((rw.Item("TEXTFILENAME").ToString)) Then
                                aryTxtFile.Add(rw.Item("TEXTFILENAME").ToString, New System.Text.StringBuilder)
                            End If
                            aryTxtFile.Item(rw.Item("TEXTFILENAME").ToString).AppendLine(rw.Item(WriteDataName))
                        Else
                            aryTxtFile.Item(txtFileName).AppendLine(rw.Item(WriteDataName))
                        End If
                    Else
                        aryTxtFile.Item(txtFileName).AppendLine(rw.Item(WriteDataName))
                    End If


                    If tbSource.Columns.Contains("ShouldAmt") Then
                        If (Not DBNull.Value.Equals(rw.Item("ShouldAmt"))) AndAlso
                            (Integer.Parse(rw.Item("ShouldAmt")) <> 0) Then
                            FShouldAmt = FShouldAmt + rw.Item("ShouldAmt")
                        End If
                    End If
                End If
                '單筆更新
                If (Integer.Parse("0" & tbDetail.Rows(0).Item("UpdateType")) = 0) AndAlso
                        (Not DBNull.Value.Equals(tbDetail.Rows(0).Item("UpdateSQL"))) Then
                    '沒錯誤才更新
                    If (Not haveError) AndAlso
                        (Integer.Parse("0" & rw.Item("DataRowType").ToString) = 0) Then
                        Dim params() As Object = Nothing
                        Try
                            'Dim updQuery As String = Nothing
                            'updQuery = tbDetail.Rows(0).Item("UpdateSQL").ToString
                            'If params IsNot Nothing Then
                            '    Array.Clear(params, 0, params.Length)
                            'End If
                            ''updQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, updQuery, "RetData", rw.Item("CodeNo"), params)
                            'For Each col As DataColumn In tbSource.Columns
                            '    updQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, updQuery, col.ColumnName, rw.Item(col.ColumnName), params)
                            'Next
                            'updQuery = ReplaceLoginInfoWhere(updQuery, params)
                            For Each updQuery As String In lstUpdSQL
                                If params IsNot Nothing Then
                                    Array.Clear(params, 0, params.Length)
                                End If

                                For Each col As DataColumn In tbSource.Columns
                                    updQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, updQuery,
                                                                                              col.ColumnName, rw.Item(col.ColumnName), params)
                                Next
                                If updQuery.ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                                    For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                                        Dim FieldName = PropertyInfo.Name
                                        If PropertyInfo IsNot Nothing Then
                                            updQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, updQuery,
                                                                                                      "LoginInfo." & FieldName,
                                                                                                      GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                                        End If
                                    Next
                                End If
                                updQuery = ReplaceLoginInfoWhere(updQuery, params)
                                'Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                                '    cmd.Connection = cn
                                '    cmd.Transaction = trans
                                '    cmd.CommandText = updQuery
                                '    cmd.ExecuteNonQuery()
                                'End Using
                                DAO.ExecNqry(updQuery)
                            Next

                        Catch ex As Exception
                            trans.Rollback()
                            Throw ex
                            Exit For
                        End Try
                    End If
                End If
            Next
            FSuccessCount = FRecordCount - FFailCount

            Dim arybtyTxt As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)
            Dim btyErr() As Byte = System.Text.Encoding.GetEncoding(950).GetBytes(sbdError.ToString)
            retFileName = Me.LoginInfo.EntryId & "-" & Now.ToString("yyyyMMddHHmmssff") & ".zip"
            For i As Int32 = 0 To aryTxtFile.Keys.Count - 1
                If Not String.IsNullOrEmpty(aryTxtFile.Item(aryTxtFile.Keys(i)).ToString) Then
                    arybtyTxt.Add(aryTxtFile.Keys(i).ToString,
                                  System.Text.Encoding.GetEncoding(950).GetBytes(aryTxtFile.Item(aryTxtFile.Keys(i)).ToString))
                End If
            Next
            Using zip As New Ionic.Zip.ZipFile(Path & "\" & retFileName, System.Text.Encoding.GetEncoding(950))

                For i As Integer = 0 To arybtyTxt.Keys.Count - 1
                    zip.AddEntry(arybtyTxt.Keys(i).ToString,
                                 arybtyTxt.Item(arybtyTxt.Keys(i).ToString))
                Next
                If btyErr.Count > 0 Then
                    zip.AddEntry(errFileName, btyErr)
                End If
                zip.Save()
            End Using
            '批次更新
            If (lstUpdSQL.Count > 0) AndAlso (Integer.Parse("0" & tbDetail.Rows(0).Item("UpdateType")) = 1) Then
                Dim params() As Object = Nothing
                For Each updSQL As String In lstUpdSQL

                    If (Not String.IsNullOrEmpty(aBatchKeyField)) AndAlso (Not String.IsNullOrEmpty(aBatchKeyField)) Then
                        updSQL = String.Format("{0} AND {1} NOT IN ({2})", updSQL, aUpdateKeyField, aBatchKeyField)
                    End If
                    'Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                    '    cmd.Connection = cn
                    '    cmd.Transaction = trans
                    '    cmd.CommandText = updSQL
                    '    cmd.ExecuteNonQuery()
                    'End Using
                    DAO.ExecNqry(updSQL)
                Next

            End If
            'trans.Commit()

            Erase btyErr
        Catch ex As Exception
            Throw
        Finally
            sbdError.Clear()
            aryTxtFile.Clear()
            'If blnAutoClose Then
            '    If trans IsNot Nothing Then
            '        trans.Dispose()
            '    End If
            '    If cn IsNot Nothing Then
            '        cn.Close()
            '        cn.Dispose()
            '    End If
            '    If blnAutoClose Then
            '        DAO.AutoCloseConn = True
            '    End If
            'End If
        End Try
        '2017/01/09 Jacky 加傳目錄位置
        Return String.Format("{0}\{1}", TxtDirName, retFileName)
    End Function
    'Public Function GetCompCode() As DataTable
    '    If Me.LoginInfo.GroupId = "0" AndAlso 1 = 0 Then
    '        Return DAO.ExecQry(_DAL.GetCompCode("0"))
    '    Else
    '        Return DAO.ExecQry(_DAL.GetCompCode(Me.LoginInfo.GroupId), New Object() {Me.LoginInfo.EntryId})
    '    End If
    'End Function

    Public Function GetCompCode() As DataTable
        Try
            'Return DAO.ExecQry(_DAL.GetCompCode("0", CableSoft.BLL.Utility.Utility.GetCompanyTableName(Me.LoginInfo, Me.DAO), Nothing))

            Return DAO.ExecQry(_DAL.GetCompCode("1",
                                                 CableSoft.BLL.Utility.Utility.GetCompanyTableName(Me.LoginInfo, Me.DAO),
                                                    CableSoft.BLL.Utility.Utility.GetLoginTableName),
                                New Object() {Me.LoginInfo.EntryId})
        Catch ex As Exception
            Throw
        End Try


        'Try
        '    If Me.LoginInfo.GroupId = "0" Then
        '        Return DAO.ExecQry(_DAL.GetCompCode("0",
        '                                            CableSoft.BLL.Utility.Utility.GetCompanyTableName(Me.LoginInfo, Me.DAO),
        '                                               CableSoft.BLL.Utility.Utility.GetLoginTableName))
        '    Else
        '        Return DAO.ExecQry(_DAL.GetCompCode("1",
        '                                            CableSoft.BLL.Utility.Utility.GetCompanyTableName(Me.LoginInfo, Me.DAO),
        '                                               CableSoft.BLL.Utility.Utility.GetLoginTableName),
        '                           New Object() {Me.LoginInfo.EntryId})
        '    End If
        'Catch ex As Exception
        '    Throw
        'End Try

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
    Public Function ChkAuthority(ByVal SysProgramId As String) As RIAResult
        Dim result As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        Dim tbMaster As DataTable = DAO.ExecQry(_DAL.QueryMaster, New Object() {SysProgramId})
        Try
            If tbMaster.Rows.Count = 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -3
                result.ErrorMessage = Language.NoFundMaster
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
                    result.ErrorMessage = Language.NoPermission
                End If
                'If Integer.Parse(DAO.ExecSclr(_DAL.chkAuthority(Me.LoginInfo.GroupId), New Object() {tbMaster.Rows(0).Item("MID")})) = 0 Then
                '    result.ResultBoolean = False
                '    result.ErrorCode = -1
                '    result.ErrorMessage = Language.NoPermission
                '    Return result
                'End If
            End If
            Return result
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
    Public Function InsertResv(ByVal SysProgramId As String, ByVal AutoSerialNo As Integer, ByVal ResvTime As Date, _
                            ByVal dsConditions As DataSet)
        Dim trans As DbTransaction = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim blnAutoClose As Boolean = False
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

    End Function
    Public Function Execute(ByVal SysProgramId As String, ByVal AutoSerialNo As Integer, ByVal isGetway As Boolean, ByVal SEQNO As Long, _
                            ByVal dsConditions As DataSet) As String
        Dim programID As String = Nothing
        Dim condProgID As String = Nothing
        Dim sqlQuery As String = Nothing
        Dim dynaCdt As CableSoft.BLL.Dynamic.Condition.DynamicCondition = Nothing
        Dim tbMaster As DataTable = DAO.ExecQry(_DAL.QueryMaster, New Object() {SysProgramId})
        Dim tbSingleDetail As DataTable = DAO.ExecQry(_DAL.QuerySingleDetail, New Object() {AutoSerialNo})
        Dim dtReturn As DataTable = Nothing
        Dim dtCdt As DataTable = Nothing
        Dim cdtProgId As String = Nothing
        Dim retFileName As String = Nothing
        Dim updSQL As String = Nothing
        Dim BeforeUpd As New List(Of String)
        Dim UpdateSQL As New List(Of String)
        Dim FinalSQL As New List(Of String)
        Dim trans As DbTransaction = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim blnAutoClose As Boolean = False
        Dim RunTime As New Stopwatch()
        Dim logSQL As New System.Text.StringBuilder()

        Dim BllUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim ParentSeqNo As Double = -1
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
        CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, tbMaster.Rows(0).Item("CAPTION"))
        Try
            programID = tbMaster.Rows(0).Item("PROGRAMID").ToString
            condProgID = tbMaster.Rows(0).Item("CondProgId").ToString
            cdtProgId = DAO.ExecSclr(_DAL.QueryDynProgId, New Object() {condProgID})
            sqlQuery = tbSingleDetail.Rows(0).Item("SQLQUERY").ToString
            dtCdt = dsConditions.Tables("Condition")
            dynaCdt = New CableSoft.BLL.Dynamic.Condition.DynamicCondition(Me.LoginInfo, Me.DAO)
            Dim params() As Object = Nothing
            dtReturn = dynaCdt.GetBuildConditionSQL(cdtProgId, dtCdt, params)
            If Not DBNull.Value.Equals(tbSingleDetail.Rows(0).Item("BEFUPDSQL")) Then
                BeforeUpd = tbSingleDetail.Rows(0).Item("BEFUPDSQL").ToString.Split(";").ToList
            End If
            If Not DBNull.Value.Equals(tbSingleDetail.Rows(0).Item("UpdateSQL")) Then
                UpdateSQL = tbSingleDetail.Rows(0).Item("UpdateSQL").ToString.Split(";").ToList
            End If
            If Not DBNull.Value.Equals(tbSingleDetail.Rows(0).Item("FinalSQL")) Then
                FinalSQL = tbSingleDetail.Rows(0).Item("FinalSQL").ToString.Split(";").ToList
            End If
            Dim aFieldName As String = Nothing
            For Each dr As DataRow In dtReturn.Rows
                sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery, dr("FieldName"), dr("ConditionSQL"), params)
                If sqlQuery.ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                    For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                        Dim FieldName = PropertyInfo.Name
                        If PropertyInfo IsNot Nothing Then
                            sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery,
                                                                                      "LoginInfo." & FieldName,
                                                                                      GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                        End If
                    Next
                End If
                For int As Integer = 0 To BeforeUpd.Count - 1
                    BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         BeforeUpd.Item(int), dr("FieldName"), dr("ConditionSQL"), params)
                    If BeforeUpd.Item(int).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, BeforeUpd.Item(int),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                Next


                For index As Integer = 0 To UpdateSQL.Count - 1
                    UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         UpdateSQL.Item(index), dr("FieldName"), dr("ConditionSQL"), params)
                    If UpdateSQL.Item(index).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, UpdateSQL.Item(index),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                Next

                For inxFinal As Integer = 0 To FinalSQL.Count - 1
                    FinalSQL.Item(inxFinal) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         FinalSQL.Item(inxFinal), dr("FieldName"), dr("ConditionSQL"), params)
                    If FinalSQL.Item(inxFinal).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                FinalSQL.Item(inxFinal) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, FinalSQL.Item(inxFinal),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                    FinalSQL.Item(inxFinal) = ReplaceLoginInfoWhere(FinalSQL.Item(inxFinal), params)
                Next

            Next


            For i As Integer = 0 To dtCdt.Rows.Count - 1
                '2017/01/09 Jacky 增加取 _1 的資料
                If Right(dtCdt.Rows(i)("FieldName"), 2) = "_0" OrElse Right(dtCdt.Rows(i)("FieldName"), 2) = "_1" Then
                    aFieldName = dtCdt.Rows(i)("FieldName").ToString.Substring(0,
                                                                               dtCdt.Rows(i)("FieldName").ToString.Length - 2)
                Else
                    aFieldName = dtCdt.Rows(i)("FieldName").ToString
                End If

                sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery, aFieldName, dtCdt.Rows(i)("FieldValue"), params)
                If sqlQuery.ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                    For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                        Dim FieldName = PropertyInfo.Name
                        If PropertyInfo IsNot Nothing Then
                            sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery,
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                        End If
                    Next
                End If

                For int As Integer = 0 To BeforeUpd.Count - 1
                    BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         BeforeUpd.Item(int), aFieldName, dtCdt.Rows(i)("FieldValue"), params)
                    If BeforeUpd.Item(int).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, BeforeUpd.Item(int),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                    BeforeUpd.Item(int) = ReplaceLoginInfoWhere(BeforeUpd.Item(int), params)
                Next
                For index As Integer = 0 To UpdateSQL.Count - 1
                    UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         UpdateSQL.Item(index), aFieldName, dtCdt.Rows(i)("FieldValue"), params)
                    If UpdateSQL.Item(index).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, UpdateSQL.Item(index),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                    UpdateSQL.Item(index) = ReplaceLoginInfoWhere(UpdateSQL.Item(index), params)
                Next

            Next
            'log sql history
            '--------------------------------------------------------------------------
            For index As Integer = 0 To BeforeUpd.Count - 1
                If index = 0 Then logSQL.Append("Before:")
                logSQL.Append(Environment.NewLine & BeforeUpd.Item(index) & ";")
            Next
            If logSQL.ToString IsNot Nothing Then
                logSQL.Append(Environment.NewLine)
            End If
            logSQL.Append("Query:")
            logSQL.Append(Environment.NewLine & sqlQuery & ";")
            If UpdateSQL.Count > 0 Then
                logSQL.Append(Environment.NewLine)
            End If

            For index As Integer = 0 To UpdateSQL.Count - 1
                If index = 0 Then logSQL.Append("Update:")
                logSQL.Append(Environment.NewLine & UpdateSQL.Item(index) & ";")
            Next
            If FinalSQL.Count > 0 Then
                logSQL.Append(Environment.NewLine)
            End If
            For index As Integer = 0 To FinalSQL.Count - 1
                If index = 0 Then logSQL.Append("Final:")
                logSQL.Append(Environment.NewLine & FinalSQL.Item(index) & ";")
            Next

            '-------------------------------------------------------------------------------
            BefUpdateData(BeforeUpd, cn, trans)


            Dim tbExecute As DataTable = DAO.ExecQry(sqlQuery)
            If SEQNO = 0 Then
                BllUtility.InsertProgramLog(SysProgramId, dsConditions.Tables("Condition"), SO.BLL.Utility.ExecType.TextFile,
                                        Nothing, True, Nothing, ParentSeqNo)
            End If

            If tbExecute.Rows.Count = 0 Then
                DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, Language.NoAnyData, DBNull.Value, logSQL.ToString, ParentSeqNo})
                If blnAutoClose Then
                    trans.Commit()
                End If
                Return "-3"
            End If

            If tbExecute.AsEnumerable.Count(Function(rwDataRowType As DataRow)
                                                Return Integer.Parse("0" & rwDataRowType.Item("DataRowType").ToString) = 0
                                            End Function) <= 0 Then
                DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, Language.NoAnyData, DBNull.Value, logSQL.ToString, ParentSeqNo})
                If blnAutoClose Then
                    trans.Commit()
                End If
                Return "-3"
            End If
            retFileName = WriteText(tbSingleDetail, tbExecute, UpdateSQL, isGetway, cn, trans)
            'If String.IsNullOrEmpty(filePath) Then
            '    retFileName = WriteText(tbSingleDetail, tbExecute, UpdateSQL, isGetway, cn, trans)
            'Else
            '    retFileName = WriteText(tbSingleDetail, filePath, tbExecute, isGetway, UpdateSQL, cn, trans)
            'End If


            For index As Integer = 0 To FinalSQL.Count - 1

                DAO.ExecNqry(FinalSQL.Item(index))
            Next
            RunTime.Stop()
            If tbExecute.Columns.Contains("ShouldAmt") Then
                retFileName = retFileName & ";" & String.Format(Language.RunOK & Environment.NewLine & _
                                                               Language.RunTotalRecord & Environment.NewLine & Language.RunSucessRecord & Environment.NewLine & _
                                                                Language.RunFailRecord & Environment.NewLine & Language.RunSucessAmt & Environment.NewLine & _
                                                               Language.RunSpendTime, FRecordCount, FSuccessCount,
                                                                FFailCount, FShouldAmt, Math.Round(RunTime.Elapsed.TotalSeconds, 1))
            Else
                retFileName = retFileName & ";" & String.Format(Language.RunOK & Environment.NewLine & _
                                                                Language.RunTotalRecord & Environment.NewLine & Language.RunSucessRecord & Environment.NewLine & _
                                                                Language.RunFailRecord & Environment.NewLine & _
                                                                Language.RunSpendTime2, FRecordCount, FSuccessCount,
                                                                FFailCount, Math.Round(RunTime.Elapsed.TotalSeconds, 1))
            End If

            If SEQNO = 0 AndAlso (Not isGetway) Then
                DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, retFileName.Split(";")(1),
                                retFileName.Split(";")(0), logSQL.ToString, ParentSeqNo})
            End If
            If blnAutoClose Then
                trans.Commit()
            End If
        Catch ex As Exception
            If SEQNO = 0 AndAlso (Not isGetway) Then
                DAO.ExecNqry(_DAL.UpdLogData, New Object() {1, ex.ToString, DBNull.Value, logSQL.ToString, ParentSeqNo})
            End If
            If blnAutoClose Then
                trans.Rollback()
            End If
            Throw ex
        Finally

            BeforeUpd.Clear()
            UpdateSQL.Clear()
            FinalSQL.Clear()
            If BllUtility IsNot Nothing Then
                BllUtility.Dispose()
                BllUtility = Nothing
            End If

            If tbMaster IsNot Nothing Then
                tbMaster.Dispose()
                tbMaster = Nothing
            End If
            If tbSingleDetail IsNot Nothing Then
                tbSingleDetail.Dispose()
                tbSingleDetail = Nothing
            End If
            If dynaCdt IsNot Nothing Then
                dynaCdt.Dispose()
                dynaCdt = Nothing
            End If
            If dtCdt IsNot Nothing Then
                dtCdt.Dispose()
                dtCdt = Nothing
            End If
            If dtReturn IsNot Nothing Then
                dtReturn.Dispose()
                dtReturn = Nothing
            End If
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
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

        Return retFileName
    End Function
    Public Function Execute(ByVal SysProgramId As String, ByVal AutoSerialNo As Integer, ByVal dsConditions As DataSet) As String
        Return Execute(SysProgramId, AutoSerialNo, False, 0, dsConditions)
        Dim programID As String = Nothing
        Dim condProgID As String = Nothing
        Dim sqlQuery As String = Nothing
        Dim dynaCdt As CableSoft.BLL.Dynamic.Condition.DynamicCondition = Nothing
        Dim tbMaster As DataTable = DAO.ExecQry(_DAL.QueryMaster, New Object() {SysProgramId})
        Dim tbSingleDetail As DataTable = DAO.ExecQry(_DAL.QuerySingleDetail, New Object() {AutoSerialNo})
        Dim dtReturn As DataTable = Nothing
        Dim dtCdt As DataTable = Nothing
        Dim cdtProgId As String = Nothing
        Dim retFileName As String = Nothing
        Dim updSQL As String = Nothing
        Dim BeforeUpd As New List(Of String)
        Dim UpdateSQL As New List(Of String)
        Dim FinalSQL As New List(Of String)
        Dim trans As DbTransaction = Nothing
        Dim cn As DbConnection = DAO.GetConn()
        Dim blnAutoClose As Boolean = False
        Dim RunTime As New Stopwatch()
        Dim BllUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim ParentSeqNo As Double = -1
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
        Try
            programID = tbMaster.Rows(0).Item("PROGRAMID").ToString
            condProgID = tbMaster.Rows(0).Item("CondProgId").ToString
            cdtProgId = DAO.ExecSclr(_DAL.QueryDynProgId, New Object() {condProgID})
            sqlQuery = tbSingleDetail.Rows(0).Item("SQLQUERY").ToString
            dtCdt = dsConditions.Tables("Condition")
            dynaCdt = New CableSoft.BLL.Dynamic.Condition.DynamicCondition(Me.LoginInfo, Me.DAO)
            Dim params() As Object = Nothing
            dtReturn = dynaCdt.GetBuildConditionSQL(cdtProgId, dtCdt, params)
            If Not DBNull.Value.Equals(tbSingleDetail.Rows(0).Item("BEFUPDSQL")) Then
                BeforeUpd = tbSingleDetail.Rows(0).Item("BEFUPDSQL").ToString.Split(";").ToList
            End If
            If Not DBNull.Value.Equals(tbSingleDetail.Rows(0).Item("UpdateSQL")) Then
                UpdateSQL = tbSingleDetail.Rows(0).Item("UpdateSQL").ToString.Split(";").ToList
            End If
            If Not DBNull.Value.Equals(tbSingleDetail.Rows(0).Item("FinalSQL")) Then
                FinalSQL = tbSingleDetail.Rows(0).Item("FinalSQL").ToString.Split(";").ToList
            End If
            Dim aFieldName As String = Nothing
            For Each dr As DataRow In dtReturn.Rows
                sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery, dr("FieldName"), dr("ConditionSQL"), params)
                If sqlQuery.ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                    For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                        Dim FieldName = PropertyInfo.Name
                        If PropertyInfo IsNot Nothing Then
                            sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery,
                                                                                      "LoginInfo." & FieldName,
                                                                                      GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                        End If
                    Next
                End If
                For int As Integer = 0 To BeforeUpd.Count - 1
                    BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         BeforeUpd.Item(int), dr("FieldName"), dr("ConditionSQL"), params)
                    If BeforeUpd.Item(int).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, BeforeUpd.Item(int),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If


                Next


                For index As Integer = 0 To UpdateSQL.Count - 1
                    UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         UpdateSQL.Item(index), dr("FieldName"), dr("ConditionSQL"), params)
                    If UpdateSQL.Item(index).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, UpdateSQL.Item(index),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                Next

                For inxFinal As Integer = 0 To FinalSQL.Count - 1
                    FinalSQL.Item(inxFinal) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         FinalSQL.Item(inxFinal), dr("FieldName"), dr("ConditionSQL"), params)
                    If FinalSQL.Item(inxFinal).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                FinalSQL.Item(inxFinal) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, FinalSQL.Item(inxFinal),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                    FinalSQL.Item(inxFinal) = ReplaceLoginInfoWhere(FinalSQL.Item(inxFinal), params)
                Next

            Next
            For i As Integer = 0 To dtCdt.Rows.Count - 1
                '2017/01/09 Jacky 增加取 _1 的資料
                If Right(dtCdt.Rows(i)("FieldName"), 2) = "_0" OrElse Right(dtCdt.Rows(i)("FieldName"), 2) = "_1" Then
                    aFieldName = dtCdt.Rows(i)("FieldName").ToString.Substring(0,
                                                                               dtCdt.Rows(i)("FieldName").ToString.Length - 2)
                Else
                    aFieldName = dtCdt.Rows(i)("FieldName").ToString
                End If
                sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery, aFieldName, dtCdt.Rows(i)("FieldValue"), params)
                If sqlQuery.ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                    For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                        Dim FieldName = PropertyInfo.Name
                        If PropertyInfo IsNot Nothing Then
                            sqlQuery = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, sqlQuery,
                                                                                      "LoginInfo." & FieldName,
                                                                                      GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                        End If
                    Next
                End If
                For int As Integer = 0 To BeforeUpd.Count - 1
                    BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         BeforeUpd.Item(int), aFieldName, dtCdt.Rows(i)("FieldValue"), params)
                    If BeforeUpd.Item(int).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                BeforeUpd.Item(int) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, BeforeUpd.Item(int),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                    BeforeUpd.Item(int) = ReplaceLoginInfoWhere(BeforeUpd.Item(int), params)
                Next
                For index As Integer = 0 To UpdateSQL.Count - 1
                    UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign,
                                                                                         UpdateSQL.Item(index), aFieldName, dtCdt.Rows(i)("FieldValue"), params)
                    If UpdateSQL.Item(index).ToUpper.IndexOf("[LOGININFO.") >= 0 Then
                        For Each PropertyInfo As Reflection.PropertyInfo In LoginInfo.GetType.GetProperties
                            Dim FieldName = PropertyInfo.Name
                            If PropertyInfo IsNot Nothing Then
                                UpdateSQL.Item(index) = CableSoft.BLL.Utility.Utility.GetFieldValueSQL(_DAL.Sign, UpdateSQL.Item(index),
                                                                                          "LoginInfo." & FieldName,
                                                                                          GetType(String), PropertyInfo.GetValue(LoginInfo, Nothing), params)
                            End If
                        Next
                    End If

                    UpdateSQL.Item(index) = ReplaceLoginInfoWhere(UpdateSQL.Item(index), params)
                Next

            Next

            BefUpdateData(BeforeUpd, cn, trans)
            Dim tbExecute As DataTable = DAO.ExecQry(sqlQuery)

            BllUtility.InsertProgramLog(SysProgramId, dsConditions.Tables("Condition"), SO.BLL.Utility.ExecType.TextFile,
                                        Nothing, True, Nothing, ParentSeqNo)

            If tbExecute.Rows.Count = 0 Then
                DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, Language.NoAnyData, DBNull.Value, ParentSeqNo})
                If blnAutoClose Then
                    trans.Commit()
                End If
                Return "-1"
            End If

            If tbExecute.AsEnumerable.Count(Function(rwDataRowType As DataRow)
                                                Return Integer.Parse("0" & rwDataRowType.Item("DataRowType").ToString) = 0
                                            End Function) <= 0 Then
                DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, Language.NoAnyData, DBNull.Value, ParentSeqNo})
                If blnAutoClose Then
                    trans.Commit()
                End If
                Return "-1"
            End If
            retFileName = WriteText(tbSingleDetail, tbExecute, UpdateSQL, False, cn, trans)
           

            For index As Integer = 0 To FinalSQL.Count - 1

                DAO.ExecNqry(FinalSQL.Item(index))
            Next
            RunTime.Stop()
            If tbExecute.Columns.Contains("ShouldAmt") Then
                retFileName = retFileName & ";" & String.Format(Language.RunOK & Environment.NewLine & _
                                                               Language.RunTotalRecord & Environment.NewLine & Language.RunSucessRecord & Environment.NewLine & _
                                                                Language.RunFailRecord & Environment.NewLine & Language.RunSucessAmt & Environment.NewLine & _
                                                               Language.RunSpendTime, FRecordCount, FSuccessCount,
                                                                FFailCount, FShouldAmt, Math.Round(RunTime.Elapsed.TotalSeconds, 1))
            Else
                retFileName = retFileName & ";" & String.Format(Language.RunOK & Environment.NewLine & _
                                                                Language.RunTotalRecord & Environment.NewLine & Language.RunSucessRecord & Environment.NewLine & _
                                                                Language.RunFailRecord & Environment.NewLine & _
                                                                Language.RunSpendTime2, FRecordCount, FSuccessCount,
                                                                FFailCount, Math.Round(RunTime.Elapsed.TotalSeconds, 1))
            End If


            DAO.ExecNqry(_DAL.UpdLogData, New Object() {0, retFileName.Split(";")(1), retFileName.Split(";")(0), ParentSeqNo})
            If blnAutoClose Then
                trans.Commit()
            End If
        Catch ex As Exception
            DAO.ExecNqry(_DAL.UpdLogData, New Object() {1, ex.ToString, DBNull.Value, ParentSeqNo})
            If blnAutoClose Then
                trans.Commit()
            End If
            Throw ex
        Finally
            BeforeUpd.Clear()
            UpdateSQL.Clear()
            FinalSQL.Clear()
            If BllUtility IsNot Nothing Then
                BllUtility.Dispose()
                BllUtility = Nothing
            End If

            If tbMaster IsNot Nothing Then
                tbMaster.Dispose()
                tbMaster = Nothing
            End If
            If tbSingleDetail IsNot Nothing Then
                tbSingleDetail.Dispose()
                tbSingleDetail = Nothing
            End If
            If dynaCdt IsNot Nothing Then
                dynaCdt.Dispose()
                dynaCdt = Nothing
            End If
            If dtCdt IsNot Nothing Then
                dtCdt.Dispose()
                dtCdt = Nothing
            End If
            If dtReturn IsNot Nothing Then
                dtReturn.Dispose()
                dtReturn = Nothing
            End If
            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
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

        Return retFileName
    End Function

    'Public Sub WriteText()

    '    'Using mem As New System.IO.StreamWriter("D:\Test.txt", False, System.Text.Encoding.ASCII)
    '    '    mem.WriteLine("ABCDEFG")
    '    '    mem.Flush()            
    '    'End Using
    '    Dim bdu As New System.Text.StringBuilder()
    '    bdu.AppendLine("ABCDEFG")
    '    bdu.AppendLine("HIJKLMN")


    '    Using mem As New System.IO.MemoryStream()
    '        Dim bty() As Byte = System.Text.Encoding.ASCII.GetBytes(bdu.ToString)
    '        mem.Write(bty, 0, bty.Length)
    '        mem.Flush()
    '        mem.Close()
    '        Using zip As New Ionic.Zip.ZipFile("D:\TEST.zip")
    '            zip.AddEntry("C:\Test.txt", bty)

    '            'zip.AddFile("D:\TEST.Txt")
    '            zip.Save()
    '        End Using
    '    End Using
    '    ' System.IO.File.Delete("D:\Test.txt")

    'End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                End If
                _DAL.Dispose()
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
