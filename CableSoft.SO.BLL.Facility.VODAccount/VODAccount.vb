Imports CableSoft.BLL.Utility
Imports System.Data.Common
Public Class VODAccount
    Inherits BLLBasic
    Private _DAL As New VODAccountDAL(Me.LoginInfo.Provider)
    Private Const FCurrectTableName As String = "VODData"
    Private Const FCurrectTableName2 As String = "ReqData"
    Private Const FPKField As String = "VODAccountID"
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
        

    End Sub
    ''' <summary>
    ''' 查詢VOD資訊
    ''' </summary>
    ''' <param name="VODAccountID">VODAccountID</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function QueryVODAccount(ByVal VODAccountID As String) As DataTable
        Return DAO.ExecQry(_DAL.QueryVODAccount, New Object() {VODAccountID})
    End Function
    ''' <summary>
    ''' 取得可選點數行銷辦法
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetSalePointcode() As DataTable
        Return DAO.ExecQry(_DAL.GetSalePointcode)
    End Function
    ''' <summary>
    ''' 查詢可選修改人員
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetReqEmpNo() As DataTable
        Return DAO.ExecQry(_DAL.GetReqEmpNo)
    End Function

    Public Function Save(ByVal EditMode As CableSoft.BLL.Utility.EditMode,
                     ByVal VODAccount As System.Data.DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        'VODAccount.Tables(0).Select(Nothing, Nothing, DataViewRowState.ModifiedOriginal)
        'Dim obj As New DataView(VODAccount.Tables(0))
        'Dim obj2 As DataRowView = VODAccount.Tables(0).Rows(0)
        'obj2.DataView.RowStateFilter = DataViewRowState.ModifiedOriginal

        Try
            If Not HavePK(EditMode, VODAccount.Tables(FCurrectTableName)) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004G NO PKField", .ResultBoolean = False}
            End If
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                Dim aWhere As String = String.Empty
                Dim aTB As DataTable = GetCorrectTable(VODAccount, EditMode)

                cmd.CommandText = "SELECT * FROM SO182LOG WHERE 1=0"
                Dim aSO182LogSchema As DataTable = cmd.ExecuteReader.GetSchemaTable
                Dim aTBSO182Log As New DataTable(FCurrectTableName2)
                For aIndex As Int32 = 0 To aSO182LogSchema.Rows.Count - 1
                    aTBSO182Log.Columns.Add(aSO182LogSchema.Rows(aIndex).Item("ColumnName").ToString,
                          aSO182LogSchema.Rows(aIndex).Item("DataType"))
                Next
                

                For i As Integer = 0 To aTB.Rows.Count - 1
                    cmd.Parameters.Clear()
                    Select Case EditMode
                        Case CableSoft.BLL.Utility.EditMode.Append
                            If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.InsertRow, aTB, "SO004G", i, cmd, aWhere) Then
                                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004G Insert Error", .ResultBoolean = False}

                            End If
                        Case CableSoft.BLL.Utility.EditMode.Edit, Global.CableSoft.BLL.Utility.EditMode.Delete
                            If VODAccount.Tables(FCurrectTableName).Columns.Contains("RowId") Then
                                aWhere = String.Format("ROWID='{0}'",
                                                   VODAccount.Tables(FCurrectTableName).Rows(i).Item("RowId"))
                            Else
                                aWhere = String.Format("VODAccountID={0}",
                                                       VODAccount.Tables(FCurrectTableName).Rows(i).Item("VODAccountID"))

                            End If
                            If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                                If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, aTB, "SO004G", i, cmd, aWhere) Then
                                    Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004G Update Error", .ResultBoolean = False}
                                End If
                            Else
                                cmd.CommandText = "DELETE SO004G WHERE " & aWhere
                            End If
                    End Select
                    cmd.ExecuteNonQuery()
                    '新增資料至SO182Log
                    cmd.Parameters.Clear()

                    aTBSO182Log.Rows.Clear()
                    aTBSO182Log.Rows.Add(
                            InsSO182Log(VODAccount.Tables(FCurrectTableName2).Rows(0),
                                    VODAccount.Tables(FCurrectTableName).Rows(i), aTBSO182Log))
                    aTBSO182Log.Rows(0).Item("UpdEn") = aTB.Rows(i).Item("UpdEn")
                    aTBSO182Log.Rows(0).Item("UpdTime") = aTB.Rows(i).Item("UpdTime")


                    If Not DAO.GetInsertCommand(aTBSO182Log, "SO182LOG", 0, cmd) Then
                        Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO182Log Insert Error", .ResultBoolean = False}
                    End If
                    'cmd.Parameters.Count
                    'cmd.Parameters(0).Value
                    'cmd.Parameters(0).ParameterName
                    cmd.ExecuteNonQuery()
                Next
                aTB.Dispose()
                aTBSO182Log.Dispose()
                aTB.Dispose()
            End Using

            trans.Commit()
            Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = "OK", .ResultBoolean = True}
        Catch ex As Exception
            If trans IsNot Nothing Then
                trans.Rollback()
            End If
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        Finally
            If trans IsNot Nothing Then
                trans.Dispose()
            End If
            If cn IsNot Nothing Then
                cn.Close()
                cn.Dispose()
            End If

        End Try
    End Function
    Private Function InsSO182Log(ByVal aRowReqDep As DataRow, ByVal aRowSO004G As DataRow,
                                 ByVal aTBSO182Log As DataTable) As DataRow
        Try
            Dim aRetRow As DataRow = aTBSO182Log.NewRow
            aRetRow.Item("ReqDep") = aRowReqDep.Item("ReqDep")
            aRetRow.Item("ReqEn") = aRowReqDep.Item("ReqEn")
            aRetRow.Item("ReqNotes") = aRowReqDep.Item("ReqNotes")
            If aRowSO004G.HasVersion(DataRowVersion.Original) Then
                aRetRow.Item("OAddCredit") = aRowSO004G.Item("PrePay", DataRowVersion.Original)
            Else

                aRetRow.Item("OAddCredit") = DBNull.Value
            End If
          

            aRetRow.Item("NAddCredit") = aRowSO004G.Item("PrePay")
            aRetRow.Item("VODAccountId") = aRowSO004G.Item("VODAccountID")
          
            Return aRetRow
        Catch ex As Exception
            Throw

        End Try

    End Function
    Private Function HavePK(ByVal aEditMode As EditMode, ByVal aTB As DataTable) As Boolean
        Try
            If aEditMode = EditMode.Append Then
                Return True
            End If
            If aTB.Columns.Contains("ROWID") Then
                Return True
            End If
            For Each s As String In FPKField.Split(",")
                If Not aTB.Columns.Contains(s) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function GetCorrectTable(ByVal aDs As DataSet,
                             ByVal EditMode As CableSoft.BLL.Utility.EditMode) As DataTable
        Try
            Dim aRetTb As DataTable = aDs.Tables(FCurrectTableName).Copy
            Dim aNow As Date = Date.Now
            If aRetTb.Rows.Count <= 0 Then Throw New Exception("無任何資料可異動!")
            If Not aRetTb.Columns.Contains("VODCredit") Then Throw New Exception("無預借點數欄位！")
           

            Select Case EditMode
                Case EditMode.Append
                    If aRetTb.Columns.Contains("RowId") Then
                        aRetTb.Columns.Remove(aRetTb.Columns("RowId"))
                    End If
                Case EditMode.Edit

                    If aRetTb.Columns.Contains("RowId") Then
                        aRetTb.Columns.Remove(aRetTb.Columns("RowId"))
                    End If
                    If aRetTb.Columns.Contains("VODAccountID") Then
                        aRetTb.Columns.Remove(aRetTb.Columns("VODAccountID"))
                    End If
                Case EditMode.Delete
                    'If Not aRetTb.Columns.Contains("RowId") Then Throw New Exception("無傳入RowId")
            End Select

            For i As Int32 = 0 To aRetTb.Rows.Count - 1
                If aRetTb.Columns.Contains("UpdTime") Then
                    aRetTb.Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(aNow)
                End If
                If aRetTb.Columns.Contains("UpdEn") Then
                    aRetTb.Rows(i).Item("UpdEn") = Me.LoginInfo.EntryName
                End If
            Next
            Return aRetTb
        Catch ex As Exception
            Throw ex
        End Try
    End Function

End Class
