Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports System.Web
Public Class IDData
    Inherits BLLBasic
    Private _DAL As New IDDataDAL(Me.LoginInfo.Provider)
    Private Const FCurrectTableName As String = "IDData"
    Private Const FPKField As String = "Custid,FaciSNO,IdKind,CreateDate"
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub
    ''' <summary>
    ''' 可新增
    ''' </summary>
    ''' <returns>RIAResult</returns>
    ''' <remarks>ErrorHandle 0: 成功 -1: 失敗</remarks>
    Public Function CanAppend() As RIAResult
        Using objUtility As New Utility.Utility(Me.LoginInfo)
            Return objUtility.ChkPriv(Me.LoginInfo.EntryId, "SO112071")
        End Using
    End Function
    ''' <summary>
    ''' 可修改
    ''' </summary>
    ''' <returns>RIAResult</returns>
    ''' <remarks>ErrorHandle 0: 成功 -1: 失敗</remarks>
    Public Function CanEdit() As RIAResult
        Using objUtility As New Utility.Utility(Me.LoginInfo)
            Return objUtility.ChkPriv(Me.LoginInfo.EntryId, "SO112072")
        End Using
    End Function
    ''' <summary>
    ''' 可刪除
    ''' </summary>
    ''' <returns>RIAResult</returns>
    ''' <remarks>ErrorHandle 0: 成功 -1: 失敗</remarks>
    Public Function CanDelete() As RIAResult
        Using objUtility As New Utility.Utility(Me.LoginInfo)
            Return objUtility.ChkPriv(Me.LoginInfo.EntryId, "SO112073")
        End Using
    End Function
    ''' <summary>
    ''' 取得可選證件種類
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks>0:申請人第一證件</remarks>
    Public Function QueryIDKind() As DataTable
        Return DAO.ExecQry(_DAL.QueryIDKind)

    End Function
    ''' <summary>
    ''' 取得證件資料
    ''' </summary>
    ''' <param name="FaciSNO">設備序號</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function QueryIDData(ByVal FaciSNO As String) As DataTable
        Return DAO.ExecQry(_DAL.QueryIDData, New Object() {FaciSNO})

    End Function
    ''' <summary>
    ''' 取得證件圖檔
    ''' </summary>
    ''' <param name="PicturePath">PicturePath</param>
    ''' <param name="PictureName">PictureName</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueryIDPicture(ByVal PicturePath As String,
                                   ByVal PictureName As String) As RIAResult
        Try
            Dim aResult As New RIAResult(0, "")
            Dim aFile As String = String.Empty
            Dim aTmpPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Templates)
            If Right(PicturePath, 1) <> "\" Then
                PicturePath = PicturePath & "\"
            End If
            If Right(aTmpPath, 1) <> "\" Then
                aTmpPath = aTmpPath & ""
            End If
            aFile = PicturePath & PictureName
            System.IO.File.Copy(aFile, aTmpPath & PictureName, True)
            Dim objUrl As New Uri(aTmpPath & PictureName)
            aResult.ResultXML = objUrl.ToString

            Return (aResult)
        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        End Try
    End Function
    'Custid,FaciSNO,IdKind,CreateDate
    Public Function Save(ByVal EditMode As CableSoft.BLL.Utility.EditMode,
                         ByVal IDData As System.Data.DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Try
            If Not HavePK(EditMode, IDData.Tables(FCurrectTableName)) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004E NO PKField", .ResultBoolean = False}
            End If
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                Dim aWhere As String = String.Empty
                Dim aTB As DataTable = GetCorrectTable(IDData, EditMode)

                For i As Integer = 0 To aTB.Rows.Count - 1                    
                    Select Case EditMode
                        Case CableSoft.BLL.Utility.EditMode.Append
                            If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.InsertRow, aTB, "SO004E", i, cmd, aWhere) Then
                                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004E Insert Error", .ResultBoolean = False}
                            End If
                        Case CableSoft.BLL.Utility.EditMode.Edit, Global.CableSoft.BLL.Utility.EditMode.Delete
                            If IDData.Tables(FCurrectTableName).Columns.Contains("RowId") Then
                                aWhere = String.Format("ROWID='{0}'",
                                                   IDData.Tables(FCurrectTableName).Rows(i).Item("RowId"))
                            Else
                                aWhere = String.Format("CustId={0} AND FaciSNO = '{1}' " &
                                                     " AND IdKind= {2} AND CreateDate = TO_DATE('{3}','YYYYMMDDHH24MISS')",
                                                     aTB.Rows(i).Item("CustId"), aTB.Rows(i).Item("FaciSNO"),
                                                     aTB.Rows(i).Item("IDKind"), Format(Date.Parse(aTB.Rows(i).Item("CreateDate")), "yyyyMMddHHmmss"))

                            End If
                            If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                                If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, aTB, "SO004E", i, cmd, aWhere) Then
                                    Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004E Update Error", .ResultBoolean = False}
                                End If
                            Else
                                cmd.CommandText = "delete SO004E where " & aWhere
                            End If                            
                    End Select
                Next
                cmd.ExecuteNonQuery()
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
    'Public Function GetWhereString(ByVal aEditMode As EditMode, ByVal aTB As DataTable, ByVal aRowIndex As Int32) As String
    '    Dim aReturn As String = String.Empty
    '    Try
    '        If aEditMode = EditMode.Append Then
    '            Return String.Empty
    '        End If
    '        If aTB.Columns.Contains("ROWID") Then
    '            Return "ROWID='" & aTB.Rows(aRowIndex).Item("ROWID") & "'"
    '        Else
    '            For Each s As String In FPKField
    '                If String.IsNullOrEmpty(aReturn) Then
    '                    If TypeOf aTB.Rows(aRowIndex).Item(s) Is Int32 Then

    '                    End If
    '                  aReturn=
    '                End If
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function
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
            If Not aRetTb.Columns.Contains("CustId") Then Throw New Exception("無客戶名稱欄位！")
            If Not aRetTb.Columns.Contains("CompCode") Then Throw New Exception("無公司別欄位！")
            If Not aRetTb.Columns.Contains("IDKind") Then Throw New Exception("無證件種類欄位！")
            If Not aRetTb.Columns.Contains("PicturePath") Then Throw New Exception("無證件圖檔路徑！")
            If Not aRetTb.Columns.Contains("IdPictureName") Then Throw New Exception("無證件圖檔名稱！")

            Select Case EditMode
                Case EditMode.Append
                    If aRetTb.Columns.Contains("RowId") Then
                        aRetTb.Columns.Remove(aRetTb.Columns("RowId"))
                    End If
                Case EditMode.Edit

                    If aRetTb.Columns.Contains("RowId") Then
                        aRetTb.Columns.Remove(aRetTb.Columns("RowId"))
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
