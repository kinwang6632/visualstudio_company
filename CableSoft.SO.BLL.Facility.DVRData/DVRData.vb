Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports System.Web

Public Class DVRData
    Inherits BLLBasic
    Private _DAL As New DVRDataDAL(Me.LoginInfo.Provider)
    Private Const FCurrectTableName As String = "DVRData"
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub
    ''' <summary>
    ''' 檢核號碼是否正確
    ''' </summary>
    ''' <param name="SeqNo">設備流水號</param>
    ''' <param name="PhoneNumber">行動電話</param>
    ''' <param name="WebConnections">Web Config 連線字串</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks>True Or False</remarks>
    Public Function ChkPhoneNumberOk(ByVal SeqNo As String, ByVal PhoneNumber As String,
                                     ByVal WebConnections As Dictionary(Of String, String)) As RIAResult
        Try
            Using objUtility As New Utility.Utility(Me.LoginInfo)
                Dim aResult As RIAResult = objUtility.ChkPriv(Me.LoginInfo.EntryId, "SO1120A")
                If aResult.ResultBoolean Then
                    Return aResult
                Else
                    aResult.ErrorCode = 0
                    aResult.ErrorMessage = ""
                    aResult.ResultBoolean = True
                    If WebConnections Is Nothing OrElse WebConnections.Count <= 0 Then
                        aResult.ErrorCode = -1
                        aResult.ErrorMessage = "Web Config無設定連線資訊"
                        aResult.ResultBoolean = False
                        Return aResult
                    End If
                    Using aTbComp As DataTable = DAO.ExecQry(_DAL.GetComps)
                        For Each aRow As DataRow In aTbComp.Rows
                            If WebConnections.ContainsKey(aRow.Item("CodeNo")) Then
                                Me.LoginInfo.CompCode = Integer.Parse(aRow("CodeNo"))
                                Me.LoginInfo.ConnectionString = WebConnections.Item(aRow("CodeNo"))
                                If Integer.Parse(DAO.ExecSclr(_DAL.ChkPhoneNumberOk,
                                              New Object() {Right(PhoneNumber, 9), SeqNo})) > 0 Then

                                    aResult.ErrorCode = -1
                                    aResult.ErrorMessage = String.Format("該行動電話存在於公司別:{0}", aRow("Description"))
                                    aResult.ResultBoolean = False
                                    Exit For
                                End If
                            End If

                        Next
                    End Using
                    Return aResult
                End If
            End Using
        Catch ex As Exception
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        End Try


    End Function
    ''' <summary>
    ''' 取得行動電話資料
    ''' </summary>
    ''' <param name="FaciSeqNo">FaciSeqNo</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function QueryPhoneNumber(ByVal FaciSeqNo As String) As DataTable
        Return DAO.ExecQry(_DAL.QueryPhoneNumber, New Object() {FaciSeqNo})
    End Function
    ''' <summary>
    ''' 存檔(Save)
    ''' </summary>
    ''' <param name="EditMode">狀態</param>
    ''' <param name="DVRData">SO004H DataSet</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function Save(ByVal EditMode As CableSoft.BLL.Utility.EditMode,
                         ByVal DVRData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Try
            'Dim obj As New Uri(Environment.GetFolderPath(Environment.SpecialFolder.Templates))

            'Dim obj As New Uri(Environment.GetFolderPath(Environment.SpecialFolder.Templates))

            'Dim obj2 As New System.Web.HttpRequest("excel4.xls", obj.PathAndQuery, Nothing)


            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                Dim aWhere As String = String.Empty

                Dim aTB As DataTable = GetCorrectTable(DVRData, EditMode)


                For i As Integer = 0 To aTB.Rows.Count - 1
                    If (DBNull.Value.Equals(aTB.Rows(i).Item("CellPhone"))) _
                            OrElse (String.IsNullOrEmpty(aTB.Rows(i).Item("CellPhone"))) Then
                        trans.Rollback()
                        Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004H CellPhone IS Null", .ResultBoolean = False}
                    End If
                    Select Case EditMode
                        Case CableSoft.BLL.Utility.EditMode.Append
                            If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.InsertRow, aTB, "SO004H", i, cmd, aWhere) Then
                                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004H Update Error", .ResultBoolean = False}
                            End If
                        Case CableSoft.BLL.Utility.EditMode.Edit
                            aWhere = String.Format("ROWID='{0}'",
                                                   DVRData.Tables(FCurrectTableName).Rows(i).Item("RowId"))

                            If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, aTB, "SO004H", i, cmd, aWhere) Then
                                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO004H Update Error", .ResultBoolean = False}
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

    Private Function GetCorrectTable(ByVal aDs As DataSet,
                                     ByVal EditMode As CableSoft.BLL.Utility.EditMode) As DataTable
        Try
            Dim aRetTb As DataTable = aDs.Tables(FCurrectTableName).Copy
            Dim twC = New System.Globalization.TaiwanCalendar()
            Dim aUpdTime As String = twC.GetYear(Date.Now).ToString + Date.Now.ToString("/MM/dd HH:mm:ss")
            Dim aNow As Date = Date.Now
            If aRetTb.Rows.Count <= 0 Then Throw New Exception("無任何資料可異動!")
            If Not aRetTb.Columns.Contains("CellPhone") Then Throw New Exception("無行動電話欄位！")

            Select Case EditMode
                Case EditMode.Append
                    If aRetTb.Columns.Contains("RowId") Then
                        aRetTb.Columns.Remove(aRetTb.Columns("RowId"))
                    End If
                Case EditMode.Edit
                    If Not aRetTb.Columns.Contains("RowId") Then Throw New Exception("無傳入RowId")
                    If aRetTb.Columns.Contains("RowId") Then
                        aRetTb.Columns.Remove(aRetTb.Columns("RowId"))
                    End If
                Case EditMode.Delete
                    If Not aRetTb.Columns.Contains("RowId") Then Throw New Exception("無傳入RowId")
            End Select

            For i As Int32 = 0 To aRetTb.Rows.Count - 1
                aRetTb.Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(aNow)
                aRetTb.Rows(i).Item("UpdEn") = Me.LoginInfo.EntryName
            Next

            Return aRetTb
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
