Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.SO.BLL.Facility
Public Class ThirdDiscount
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New ThirdDiscountDAL(Me.LoginInfo.Provider)
    Private Const FCurrectTableName As String = "ThirdDiscount"
    Private Const FPKField As String = "CustId,CompCode,ServiceType,FaciSeqNoB"
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
    

    Public Function QueryThirdDiscount(ByVal aCustId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.QueryThirdDiscount(),
                                          New Object() {aCustId})


    End Function
    Public Function QueryCanChooseFaci(ByVal aCustId As Integer,
                                            ByVal aServiceType As String) As DataTable
        Dim obj As New Facility.Facility(Me.LoginInfo, Me.DAO)
        Try
            Return obj.QueryCanChooseFaci(aCustId, aServiceType, False, False, False)
        Finally
            obj.Dispose()
        End Try
    End Function
    Public Function CanView() As CableSoft.BLL.Utility.RIAResult        
        Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
    End Function
    Public Function CanEdit() As CableSoft.BLL.Utility.RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO1100S1")
        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1100S1")
    End Function

    Public Function Save(ByVal EditMode As CableSoft.BLL.Utility.EditMode,
                         ByVal ThirdDiscount As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Try
            If Not HavePK(EditMode, ThirdDiscount.Tables(FCurrectTableName)) Then
                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO003B No PKField", .ResultBoolean = False}
            End If
            CSLog = New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo, Me.DAO)
           
            If DAO.Transaction IsNot Nothing Then

                trans = DAO.Transaction
            Else
                cn.ConnectionString = Me.LoginInfo.ConnectionString
                cn.Open()
                trans = cn.BeginTransaction
                DAO.Transaction = trans
                blnAutoClose = True
            End If

            DAO.AutoCloseConn = False
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId)
            Using cmd As System.Data.Common.DbCommand = DAO._factory.CreateCommand()
                cmd.Connection = cn
                cmd.Transaction = trans
                Dim aDtClone As DataTable = ThirdDiscount.Tables(0).Copy
                If aDtClone.Columns.IndexOf(aDtClone.Columns("FaciSNoA")) >= 0 Then
                    aDtClone.Columns.Remove(aDtClone.Columns("FaciSNoA"))
                End If
                If aDtClone.Columns.IndexOf(aDtClone.Columns("FaciSNoB")) >= 0 Then
                    aDtClone.Columns.Remove(aDtClone.Columns("FaciSNoB"))
                End If
                If aDtClone.Columns.IndexOf(aDtClone.Columns("RowId")) >= 0 Then
                    aDtClone.Columns.Remove(aDtClone.Columns("RowId"))
                End If
                'aDtClone.Columns.Remove(aDtClone.Columns("FaciSNoB"))
                Dim aWhere As String = String.Empty
                For i As Integer = 0 To ThirdDiscount.Tables(0).Rows.Count - 1
                    cmd.Parameters.Clear()
                    Select Case EditMode
                        Case CableSoft.BLL.Utility.EditMode.Append
                            If Not DAO.GetInsertCommand(ThirdDiscount.Tables(0), "SO003B", i, cmd) Then
                                Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO003B Insert Error", .ResultBoolean = False}
                            End If
                        Case CableSoft.BLL.Utility.EditMode.Edit, CableSoft.BLL.Utility.EditMode.Delete

                            'If ThirdDiscount.Tables(0).Columns.IndexOf(ThirdDiscount.Tables(0).Columns("RowId")) < 0 Then

                            '    Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = " 需傳入RowID ", .ResultBoolean = False}
                            'Else
                            '    aWhere = String.Format("SO003B.RowID='{0}'", ThirdDiscount.Tables(0).Rows(i).Item("RowID"))
                            'End If
                            If ThirdDiscount.Tables(FCurrectTableName).Columns.Contains("RowId") Then
                                aWhere = String.Format("SO003B.RowID='{0}'", ThirdDiscount.Tables(FCurrectTableName).Rows(i).Item("RowID"))
                            Else
                                aWhere = String.Format("SO003B.CustId = {0} AND SO003B.CompCode = {1} " &
                                                      "SO003B.ServiceType='{2}' AND SO003B.FaciSeqNoB='{3}' ",
                                                      ThirdDiscount.Tables(FCurrectTableName).Rows(i).Item("CUSTID"),
                                                        Me.LoginInfo.CompCode,
                                                        ThirdDiscount.Tables(FCurrectTableName).Rows(i).Item("SERVICETYPE"),
                                                        ThirdDiscount.Tables(FCurrectTableName).Rows(i).Item("FACISEQNOB"))
                            End If

                            If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                                If aDtClone.Columns.Contains("UpdTime") Then
                                    aDtClone.Rows(i).Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(Now.Date)
                                End If
                                If aDtClone.Columns.Contains("UpdEn") Then
                                    aDtClone.Rows(i).Item("UpdEn") = Me.LoginInfo.EntryName
                                End If

                                If Not DAO.GetInsertOrUpdateCommand(CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, aDtClone, "SO003B", i, cmd, aWhere) Then
                                    Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO003B Update Error", .ResultBoolean = False}
                                End If
                            Else
                                cmd.CommandText = "DELETE SO003B WHERE " & aWhere
                            End If

                            'If Not DAO.GetUpdateCommand(ThirdDiscount.Tables(0), "SO003B", i, "", cmd) Then
                            '    Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = "SO003B Update Error", .ResultBoolean = False}
                            'End If

                    End Select
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                    Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update,
                                                                      "SO003B", ThirdDiscount.Tables(FCurrectTableName),
                                                                      Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                    If Not aResult.ResultBoolean Then
                        Select Case aResult.ErrorCode
                            Case -5
                            Case -6
                                If blnAutoClose Then
                                    trans.Rollback()
                                    Return aResult
                                End If

                        End Select

                    End If
                Next
            End Using

            If blnAutoClose Then
                trans.Commit()
            End If


        Catch ex As Exception
            If (trans IsNot Nothing) AndAlso (blnAutoClose) Then
                trans.Rollback()
            End If
            Return New RIAResult() With {.ErrorCode = -1, .ErrorMessage = ex.Message, .ResultBoolean = False}
        Finally
            If blnAutoClose Then
                DAO.AutoCloseConn = True
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()

                End If
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                End If
            End If
            
        End Try
        Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = "OK", .ResultBoolean = True}
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

    Protected Overrides Sub Finalize()

        MyBase.Finalize()
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
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
