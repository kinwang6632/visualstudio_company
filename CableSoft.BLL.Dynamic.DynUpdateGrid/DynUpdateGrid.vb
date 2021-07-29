Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports System.Reflection
Public Class DynUpdateGrid
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New DynUpdateGridDALMultiDB(Me.LoginInfo.Provider)
    'Private tbMaster As DataTable = Nothing
    'Private tbDetail As DataTable = Nothing
    'Private Const DefaultField As String = "FinalValue"
    'Private Const tbMasterName As String = "Master"
    'Private Const tbDetailName As String = "Detail"
    'Private Const LoginInfoString As String = "LoginInfo"
    'Private Const SeqNoString As String = "SEQNO"
    'Private fFieldsAndValues As Dictionary(Of String, Object)
    'Private fWhereFieldsAndValues As Dictionary(Of String, Object)
    'Private tbSechema As DataTable = Nothing
    Private lang As New CableSoft.BLL.Language.SO61.DynUpdateGridLanguage()
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
        Try
            tbMaster.TableName = "Master"
            tbDetail.TableName = "Detail"
            tbMaster.Columns.Add("CanAppend", GetType(Boolean))
            tbMaster.Columns.Add("CanEdit", GetType(Boolean))
            tbMaster.Columns.Add("CanDelete", GetType(Boolean))
            'tbMaster.Columns.Add("CanCopyOtherDB", GetType(Boolean))
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
    Public Function GetCanChooseComp() As DataTable
        Return DAO.ExecQry(_DAL.QueryCD052).Copy
    End Function
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
    'Public Function CanAppend(ByVal PrivMid As String) As RIAResult
    '    Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, PrivMid)
    'End Function

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
                If lang IsNot Nothing Then
                    lang.Dispose()
                    lang = Nothing
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
