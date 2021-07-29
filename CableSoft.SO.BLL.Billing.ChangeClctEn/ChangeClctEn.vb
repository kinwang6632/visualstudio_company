Imports CableSoft.BLL.Utility
Imports System.Data.Common

Public Class ChangeClctEn
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New ChangeClctEnDALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.ChangeClctEnLanguage
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
    Public Function GetAllData() As DataSet
        Using dsAll As New DataSet
            Using dtComp As DataTable = GetCompCode()
                dtComp.TableName = "COMPCODE"
                dsAll.Tables.Add(dtComp.Copy)
            End Using
            Using dtClctEn As DataTable = GetClctEn()
                dtClctEn.TableName = "CLCTEN"
                dsAll.Tables.Add(dtClctEn.Copy)
            End Using
            Using dtStrtCode As DataTable = GetStrtCode()
                dtStrtCode.TableName = "STRCODE"
                dsAll.Tables.Add(dtStrtCode.Copy)
            End Using
            Using dtServiceType As DataTable = GetServiceType()
                dtServiceType.TableName = "SERVICETYPE"
                dsAll.Tables.Add(dtServiceType.Copy)
            End Using
            Using dtMduId As DataTable = GetMduId()
                dtMduId.TableName = "MDUID"
                dsAll.Tables.Add(dtMduId.Copy)
            End Using
           
            Return dsAll.Copy
        End Using
       
    End Function
    Public Function ChkAuthority(ByVal refNo As Integer) As RIAResult
        Dim result As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}

        Dim strKey As String = "SO3100"
        Dim strItem As String = "SO3130"
        If refNo = 1 Then
            strItem = "SO3253"
        End If
        Using obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
            result = obj.ChkPriv(Me.LoginInfo.EntryId, strItem)
        End Using
        Return result
    End Function
    '取得可選公司別
    Public Function GetCompCode() As DataTable
        Try
            If Me.LoginInfo.GroupId = "0" AndAlso 1 = 0 Then
                Return DAO.ExecQry(_DAL.GetCompCode("0"))
            Else
                Return DAO.ExecQry(_DAL.GetCompCode("1"), New Object() {Me.LoginInfo.EntryId})
            End If
        Catch ex As Exception
            Throw
        End Try

    End Function
    '取得可選收費人員
    Public Function GetClctEn() As DataTable
        Try
            Using dtReturn As DataTable = DAO.ExecQry(_DAL.GetClctEn)
                'Dim rw As DataRow = dtReturn.NewRow
                'dtReturn.Rows.InsertAt(rw, 0)
                'dtReturn.AcceptChanges()
                Return dtReturn
            End Using
            'Dim dtCopy As DataTable = DAO.ExecQry(_DAL.GetClctEn).Copy
            'Return dtCopy
            'Return DAO.ExecQry(_DAL.GetClctEn)
        Catch ex As Exception
            Throw
        End Try
    End Function
    '取得可選街道編號
    Public Function GetStrtCode() As DataTable
        Return DAO.ExecQry(_DAL.GetStrtCode)
    End Function
    Public Function GetServiceType() As DataTable
        Return DAO.ExecQry(_DAL.GetServiceType)
    End Function
    Public Function GetMduId() As DataTable
        Return DAO.ExecQry(_DAL.GetMduId)
    End Function
    '取得收費員街道群組資料
    Public Function GetClctStrtGroupData(ByVal ClctEnStr As String, ByVal StrtCodeStr As String) As DataTable
        Return DAO.ExecQry(_DAL.GetClctStrtGroupData(ClctEnStr, StrtCodeStr))
    End Function
    Public Function GetGroupData(ByVal RefNo As Integer, ByVal tbWhere As DataTable) As DataTable
        Return DAO.ExecQry(_DAL.GetGroupData(RefNo, tbWhere))
    End Function
    Public Function Execute(ByVal ModifyData As DataTable, ByVal tbPara As DataTable,
                            ByVal RefNo As Integer,
                            ByVal GroupByStr As Integer, ByVal ModifyType As Integer) As Integer
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aNow As Date = Date.Now
        Try
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
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId, Language.ClientInfoString)
            Select Case ModifyType
                Case 0

            End Select

            For Each rw As DataRow In ModifyData.Rows

                Select Case ModifyType
                    Case 0
                        DAO.ExecNqry(_DAL.Execute(RefNo, GroupByStr, ModifyType, tbPara, rw), New Object() {rw.Item("NewClctEn"),
                                                           rw.Item("NewClctName"),
                                                           rw.Item("NewClctEn"),
                                                           rw.Item("NewClctName"),
                                                            CableSoft.BLL.Utility.DateTimeUtility.GetDTString(aNow),
                                                           Me.LoginInfo.EntryName,
                                                            aNow,
                                                           rw.Item("GroupCode")})
                    Case 1, 2
                        DAO.ExecNqry(_DAL.Execute(RefNo, GroupByStr, ModifyType, tbPara, rw),
                                                New Object() {rw.Item("NewClctEn"),
                                                           rw.Item("NewClctName"),
                                                           CableSoft.BLL.Utility.DateTimeUtility.GetDTString(aNow),
                                                           Me.LoginInfo.EntryName,
                                                            aNow,
                                                           rw.Item("GroupCode")})
                End Select

            Next
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(Me.DAO)
                DAO.AutoCloseConn = True
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If

            End If
            ModifyData.Dispose()
        End Try
        Return 0
    End Function

    Public Function Execute(ByVal ModifyData As DataTable) As Int32
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aNow As Date = Date.Now
        Try
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
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId)
            For Each rw As DataRow In ModifyData.Rows
                DAO.ExecNqry(_DAL.Execute(rw), New Object() {rw.Item("NewClctEn"),
                                                             rw.Item("NewClctName"),
                                                             rw.Item("NewClctEn"),
                                                             rw.Item("NewClctName"),
                                                              CableSoft.BLL.Utility.DateTimeUtility.GetDTString(aNow),
                                                             Me.LoginInfo.EntryName,
                                                             rw.Item("strtCode")})
            Next
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw
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

            End If
            ModifyData.Dispose()
        End Try
        Return 0
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
