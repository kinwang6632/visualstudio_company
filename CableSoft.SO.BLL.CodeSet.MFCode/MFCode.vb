Imports System.Data.Common
Imports CableSoft.BLL.Utility

Public Class MFCode
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New MFCodeDALMultiDB(Me.LoginInfo.Provider)
    Private FNowDate As DateTime = Now
    Private Language As New CableSoft.BLL.Language.SO61.MFCodeLanguage
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

    Public Function CopyToOtherDB(ByVal IsCoverData As Boolean, ByVal dsSource As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim ErrMsg As String = Nothing
        Dim result As New RIAResult
        Dim strRetMsg As New System.Text.StringBuilder()
        Dim tbCD011A As DataTable = Nothing
        Dim tbCD011 As DataTable = Nothing
        Dim MFCode As String = Nothing
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
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryName)
        MFCode = dsSource.Tables("CD011").Rows(0).Item("CodeNo")

        Try
            tbCD011 = DAO.ExecQry(_DAL.QueryCD011Code, New Object() {MFCode})
            tbCD011A = DAO.ExecQry(_DAL.QueryCD011ACode, New Object() {MFCode})
           

            For Each rwCompCode As DataRow In dsSource.Tables("SelectedComp").Rows
                Dim otherLoginInfo As LoginInfo = CableSoft.BLL.Utility.Utility.GetRealLoginInfo(LoginInfo, rwCompCode.Item("CompId"))
                Dim otherDao As New CableSoft.Utility.DataAccess.DAO(otherLoginInfo.Provider, otherLoginInfo.ConnectionString)
                Dim Othertrans As DbTransaction = Nothing

                If otherDao.GetConn.State <> ConnectionState.Open Then
                    otherDao.GetConn.Open()
                End If
                If otherDao.Transaction IsNot Nothing Then
                    Othertrans = otherDao.Transaction
                Else
                    Othertrans = otherDao.GetConn.BeginTransaction
                    otherDao.Transaction = Othertrans
                End If
                otherDao.AutoCloseConn = False
                CableSoft.BLL.Utility.Utility.SetClientInfo(otherDao, otherLoginInfo.EntryName)

                Try

                    If Not IsCoverData Then
                        '處理CD011
                        If Integer.Parse(otherDao.ExecSclr(_DAL.QueryMasterExists, New Object() {MFCode})) = 0 Then
                            otherDao.ExecNqry(_DAL.DeleteCD011, MFCode)
                            otherDao.ExecNqry(_DAL.InsertCD011, New Object() {tbCD011.Rows(0).Item("CodeNo"),
                                                                tbCD011.Rows(0).Item("Description"),
                                                                 tbCD011.Rows(0).Item("RefNo"),
                                                                 CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                                                                 otherLoginInfo.EntryName,
                                                                tbCD011.Rows(0).Item("ServiceType"),
                                                                 tbCD011.Rows(0).Item("StopFlag")})
                            strRetMsg.AppendLine(String.Format(Language.CopyOK, otherLoginInfo.CompCode))
                        Else
                            strRetMsg.AppendLine(String.Format(Language.CodeExists, otherLoginInfo.CompCode))
                        End If
                        '處理CD011A
                        For Each rwCD011A As DataRow In tbCD011A.Rows
                            If Integer.Parse(otherDao.ExecSclr(_DAL.QueryDetailExists, New Object() {MFCode, rwCD011A.Item("CodeNo")})) = 0 Then
                                otherDao.ExecNqry(_DAL.InsertCD011A, New Object() {MFCode,
                                                               rwCD011A.Item("CodeNo"),
                                                               rwCD011A.Item("Description"),
                                                               rwCD011A.Item("ServiceType")})
                            End If
                        Next
                    Else
                        otherDao.ExecNqry(_DAL.DeleteCD011, MFCode)
                        otherDao.ExecNqry(_DAL.DeleteCD011A, MFCode)
                        otherDao.ExecNqry(_DAL.InsertCD011, New Object() {tbCD011.Rows(0).Item("CodeNo"),
                                                               tbCD011.Rows(0).Item("Description"),
                                                                tbCD011.Rows(0).Item("RefNo"),
                                                                CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                                                                otherLoginInfo.EntryName,
                                                               tbCD011.Rows(0).Item("ServiceType"),
                                                                tbCD011.Rows(0).Item("StopFlag")})
                        For Each rwCD011A As DataRow In tbCD011A.Rows
                            If Integer.Parse(otherDao.ExecSclr(_DAL.QueryDetailExists, New Object() {MFCode, rwCD011A.Item("CodeNo")})) = 0 Then
                                otherDao.ExecNqry(_DAL.InsertCD011A, New Object() {MFCode,
                                                               rwCD011A.Item("CodeNo"),
                                                               rwCD011A.Item("Description"),
                                                               rwCD011A.Item("ServiceType")})
                            End If
                        Next

                        strRetMsg.AppendLine(String.Format(Language.CopyOK, otherLoginInfo.CompCode))
                    End If
                    Othertrans.Commit()
                Catch ex As Exception
                    Othertrans.Rollback()
                    strRetMsg.AppendLine(String.Format(Language.CopyErr, otherLoginInfo.CompCode, ex.ToString))
                Finally
                    otherLoginInfo = Nothing
                    If Othertrans IsNot Nothing Then
                        Othertrans.Dispose()
                        Othertrans = Nothing
                    End If
                    If otherDao IsNot Nothing Then
                        otherDao.Dispose()
                        otherDao = Nothing
                    End If
                End Try
            Next
            result.ResultBoolean = True
            result.ResultXML = strRetMsg.ToString
        Catch ex As Exception
            result.ResultBoolean = False
            result.ErrorCode = -99
            result.ErrorMessage = ex.ToString
            Return result
        Finally
            If tbCD011 IsNot Nothing Then
                tbCD011.Dispose()
                tbCD011 = Nothing
            End If
            If tbCD011A IsNot Nothing Then
                tbCD011A.Dispose()
                tbCD011A = Nothing
            End If
            If blnAutoClose Then
                If DAO IsNot Nothing Then
                    DAO.AutoCloseConn = True
                End If

                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
            End If
        End Try
        Return result
    End Function
    Public Function Execute(ByVal EditMode As EditMode, ByVal dsMaster As DataSet) As RIAResult
        Dim result As New RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing

        Dim blnAutoClose As Boolean = False
        Dim dsReturn As New DataSet
        Dim dtOK As DataTable = Nothing
        Dim dtError As DataTable = Nothing
        Dim dsResult As New DataSet
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
        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, Me.LoginInfo.EntryId)
        Dim tbCD011A As DataTable = Nothing
     
        Try

            Select Case EditMode
                Case CableSoft.BLL.Utility.EditMode.Edit
                    DAO.ExecNqry(_DAL.UpdateCD011, New Object() {dsMaster.Tables("CD011").Rows(0).Item("Description"),
                                                              dsMaster.Tables("CD011").Rows(0).Item("RefNo"),
                                                                 CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                                                                 Me.LoginInfo.EntryName,
                                                                 dsMaster.Tables("CD011").Rows(0).Item("ServiceType"),
                                                                 dsMaster.Tables("CD011").Rows(0).Item("StopFlag"),
                                                                 dsMaster.Tables("CD011").Rows(0).Item("CodeNo")})
                Case CableSoft.BLL.Utility.EditMode.Append
                
                    DAO.ExecNqry(_DAL.InsertCD011, New Object() {dsMaster.Tables("CD011").Rows(0).Item("CodeNo"),
                                                                dsMaster.Tables("CD011").Rows(0).Item("Description"),
                                                                 dsMaster.Tables("CD011").Rows(0).Item("RefNo"),
                                                                 CableSoft.BLL.Utility.DateTimeUtility.GetDTString(FNowDate),
                                                                 Me.LoginInfo.EntryName,
                                                                dsMaster.Tables("CD011").Rows(0).Item("ServiceType"),
                                                                 dsMaster.Tables("CD011").Rows(0).Item("StopFlag")})
                   
            End Select
            '更新CD011A
            DAO.ExecNqry(_DAL.DeleteCD011A, New Object() {dsMaster.Tables("CD011").Rows(0).Item("CodeNo")})
            For Each rw As DataRow In dsMaster.Tables("CD011A").Rows
                'Return String.Format("Insert into CD011A (MFCode,CodeNo,Description,ServiceType " & _
                '                      " Values ({0}0,{0}1,{0}2,{0}3)", Sign)
                DAO.ExecNqry(_DAL.InsertCD011A, New Object() {dsMaster.Tables("CD011").Rows(0).Item("CodeNo"),
                                                              rw.Item("CodeNo"),
                                                              rw.Item("Description"),
                                                              dsMaster.Tables("CD011").Rows(0).Item("ServiceType")})

            Next
            tbCD011A = GetCD011A()
            tbCD011A.TableName = "CD011A"
            dsResult.Tables.Add(tbCD011A.Copy)
            trans.Commit()
            result.ResultBoolean = True
            result.ResultDataSet = dsResult

        Catch ex As Exception
            trans.Rollback()
            result.ErrorCode = -99
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            Return result
        Finally
            If tbCD011A IsNot Nothing Then
                tbCD011A.Dispose()
                tbCD011A = Nothing
            End If
            If cmd IsNot Nothing Then
                cmd.Dispose()
            End If

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
        Return result
    End Function
    Public Function QueryCD011B(ByVal ServiceType As String) As DataSet
        Dim dsResult As New DataSet
        Try

            dsResult.Tables.Add(DAO.ExecQry(_DAL.QueryCD011B(ServiceType)).Copy)
        Catch ex As Exception
            Throw ex
        End Try
        Return dsResult
    End Function
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
    Public Function GetServiceType() As DataTable
        Return DAO.ExecQry(_DAL.GetServiceType)
    End Function
    Public Function GetCD011A() As DataTable
        Return DAO.ExecQry(_DAL.GetCD011A)
    End Function
    Public Function GetSO041() As DataTable
        Return DAO.ExecQry(_DAL.GetSO041, New Object() {LoginInfo.CompCode})
    End Function
    Public Function GetMaxCode() As String
        Return DAO.ExecSclr(_DAL.GetMaxCode).ToString
    End Function
    Public Function GetCD011Sechema() As DataTable
        Return DAO.ExecQry(_DAL.GetCD011Sechema)
    End Function
    Public Function GetAllData() As DataSet
        Dim dsResult As New DataSet
        Dim tbCompCode As DataTable = Nothing
        Dim tbServiceType As DataTable = Nothing
        Dim tbCD011A As DataTable = Nothing
        Dim tbSO041 As DataTable = Nothing
        Dim tbCD011Sechema As DataTable = Nothing
        Dim objUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
        Dim tbPriv As DataTable = Nothing
        Try
            tbCompCode = GetCompCode()
            tbCompCode.TableName = "CompCode"
            tbServiceType = GetServiceType()
            tbServiceType.TableName = "ServiceType"
            tbCD011A = GetCD011A()
            tbCD011A.TableName = "CD011A"
            tbPriv = objUtility.GetPriv(Me.LoginInfo.EntryId, "SO6250")
            tbPriv.TableName = "Priv"
            tbSO041 = GetSO041()
            tbSO041.TableName = "SO041"
            tbCD011Sechema = GetCD011Sechema()
            tbCD011Sechema.TableName = "CD011"
            dsResult.Tables.Add(tbCompCode.Copy)
            dsResult.Tables.Add(tbServiceType.Copy)
            dsResult.Tables.Add(tbCD011A.Copy)
            dsResult.Tables.Add(tbPriv.Copy)
            dsResult.Tables.Add(tbSO041.Copy)
            dsResult.Tables.Add(tbCD011Sechema.Copy)
        Catch ex As Exception
            Throw
        Finally
            If tbCompCode IsNot Nothing Then
                tbCompCode.Dispose()
                tbCompCode = Nothing
            End If
            If tbServiceType IsNot Nothing Then
                tbServiceType.Dispose()
                tbServiceType = Nothing
            End If
            If tbCD011A IsNot Nothing Then
                tbCD011A.Dispose()
                tbCD011A = Nothing
            End If
            If objUtility IsNot Nothing Then
                objUtility.Dispose()
            End If
        End Try
        Return dsResult
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
                If _DAL IsNot Nothing Then
                    _DAL.Dispose()
                    _DAL = Nothing
                End If
                If Language IsNot Nothing Then
                    Language.Dispose()
                    Language = Nothing
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
