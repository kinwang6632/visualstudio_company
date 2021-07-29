Imports System.Data.Common
Imports CableSoft.BLL.Utility
Public Class SaveData
    Inherits BLLBasic
    Implements IDisposable
    Private Const fMaintain_Wip As String = "Wip"
    Private Const fMaintain_Facility As String = "Facility"
    Private Const fMaintain_PRFacility As String = "PRFacility"
    Private Const fMaintain_Charge As String = "Charge"
    Private Const fMaintain_ChangeFacility As String = "ChangeFacility"
    Private Const fMaintain_Parameter As String = "WipPara"
    Private Language As New CableSoft.BLL.Language.SO61.MaintainLanguage
    Private _DAL As New MaintainDALMultiDB(Me.LoginInfo.Provider)
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Function Save(ByVal EditMode As EditMode,
                         ByVal ShouldReg As Boolean,
                         ByVal WipData As DataSet) As Boolean
        Return Me.Save(EditMode, ShouldReg, WipData, False).ResultBoolean
    End Function
    Public Function Save(ByVal EditMode As EditMode,
                         ByVal ShouldReg As Boolean,
                        ByVal WipData As DataSet, ByVal ReturnFlag As Boolean) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aRiaresult As New RIAResult()
        Dim tbSO008Log As DataTable = Nothing
        aRiaresult.ResultBoolean = False
        aRiaresult.ErrorCode = -99
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

        CSLog = New CableSoft.SO.BLL.DataLog.DataLog(Me.LoginInfo, Me.DAO)
        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans
        If blnAutoClose Then
            Dim aAction As String = Nothing
            Select Case EditMode
                Case CableSoft.BLL.Utility.EditMode.Edit
                    aAction = Language.EditClientInfo
                Case CableSoft.BLL.Utility.EditMode.Append
                    aAction = Language.AddClientInfo
                Case Else
                    aAction = Language.EditClientInfo
            End Select
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, aAction)
        End If        
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.SaveData(Me.LoginInfo, DAO)
        Dim objSF As New CableSoft.SO.BLL.Utility.Wip(Me.LoginInfo, DAO)
        Dim aNowDate As Date = Date.Now
        Try
            If EditMode <> CableSoft.BLL.Utility.EditMode.Append Then
                'tbSO008Log = DAO.ExecQry("Select A.rowid,A.* From SO008  A where SNO = '" & WipData.Tables(fMaintain_Wip).Rows(0).Item("SNO") & "'")
                tbSO008Log = DAO.ExecQry(_DAL.QuerySO008Log, New Object() {WipData.Tables(fMaintain_Wip).Rows(0).Item("SNO")})
                If tbSO008Log.Rows.Count = 0 Then
                    tbSO008Log.Dispose()
                    tbSO008Log = Nothing
                End If
            End If
            With WipData.Tables(fMaintain_Wip).Rows(0)
                .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(Now)
                .Item("NewUpdTime") = Now
            End With

            If Not objWipUtility.ChangeWip(EditMode, BLL.Utility.InvoiceType.Maintain, WipData, ShouldReg) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeWipErr)

                aRiaresult.ErrorMessage = Language.ChangeWipErr
                Return aRiaresult
            End If

            If Not objWipUtility.ChangeCharge(EditMode, ShouldReg, WipData) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeChargeErr)
                aRiaresult.ErrorMessage = Language.ChangeChargeErr
                Return aRiaresult
            End If



            If Not objWipUtility.ChangeCommandData(WipData.Tables(fMaintain_Wip).Rows(0).Item("SNO"),
                                                   WipData.Tables(fMaintain_Wip)) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeCommandDataErr)
                aRiaresult.ErrorMessage = Language.ChangeCommandDataErr
                Return aRiaresult
            End If

            With WipData.Tables(fMaintain_Wip).Rows(0)
                objSF.SF_NEWWIP2(cmd, Int32.Parse(.Item("CustId")),
                            Me.LoginInfo.CompCode, .Item("ServiceType"))
            End With

            If Not objWipUtility.ChangeWipFinalProcess(EditMode, BLL.Utility.InvoiceType.Maintain, WipData) Then
                Throw New Exception("WipUtil.ChangeWipFinalProcess")
            End If

            Using ControlCMD As New CableSoft.SO.BLL.Wip.ControlCommand.ControlCommand(LoginInfo, DAO)
                Dim RiaSendCmd As New RIAResult

                RiaSendCmd = ControlCMD.Execute(EditMode, WipData)

                If Not RiaSendCmd.ResultBoolean Then
                    If Not String.IsNullOrEmpty(RiaSendCmd.ErrorMessage) Then
                        Throw New Exception("SendCmd Error: " & RiaSendCmd.ErrorMessage)
                    Else
                        Throw New Exception(Language.SendCmdNotRetValue)
                    End If
                Else
                    WipData = RiaSendCmd.ResultDataSet
                End If

            End Using
            If tbSO008Log IsNot Nothing Then
                CableSoft.BLL.Utility.Utility.CopyDataRow(WipData.Tables(fMaintain_Wip).Rows(0), tbSO008Log.Rows(0))
                'Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update, "SO106", Account.Tables(FNewAccountTableName), Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                Dim aResult As RIAResult = CSLog.SummaryExpansion(cmd, DataLog.OpType.Update, "SO008", tbSO008Log, Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                If Not aResult.ResultBoolean Then
                    Select Case aResult.ErrorCode
                        Case -5
                        Case -6
                            If blnAutoClose Then
                                trans.Rollback()
                                Throw New Exception(aResult.ErrorMessage)
                                'Return aResult
                            End If

                    End Select

                End If
            End If
            If blnAutoClose Then
                trans.Commit()
            End If

        Catch ex As Exception
            If blnAutoClose Then
                trans.Rollback()
            End If
            Throw ex
            aRiaresult.ErrorCode = -1
            aRiaresult.ErrorMessage = ex.ToString
        Finally

            If objWipUtility IsNot Nothing Then
                objWipUtility.Dispose()
                objWipUtility = Nothing
            End If
            If objSF IsNot Nothing Then
                objSF.Dispose()
                objSF = Nothing
            End If
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
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
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                    CSLog = Nothing
                End If
            End If
        End Try
        If tbSO008Log IsNot Nothing Then
            tbSO008Log.Dispose()
            tbSO008Log = Nothing
        End If
        aRiaresult.ResultXML = WipData.Tables(fMaintain_Wip).Rows(0).Item("SNO")
        aRiaresult.ErrorCode = 0
        aRiaresult.ResultBoolean = True
        Return aRiaresult
    End Function
    Public Function Save1(ByVal EditMode As EditMode,
                         ByVal ShouldReg As Boolean,
                         ByVal ServCode As String,
                         ByVal OldSNo As String,
                         ByVal WipData As DataSet) As Boolean
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        'Dim aRiaresult As New RIAResult()

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

        'trans = cn.BeginTransaction
        Dim cmd As DbCommand = cn.CreateCommand
        cmd.Connection = cn
        cmd.Transaction = trans
        'DAO.AutoCloseConn = False
        'DAO.Transaction = trans
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.SaveData(Me.LoginInfo, DAO)
        Dim objSF As New CableSoft.SO.BLL.Utility.Wip(Me.LoginInfo, DAO)
        Dim aNowDate As Date = Date.Now

        'aRiaresult.ResultBoolean = False
        'aRiaresult.ErrorCode = -1
        Try
            With WipData.Tables(fMaintain_Wip).Rows(0)
                .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(Now)
                '.Item("UpdEn") = Me.LoginInfo.EntryName
                '.Item("AcceptEn") = Me.LoginInfo.EntryId
                '.Item("AcceptName") = Me.LoginInfo.EntryName
                '.Item("AcceptTime") = New Date(aNowDate.Year, aNowDate.Month,
                '                                  aNowDate.Day, aNowDate.Hour, aNowDate.Minute, 0)
                'If .IsNull("AcceptTime") Then

                'End If

            End With

            If Not objWipUtility.ChangeWip(EditMode, BLL.Utility.InvoiceType.Maintain, WipData) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeWipErr)
                'aRiaresult.ErrorMessage = "ChangeWip 失敗！"
                'Return aRiaresult
            End If
            If Not objWipUtility.ChangeFacility(EditMode, BLL.Utility.InvoiceType.Maintain, WipData) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeFacilityErr)
                'aRiaresult.ErrorMessage = "ChangeFacility 失敗！"
                'Return aRiaresult
            End If
            If Not objWipUtility.ChangePRFacility(EditMode, BLL.Utility.InvoiceType.Maintain, WipData) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangePRFacilityErr)
                'aRiaresult.ErrorMessage = "ChangePRFacility 失敗！"
                'Return aRiaresult
            End If
            'If Not objWipUtility.ChangeChangeFacility(EditMode, BLL.Utility.InvoiceType.Maintain, WipData) Then
            '    If blnAutoClose Then
            '        trans.Rollback()
            '    End If
            '    Throw New Exception(Language.ChangeChangeFacilityErr)
            '    'aRiaresult.ErrorMessage = "ChangeChangeFacility 失敗！"
            '    'Return aRiaresult
            'End If
            If Not objWipUtility.ChangeCharge(EditMode, ShouldReg, WipData) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeChargeErr)
                'aRiaresult.ErrorMessage = "ChangeCharge 失敗！"
                'Return aRiaresult
            End If
            If String.IsNullOrEmpty(ServCode) Then
                'ServCode = DAO.ExecSclr(String.Format("SELECT ServCode FROM SO001 WHERE CUSTID ={0}",
                '                                     WipData.Tables(fMaintain_Wip).Rows(0)("CustId")))
                ServCode = DAO.ExecSclr(_DAL.QueryServCode, New Object() {WipData.Tables(fMaintain_Wip).Rows(0)("CustId")})
            End If
            With WipData.Tables(fMaintain_Wip).Rows(0)
                Dim CloseWip As Boolean = False
                If (Not .IsNull("SignDate")) AndAlso (Not String.IsNullOrEmpty(.Item("SignDate"))) Then
                    CloseWip = True
                End If
                If Not objWipUtility.ChangeResvDetail(.Item("SNO"), Me.LoginInfo.CompCode,
                                     Date.Parse(.Item("ResvTime")), ServCode, .Item("ServiceType"),
                                    CloseWip) Then
                    If blnAutoClose Then
                        trans.Rollback()
                    End If
                    Throw New Exception(Language.ChangeResvDetailErr)
                    'aRiaresult.ErrorMessage = "ChangeResvDetail 失敗！"
                    'Return aRiaresult
                End If
            End With
            If Not objWipUtility.ChangeResvLog(BLL.Utility.InvoiceType.Maintain, WipData.Tables(fMaintain_Wip)) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeResvLogErr)
                'aRiaresult.ErrorMessage = "ChangeResvLog 失敗！"
                'Return aRiaresult
            End If
            If Not objWipUtility.ChangeCommandData(OldSNo, WipData.Tables(fMaintain_Wip)) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeCommandDataErr)
                'aRiaresult.ErrorMessage = "ChangeCommandData 失敗！"
                'Return aRiaresult
            End If
            If Not objWipUtility.ChangeResvTempPoint(BLL.Utility.InvoiceType.Maintain, WipData.Tables(fMaintain_Wip)) Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.ChangeResvTempPoint)
                'aRiaresult.ErrorMessage = "ChangeResvTempPoint 失敗！"
                'Return aRiaresult
            End If
            If Not objWipUtility.DelResvPoint Then
                If blnAutoClose Then
                    trans.Rollback()
                End If
                Throw New Exception(Language.DelResvPoint)
                'aRiaresult.ErrorMessage = "DelResvPoint 失敗！"

            End If
            With WipData.Tables(fMaintain_Wip).Rows(0)
                objSF.SF_NEWWIP2(cmd, Int32.Parse(.Item("CustId")),
                            Me.LoginInfo.CompCode, .Item("ServiceType"))
            End With


            If Not objWipUtility.ChangeWipFinalProcess(EditMode, BLL.Utility.InvoiceType.Maintain, WipData) Then
                Throw New Exception("WipUtil.ChangeWipFinalProcess")
            End If

            Using ControlCMD As New CableSoft.SO.BLL.Wip.ControlCommand.ControlCommand(LoginInfo, DAO)
                Dim RiaSendCmd As New RIAResult

                RiaSendCmd = ControlCMD.Execute(EditMode, WipData)

                If Not RiaSendCmd.ResultBoolean Then
                    If Not String.IsNullOrEmpty(RiaSendCmd.ErrorMessage) Then
                        Throw New Exception("SendCmd Error: " & RiaSendCmd.ErrorMessage)
                    Else
                        Throw New Exception(Language.SendCmdNotRetValue)
                    End If
                Else
                    WipData = RiaSendCmd.ResultDataSet
                End If

            End Using

            If blnAutoClose Then
                trans.Commit()
            End If

        Catch ex As Exception
            If blnAutoClose Then
                trans.Rollback()
            End If
            Throw ex
            'aRiaresult.ErrorCode = -99
            'aRiaresult.ErrorMessage = ex.Message
        Finally
            'aRiaresult.ResultXML = WipData.Tables(fMaintain_Wip).Rows(0).Item("SNO")
            If _DAL IsNot Nothing Then
                _DAL.Dispose()
                _DAL = Nothing
            End If
            If objWipUtility IsNot Nothing Then
                objWipUtility.Dispose()
                objWipUtility = Nothing
            End If
            If objSF IsNot Nothing Then
                objSF.Dispose()
                objSF = Nothing
            End If
            If cmd IsNot Nothing Then
                cmd.Dispose()
                cmd = Nothing
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
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                    CSLog = Nothing
                End If
            End If


        End Try

        Return True
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
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
