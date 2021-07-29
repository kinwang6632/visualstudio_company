Imports System.Data.Common
Imports CableSoft.BLL.Utility
Public Class Validate
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New MaintainDALMultiDB(Me.LoginInfo.Provider)

    Private Const fMaintain_Wip As String = "Wip"
    Private Const fMaintain_Facility As String = "Facility"
    Private Const fMaintain_PRFacility As String = "PRFacility"
    Private Const fMaintain_Charge As String = "Charge"
    Private Const fMaintain_ChangeFacility As String = "ChangeFacility"
    Private Const fMaintain_OldWip As String = "OldWip"
    Private Language As New CableSoft.BLL.Language.SO61.MaintainLanguage
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
    Public Function ChkCanResv(ByVal ServCode As String, ByVal WipCode As Int32,
                            ByVal MCode As String, ByVal ServiceType As String,
                            ByVal ResvTime As Date,
                            ByVal AcceptTime As Date, ByVal OldResvTime As Date,
                            ByVal Resvdatebefore As Int32, ByVal WorkUnit As Decimal,
                            ByVal IsBooking As Boolean) As RIAResult
        Return ChkCanResv(ServCode, WipCode, MCode, ServiceType, ResvTime,
                         AcceptTime, OldResvTime, Resvdatebefore, WorkUnit, IsBooking, ServCode)
    End Function
    ''' <summary>
    ''' 檢查預約時段是否可派工
    ''' </summary>
    ''' <param name="ServCode">服務區</param>
    ''' <param name="WipCode"></param>
    ''' <param name="MCode">裝機類別名稱</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="ResvTime">預約時間</param>
    ''' <param name="AcceptTime">受理時間</param>
    ''' <param name="OldResvTime">舊預約時間</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChkCanResv(ByVal ServCode As String, ByVal WipCode As Int32,
                            ByVal MCode As String, ByVal ServiceType As String,
                            ByVal ResvTime As Date,
                            ByVal AcceptTime As Date, ByVal OldResvTime As Date,
                            ByVal Resvdatebefore As Int32, ByVal WorkUnit As Decimal,
                            ByVal IsBooking As Boolean, ByVal oldServCode As String) As RIAResult

        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Dim aRet As RIAResult = obj.ChkCanResv(BLL.Utility.InvoiceType.Maintain,
                                  WipCode, ServCode, MCode, ServiceType,
                                  ResvTime, AcceptTime, OldResvTime, Resvdatebefore, WorkUnit, IsBooking, Nothing, oldServCode)
            Return aRet
        Finally
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If
        End Try
    End Function
    Public Function ChkCloseDate(CloseDate As String, _
        ServiceType As String) As RIAResult

        Dim aRet As New RIAResult()
        aRet.ResultBoolean = True
        Dim aCloseDate As Date = Date.Now
        Dim intDayCut As Int32 = DAO.ExecSclr(_DAL.QueryDayCut,
                                              New Object() {Me.LoginInfo.CompCode})
        Dim aWhere As String = String.Empty

        'If Not String.IsNullOrEmpty(ServiceType) Then
        '    aWhere = String.Format(" AND SERVICETYPE='{0}'", ServiceType)
        'End If
        'Dim strTranDate As String = DAO.ExecSclr(String.Format("SELECT TRANDATE FROM SO062 WHERE COMPCODE={0}0 " & _
        '                                                     aWhere & " AND TYPE = 1 ORDER BY TRANDATE DESC ", _DAL.Sign),
        '                                                 New Object() {Me.LoginInfo.CompCode}).ToString

        Dim strTranDate As String = DAO.ExecSclr(_DAL.QueryTranDate(ServiceType), New Object() {Me.LoginInfo.CompCode}).ToString
        'Dim intPara6 As Integer = Int32.Parse(DAO.ExecSclr(String.Format("SELECT NVL(PARA6,0) FROM SO043 WHERE " & _
        '                                     "COMPCODE = {0}0 " & aWhere, _DAL.Sign), New Object() {Me.LoginInfo.CompCode}))
        Dim intPara6 As Integer = Integer.Parse(DAO.ExecSclr(_DAL.QueryPara6(ServiceType), New Object() {Me.LoginInfo.CompCode}))
        If String.IsNullOrEmpty(CloseDate) Then
            aRet.ErrorCode = -1
            aRet.ErrorMessage = Language.IsNullField
            aRet.ResultBoolean = False
            Return aRet
        End If
        If Not Date.TryParse(CloseDate, aCloseDate) Then
            aRet.ErrorCode = -2
            aRet.ErrorMessage = Language.DateFmtErr
            aRet.ResultBoolean = False
            Return aRet
        End If
        If aCloseDate > Date.Now Then
            aRet.ErrorCode = -3
            aRet.ErrorMessage = Language.DateExceed
            aRet.ResultBoolean = False
            Return aRet
        End If
        If (Date.Now - aCloseDate).Days > intPara6 Then
            aRet.ErrorCode = -4
            aRet.ErrorMessage = Language.DateExceedSave
            aRet.ResultBoolean = False
            Return aRet
        End If
        If Not String.IsNullOrEmpty(strTranDate) Then
            Dim aTranDate As Date = Date.Parse(strTranDate)
            If intDayCut = 1 AndAlso Date.Parse(CloseDate) >= aTranDate Then
                Return aRet
            Else
                If Date.Parse(CloseDate) >= aTranDate Then

                Else
                    aRet.ErrorCode = -5
                    aRet.ErrorMessage = Language.DateHasClose
                    aRet.ResultBoolean = False
                End If


                Return aRet
            End If
        End If
        Return aRet


    End Function
    Public Function ChkSaveDataOK(ByVal EditMode As EditMode, ByVal CloseDate As String, _
        ServiceType As String, ByVal WipData As DataSet, ByVal ShouldReg As Boolean) As RIAResult
        Dim aResultRia As New RIAResult
        Me.DAO.AutoCloseConn = False
        Try
            aResultRia = Me.ChkDataOk(EditMode, WipData, False)
            If aResultRia.ResultBoolean Then
                aResultRia = Me.ChkCloseDate(CloseDate, ServiceType)
            End If

        Catch ex As Exception
            aResultRia.ResultBoolean = False
            aResultRia.ErrorMessage = ex.ToString
            aResultRia.ErrorCode = -99
        Finally
            Me.DAO.AutoCloseConn = True
        End Try
        Return aResultRia
    End Function
    Public Function ChkDataOk(ByVal EditMode As EditMode, ByVal WipData As DataSet) As RIAResult
        Return ChkDataOk(EditMode, WipData, False)
    End Function
    ''' <summary>
    ''' 檢查派工單是否正確
    ''' </summary>
    ''' <param name="WipData">派工資料</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function ChkDataOk(ByVal EditMode As EditMode, ByVal WipData As DataSet, ByVal ShouldReg As Boolean) As RIAResult
        Dim aRet As New RIAResult()
        Dim WorkCode As DataTable = Nothing
        Dim WipSystem As DataTable = Nothing
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Validate(Me.LoginInfo, DAO)
        Dim aNowDate As Date = Now.Date
        Try
            aRet.ResultBoolean = True
            If WipData.Tables(fMaintain_Wip).Rows.Count <= 0 Then
                aRet.ResultBoolean = False
                aRet.ErrorCode = -1
                aRet.ErrorMessage = Language.NoMaintainData
                Return aRet
            End If
            Dim RefNo As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.QueryCD006RefNo,
                                                          New Object() {WipData.Tables(fMaintain_Wip).Rows(0).Item("ServiceCode")}))
            WorkCode = DAO.ExecQry(_DAL.GetCD006,
                           New Object() {WipData.Tables(fMaintain_Wip).Rows(0).Item("ServiceCode")})
            If (RefNo = 1) Then
                If WipData.Tables(fMaintain_ChangeFacility).Rows.Count <= 0 Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -2
                    aRet.ErrorMessage = Language.MustChangeFaci
                    Return aRet
                End If
            End If
            If (Not WipData.Tables(fMaintain_Wip).Rows(0).IsNull("FinTime")) Then
                If Int32.Parse(DAO.ExecSclr(_DAL.QueryCheckMFCode,
                                            New Object() {WipData.Tables(fMaintain_Wip).Rows(0).Item("SERVICETYPE")})) = 1 Then
                    If WipData.Tables(fMaintain_Wip).Rows(0).IsNull("MFCode1") Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -3
                        aRet.ErrorMessage = Language.MustFixNum1
                        Return aRet
                    End If
                    If WipData.Tables(fMaintain_Wip).Rows(0).IsNull("MFCode2") Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -4
                        aRet.ErrorMessage = Language.MustFixNum2
                        Return aRet
                    End If
                End If
            End If
            aRet.ResultBoolean = True
            aRet.ErrorCode = 0
            aRet.ErrorMessage = String.Empty
            WipSystem = DAO.ExecQry(_DAL.GetSO042,
                                                New Object() {WipData.Tables(fMaintain_Wip).Rows(0).Item("ServiceType")})
            With WipData
                aRet = obj.ChkDataOk(EditMode, BLL.Utility.InvoiceType.Maintain, .Tables(fMaintain_Wip), .Tables(fMaintain_OldWip), _
                              WorkCode, .Tables(fMaintain_Facility), .Tables(fMaintain_PRFacility), _
                              .Tables(fMaintain_Charge), Nothing, WipSystem, .Tables(fMaintain_ChangeFacility), ShouldReg)


            End With

        Catch ex As Exception
            aRet.ResultBoolean = False
            aRet.ErrorCode = -99
            aRet.ErrorMessage = ex.ToString
        Finally
            'WipData.Dispose()
            If WipSystem IsNot Nothing Then
                WipSystem.Dispose()
                WipSystem = Nothing
            End If
            If WorkCode IsNot Nothing Then
                WorkCode.Dispose()
                WorkCode = Nothing
            End If
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If

        End Try
        Return aRet

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
