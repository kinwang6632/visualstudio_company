Imports System.Data.Common
Imports CableSoft.BLL.Utility
'Imports System.Windows.Forms
Public Class Maintain
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New MaintainDALMultiDB(Me.LoginInfo.Provider)

    Private Const fMaintain_Wip As String = "Wip"
    Private Const fMaintain_Facility As String = "Facility"
    Private Const fMaintain_PRFacility As String = "PRFacility"
    Private Const fMaintain_Charge As String = "Charge"
    Private Const fMaintain_ChangeFacility As String = "ChangeFacility"
    Private Const fCommon_SO001 As String = "SO001"
    Private Const fCommon_SO002 As String = "SO002"
    Private Const fCommon_MaintainCode As String = "MaintainCode"
    Private Const fCommon_GroupCode As String = "GroupCode"
    Private Const fCommon_Priv As String = "Priv"
    Private Const fCommon_PrivSO1132 As String = "PrivSO1132"
    Private Const fCommon_FaciFinishPrivFlag As String = "FaciFinishPrivFlag"
    Private Const fCommon_SO014 As String = "SO014"
    Private Const fCommon_ReInstAddrNo As String = "ReInstAddrNo"
    Private Const fCommon_ServiceType As String = "ServiceType"
    Private Const fCommon_WorkerEn1 As String = "WorkerEn1"
    Private Const fCommon_WorkerEn2 As String = "WorkerEn2"
    Private Const fCommon_FalseSNo As String = "FalseSNo"
    Private Const fCommon_ReturnCode As String = "ReturnCode"
    Private Const fCommon_ReturnDescCode As String = "ReturnDescCode"
    Private Const fCommon_SatiCode As String = "SatiCode"
    Private Const fCommon_SignEn As String = "SignEn"
    Private Const fCommon_MFCode As String = "MFCode"
    Private Const fCommon_MFCode2 As String = "MFCode2"
    Private Const fChooseFacility_ALL As String = "ChooseFacility"
    Private Const fChooseFacility_One As String = "Facility"
    Private Const fChooseCharge As String = "Simple"
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
    'Public Function GetServiceTypeChangeData(
    '                                         ByVal CustId As Integer,
    '                                         ByVal ServiceType As String,
    '                                         ByVal IsGetFalseSNo As Boolean) As DataSet
    '    Dim dsResult As New DataSet
    '    Try
    '        Me.DAO.AutoCloseConn = False
    '        Dim dtSO001 As DataTable = Me.GetSO001(CustId)
    '        dtSO001.TableName = fCommon_SO001
    '        Dim dtCustomer As DataTable = Me.GetCustomer(CustId, ServiceType)
    '        dtCustomer.TableName = fCommon_SO002
    '        Dim dtFalseSNO As New DataTable(fCommon_FalseSNo)
    '        dtFalseSNO.Columns.Add("SNO", GetType(String))
    '        Dim dtMaintainCode As DataTable = Me.GetMaitainCode(ServiceType)
    '        dtMaintainCode.TableName = fCommon_MaintainCode
    '        Dim dtReturnCode As DataTable = Me.GetReturnCode(ServiceType)
    '        dtReturnCode.TableName = fCommon_ReturnCode
    '        Dim dtReturnDescCode As DataTable = Me.GetReturnDescCode(ServiceType)
    '        dtReturnDescCode.TableName = fCommon_ReturnDescCode
    '        Dim dtSatiCode As DataTable = Me.GetSatiCode(ServiceType)
    '        dtSatiCode.TableName = fCommon_SatiCode
    '        Dim dtMFCode As DataTable = Me.GetMFCode1(ServiceType)
    '        dtMFCode.TableName = fCommon_MFCode
    '        If IsGetFalseSNo Then
    '            Dim rwNew As DataRow = dtFalseSNO.NewRow
    '            rwNew.Item("SNO") = Me.GetFalseSNO(ServiceType).ResultXML
    '            dtFalseSNO.Rows.Add(rwNew)
    '            dtFalseSNO.AcceptChanges()
    '        End If

    '        dsResult.Tables.Add(dtCustomer.Copy)
    '        dsResult.Tables.Add(dtFalseSNO.Copy)
    '        dsResult.Tables.Add(dtMaintainCode.Copy)
    '        dsResult.Tables.Add(dtMFCode.Copy)
    '        dsResult.Tables.Add(dtReturnCode.Copy)
    '        dsResult.Tables.Add(dtReturnDescCode.Copy)
    '        dsResult.Tables.Add(dtSatiCode.Copy)
    '        dsResult.Tables.Add(dtSO001.Copy)
    '    Catch ex As Exception
    '        Throw
    '    Finally
    '        Me.DAO.AutoCloseConn = True
    '    End Try
    '    Return dsResult
    'End Function
    Public Function GetServiceCodeChangeData(ByVal CustId As Integer,
                                             ByVal ServiceType As String,
                                             ByVal ServiceCode As String,
                                             ByVal AcceptTime As String,
                                             ByVal ResvTime As String,
                                            ByVal IsHaveContactTable As Boolean) As DataSet
        'Me.GetDefaultResvTime()
        Return Nothing
    End Function


    Public Function GetServiceTypeChangeData(ByVal SNo As String,
                                             ByVal CustId As Integer,
                                             ByVal ServiceType As String,
                                             ByVal ServiceCode As String,
                                             ByVal MFCode As String,
                                             ByVal IsGetFalseSNo As Boolean) As DataSet
        Dim dsResult As New DataSet
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim dtWipName As New DataTable("WipTableName")
        dtWipName.Columns.Add("WipName", GetType(String))
        Dim dtFalseSNO As New DataTable(fCommon_FalseSNo)
        dtFalseSNO.Columns.Add("SNO", GetType(String))
        Try
            Me.DAO.AutoCloseConn = False
            If IsGetFalseSNo Then
                Dim rwNew As DataRow = dtFalseSNO.NewRow
                rwNew.Item("SNO") = Me.GetFalseSNO(ServiceType).ResultXML
                SNo = rwNew.Item("SNO")
                dtFalseSNO.Rows.Add(rwNew)
                dtFalseSNO.AcceptChanges()
            End If
            Using dsWip As DataSet = obj.GetWipDetail(SNo, False, BLL.Utility.InvoiceType.Maintain)
                If dsWip.Tables(fMaintain_Wip).Rows.Count > 0 Then
                    ServiceType = dsWip.Tables(fMaintain_Wip).Rows(0).Item("ServiceType")
                    ServiceCode = dsWip.Tables(fMaintain_Wip).Rows(0).Item("ServiceCode").ToString
                    CustId = Integer.Parse(dsWip.Tables(fMaintain_Wip).Rows(0).Item("CustId").ToString)
                    MFCode = dsWip.Tables(fMaintain_Wip).Rows(0).Item("MFCode1").ToString
                End If

                For Each tbWip As DataTable In dsWip.Tables
                    Dim rwWip As DataRow = dtWipName.NewRow
                    rwWip.Item("WipName") = tbWip.TableName
                    dtWipName.Rows.Add(rwWip)
                    dsResult.Tables.Add(tbWip.Copy)
                Next
                Using tbFaciFinishPrivFlag As DataTable = ChkFaciFinishPrivFlag(dsWip)
                    dsResult.Tables.Add(tbFaciFinishPrivFlag.Copy)
                End Using
                Dim tbHaveCM As New DataTable("HaveCM")

                tbHaveCM.Columns.Add("ResultBoolean", GetType(Boolean))
                tbHaveCM.Columns.Add("ErrorCode", GetType(Integer))
                tbHaveCM.Columns.Add("ErrorMessage", GetType(String))
                tbHaveCM.Columns.Add("ResultXML", GetType(String))
                Dim aRiaResult As RIAResult = ChkHaveCM(dsWip, False, "I")
                Dim rw As DataRow = tbHaveCM.NewRow
                rw.Item("ResultBoolean") = aRiaResult.ResultBoolean
                rw.Item("ErrorCode") = aRiaResult.ErrorCode
                rw.Item("ErrorMessage") = aRiaResult.ErrorMessage
                rw.Item("ResultXML") = aRiaResult.ResultXML
                tbHaveCM.Rows.Add(rw)
                dsResult.Tables.Add(tbHaveCM.Copy)
                tbHaveCM.Dispose()
            End Using

            '#8475 By Kin 2019/08/19
            Dim strCD002Code As String = "-X"
            If dsResult.Tables("Wip").Rows.Count > 0 AndAlso Not DBNull.Value.Equals(dsResult.Tables("Wip").Rows(0).Item("ServCode")) Then
                strCD002Code = dsResult.Tables("Wip").Rows(0).Item("ServCode")
            End If
            Dim dtCD002 As DataTable = GetCD002(strCD002Code)
            Dim dtServiceType As DataTable = Me.GetServiceType
            dtServiceType.TableName = fCommon_ServiceType
            Dim dtWorkerEn1 As DataTable = Me.GetWorkerEn(0)
            dtWorkerEn1.TableName = fCommon_WorkerEn1
            Dim dtWorkerEn2 As DataTable = dtWorkerEn1.Copy
            dtWorkerEn2.TableName = fCommon_WorkerEn2
            Dim dtSignEn As DataTable = Me.GetSignEn
            dtSignEn.TableName = fCommon_SignEn
            Dim dtPriv As DataTable = Me.GetPriv()
            dtPriv.TableName = fCommon_Priv
            dsResult.Tables.Add(dtPriv.Copy)
            dsResult.Tables.Add(dtSignEn.Copy)
            dsResult.Tables.Add(dtWorkerEn1.Copy)
            dsResult.Tables.Add(dtWorkerEn2.Copy)
            dsResult.Tables.Add(dtServiceType.Copy)


            Dim dtSO001 As DataTable = Me.GetSO001(CustId)
            dtSO001.TableName = fCommon_SO001
            Dim dtCustomer As DataTable = Me.GetCustomer(CustId, ServiceType)
            dtCustomer.TableName = fCommon_SO002

            Dim dtMaintainCode As DataTable = Me.GetMaitainCode(ServiceType)
            dtMaintainCode.TableName = fCommon_MaintainCode
            Dim dtReturnCode As DataTable = Me.GetReturnCode(ServiceType)
            dtReturnCode.TableName = fCommon_ReturnCode
            Dim dtReturnDescCode As DataTable = Me.GetReturnDescCode(ServiceType)
            dtReturnDescCode.TableName = fCommon_ReturnDescCode
            Dim dtSatiCode As DataTable = Me.GetSatiCode(ServiceType)
            dtSatiCode.TableName = fCommon_SatiCode
            Dim dtMFCode As DataTable = Me.GetMFCode1(ServiceType)
            dtMFCode.TableName = fCommon_MFCode
            If String.IsNullOrEmpty(ServiceCode) Then
                ServiceCode = "X"
                If dtSO001.Rows.Count > 0 Then
                    ServiceCode = dtSO001.Rows(0).Item("ServCode")
                End If
            End If
            Dim dtGroupCode As DataTable = Me.GetGroupCode(ServiceCode)
            dtGroupCode.TableName = fCommon_GroupCode
            If String.IsNullOrEmpty(MFCode) Then
                MFCode = "-1"
            End If
            Dim dtMFCode2 As DataTable = Me.GetMFCode2(Integer.Parse(MFCode), ServiceType)
            dtMFCode2.TableName = fCommon_MFCode2
            Using bll As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                Using dtFieldPriv As DataTable = bll.GetFieldPrivMappingData("SO1112A", IIf(ServiceType = "X", EditMode.Append, EditMode.Edit))
                    dtFieldPriv.TableName = "FieldPriv"
                    dsResult.Tables.Add(dtFieldPriv.Copy)
                    dtFieldPriv.Dispose()
                End Using
             
                bll.Dispose()
            End Using
            dsResult.Tables.Add(dtWipName)
            dsResult.Tables.Add(dtCustomer.Copy)
            dsResult.Tables.Add(dtFalseSNO.Copy)
            dsResult.Tables.Add(dtMaintainCode.Copy)
            dsResult.Tables.Add(dtMFCode.Copy)
            dsResult.Tables.Add(dtReturnCode.Copy)
            dsResult.Tables.Add(dtReturnDescCode.Copy)
            dsResult.Tables.Add(dtSatiCode.Copy)
            dsResult.Tables.Add(dtSO001.Copy)
            dsResult.Tables.Add(dtGroupCode.Copy)
            dsResult.Tables.Add(dtMFCode2.Copy)
            dsResult.Tables.Add(dtCD002.Copy)

            dtCustomer.Dispose()
            dtMaintainCode.Dispose()
            dtMFCode.Dispose()
            dtReturnCode.Dispose()
            dtReturnDescCode.Dispose()
            dtSO001.Dispose()
            dtSatiCode.Dispose()
            dtCD002.Dispose()
            dsResult.AcceptChanges()
        Catch ex As Exception
            Throw
        Finally
            Me.DAO.AutoCloseConn = True
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If

        End Try
        Return dsResult
    End Function
    Public Function GetCD002(ByVal codeNo As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetCD002, New Object() {codeNo, LoginInfo.CompCode})
            dt.TableName = "CD002"
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function GetNullServiceType(ByVal SNo As String, ByVal CustId As Integer) As DataSet
        Dim dsResult As New DataSet
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim dtWipName As New DataTable("WipTableName")
        dtWipName.Columns.Add("WipName", GetType(String))
        Try
            Me.DAO.AutoCloseConn = False
            Dim dtServiceType As DataTable = Me.GetServiceType
            dtServiceType.TableName = fCommon_ServiceType
            Dim dtWorkerEn1 As DataTable = Me.GetWorkerEn(0)
            dtWorkerEn1.TableName = fCommon_WorkerEn1
            Dim dtWorkerEn2 As DataTable = dtWorkerEn1.Copy
            dtWorkerEn2.TableName = fCommon_WorkerEn2
            Dim dtSignEn As DataTable = Me.GetSignEn
            dtSignEn.TableName = fCommon_SignEn
            Dim dtPriv As DataTable = Me.GetPriv()
            dtPriv.TableName = fCommon_Priv
            dsResult.Tables.Add(dtPriv.Copy)
            dsResult.Tables.Add(dtSignEn.Copy)
            dsResult.Tables.Add(dtWorkerEn1.Copy)
            dsResult.Tables.Add(dtWorkerEn2.Copy)
            dsResult.Tables.Add(dtServiceType.Copy)

            'dsResult.AcceptChanges()
            Dim dtSO001 As DataTable = Me.GetSO001(CustId)
            dtSO001.TableName = fCommon_SO001
            Dim dtCustomer As DataTable = Me.GetCustomer(CustId, "")
            dtCustomer.TableName = fCommon_SO002
            Dim dtFalseSNO As New DataTable(fCommon_FalseSNo)
            dtFalseSNO.Columns.Add("SNO", GetType(String))
            Dim dtMaintainCode As DataTable = Me.GetMaitainCode("")
            dtMaintainCode.TableName = fCommon_MaintainCode
            Dim dtReturnCode As DataTable = Me.GetReturnCode("")
            dtReturnCode.TableName = fCommon_ReturnCode
            Dim dtReturnDescCode As DataTable = Me.GetReturnDescCode("")
            dtReturnDescCode.TableName = fCommon_ReturnDescCode
            Dim dtSatiCode As DataTable = Me.GetSatiCode("")
            dtSatiCode.TableName = fCommon_SatiCode
            Dim dtMFCode As DataTable = Me.GetMFCode1("")
            dtMFCode.TableName = fCommon_MFCode
            If True Then
                Dim rwNew As DataRow = dtFalseSNO.NewRow
                rwNew.Item("SNO") = Me.GetFalseSNO("").ResultXML

                SNo = rwNew.Item("SNO")
                dtFalseSNO.Rows.Add(rwNew)
                dtFalseSNO.AcceptChanges()
            End If
            Using dsWip As DataSet = obj.GetWipDetail(SNo, False, BLL.Utility.InvoiceType.Maintain)
                For Each tbWip As DataTable In dsWip.Tables
                    Dim rw As DataRow = dtWipName.NewRow
                    rw.Item("WipName") = tbWip.TableName
                    dtWipName.Rows.Add(rw)
                    dsResult.Tables.Add(tbWip.Copy)
                Next
            End Using
            Using bll As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                Using dtFieldPriv As DataTable = bll.GetFieldPrivMappingData("SO1112A", EditMode.Append)
                    dtFieldPriv.TableName = "FieldPriv"
                    dsResult.Tables.Add(dtFieldPriv.Copy)
                    dtFieldPriv.Dispose()
                End Using
                bll.Dispose()
            End Using
            dsResult.Tables.Add(dtWipName)
            dsResult.Tables.Add(dtCustomer.Copy)
            dsResult.Tables.Add(dtFalseSNO.Copy)
            dsResult.Tables.Add(dtMaintainCode.Copy)
            dsResult.Tables.Add(dtMFCode.Copy)
            dsResult.Tables.Add(dtReturnCode.Copy)
            dsResult.Tables.Add(dtReturnDescCode.Copy)
            dsResult.Tables.Add(dtSatiCode.Copy)
            dsResult.Tables.Add(dtSO001.Copy)

            dtCustomer.Dispose()

            dtMaintainCode.Dispose()
            dtMFCode.Dispose()
            dtReturnCode.Dispose()
            dtReturnDescCode.Dispose()
            dtSO001.Dispose()
            dtSatiCode.Dispose()
            dsResult.AcceptChanges()
        Catch ex As Exception
            Throw
        Finally
            Me.DAO.AutoCloseConn = True
            obj.Dispose()
        End Try
        Return dsResult
    End Function
    Public Function GetCommonData() As DataSet

        Dim dsResult As New DataSet
        Try
            Me.DAO.AutoCloseConn = False
            Dim dtServiceType As DataTable = Me.GetServiceType
            dtServiceType.TableName = fCommon_ServiceType
            Dim dtWorkerEn1 As DataTable = Me.GetWorkerEn(0)
            dtWorkerEn1.TableName = fCommon_WorkerEn1
            Dim dtWorkerEn2 As DataTable = dtWorkerEn1.Copy
            dtWorkerEn2.TableName = fCommon_WorkerEn2
            Dim dtSignEn As DataTable = Me.GetSignEn
            dtSignEn.TableName = fCommon_SignEn
            Dim dtPriv As DataTable = Me.GetPriv()
            dtPriv.TableName = fCommon_Priv
            dsResult.Tables.Add(dtPriv.Copy)
            dsResult.Tables.Add(dtSignEn.Copy)
            dsResult.Tables.Add(dtWorkerEn1.Copy)
            dsResult.Tables.Add(dtWorkerEn2.Copy)
            dsResult.Tables.Add(dtServiceType.Copy)
            dsResult.AcceptChanges()
        Catch ex As Exception
            Throw
        Finally
            Me.DAO.AutoCloseConn = True
        End Try
        Return dsResult

    End Function

    Public Function GetReInstAddrNo(ByVal CustId As Int32, ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetReInstAddrNo, New Object() {CustId, ServiceType})
    End Function
    Public Function GetSO014(ByVal AddrNo As Int32) As DataTable
        Return DAO.ExecQry(_DAL.GetSO014, New Object() {AddrNo, Me.LoginInfo.CompCode})
    End Function
    Public Function GetInvoiceNo(ByVal ServiceType As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRiaresult As New RIAResult()
        aRiaresult.ResultBoolean = True
        Try
            aRiaresult.ResultXML = obj.GetInvoiceNo(BLL.Utility.InvoiceType.Maintain, ServiceType)
        Finally
            obj.Dispose()
        End Try
        Return aRiaresult
    End Function
    Public Function GetFalseSNO(ByVal InvoiceType As BLL.Utility.InvoiceType,
                                ByVal ServiceType As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRiaresult As New RIAResult()
        aRiaresult.ResultBoolean = True
        Try
            aRiaresult.ResultXML = obj.GetFalseSNo(BLL.Utility.InvoiceType.Maintain, ServiceType)
        Finally
            obj.Dispose()
        End Try
        Return aRiaresult
    End Function
    ''' <summary>
    ''' 取得維修資料
    ''' </summary>
    ''' <param name="SNo">工單單號</param>
    ''' <returns>DataSet</returns>
    ''' <remarks></remarks>
    Public Function GetMaintainData(ByVal SNo As String) As DataSet
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return obj.GetWipDetail(SNo, False, BLL.Utility.InvoiceType.Maintain)
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得客戶基本資料(SO002)
    ''' </summary>
    ''' <param name="CustId">客編</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>TABLE</returns>
    ''' <remarks></remarks>
    Public Function GetCustomer(ByVal CustId As Int32, ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetCustomer, New Object() {CustId, ServiceType})
    End Function
    ''' <summary>
    ''' 刪除預約時間
    ''' </summary>
    ''' <returns>True Or False</returns>
    ''' <remarks>使用者沒有儲存就離開要呼叫這個動作</remarks>
    Public Function DelResvPoint() As Boolean
        Dim aRet As Boolean = False
        Using obj As New CableSoft.SO.BLL.Wip.Utility.SaveData(Me.LoginInfo, DAO)
            aRet = obj.DelResvPoint

            'Try

            'Finally
            '    obj.Dispose()
            'End Try

        End Using

        Return aRet
    End Function
    ''' <summary>
    ''' 取得客戶基本資料(SO001)
    ''' </summary>
    ''' <param name="CustId">客編</param>
    ''' <returns>Table</returns>
    ''' <remarks></remarks>
    Public Function GetSO001(ByVal CustId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetSO001, New Object() {CustId})
    End Function
    ''' <summary>
    ''' 取得可選維修類別
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>TABLE</returns>
    ''' <remarks></remarks>
    Public Function GetMaitainCode(ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetMaitainCode, New Object() {ServiceType})
    End Function
    
    ''' <summary>
    ''' 取得可選工程組別
    ''' </summary>
    ''' <param name="ServCode">服務區</param>
    ''' <returns>TABLE</returns>
    ''' <remarks></remarks>
    Public Function GetGroupCode(ByVal ServCode As String) As DataTable
        Dim aRet As DataTable = Nothing
        aRet = DAO.ExecQry(_DAL.GetGroupCode, New Object() {ServCode})
        If aRet.Rows.Count <= 0 Then
            aRet = DAO.ExecQry(_DAL.GetGroupCode2)
        End If
        Return aRet
    End Function
    ''' <summary>
    ''' 取得可選工作人員
    ''' </summary>
    ''' <param name="Type">工程人員種類</param>
    ''' <returns>TABLE</returns>
    ''' <remarks>TYPE = ( 0: 工程人員 , 1: 工程人員2 )</remarks>
    Public Function GetWorkerEn(ByVal Type As Int32) As DataTable
        Return DAO.ExecQry(_DAL.GetWorkerEn)
    End Function
    ''' <summary>
    ''' 取得可選退單原因
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>TABLE</returns>
    ''' <remarks></remarks>
    Public Function GetReturnCode(ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetReturnCode, New Object() {ServiceType})
    End Function
    ''' <summary>
    ''' 取得可選退單原因分類
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Table</returns>
    ''' <remarks></remarks>
    Public Function GetReturnDescCode(ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetReturnDescCode(ServiceType))
    End Function
    ''' <summary>
    ''' 取得可選服務滿意度
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Table</returns>
    ''' <remarks></remarks>
    Public Function GetSatiCode(ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetSatiCode, New Object() {ServiceType})
    End Function
    ''' <summary>
    ''' 取得可選簽收人員
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSignEn() As DataTable
        Return DAO.ExecQry(_DAL.GetSignEn)
    End Function
    ''' <summary>
    ''' 取得可選故障代號1
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetMFCode1(ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetMFCode1, New Object() {ServiceType})
    End Function
    ''' <summary>
    '''取得可選故障代號2
    ''' </summary>
    ''' <param name="MFCode">故障代號1</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Table</returns>
    ''' <remarks></remarks>
    Public Function GetMFCode2(ByVal MFCode As Int32, ByVal ServiceType As String) As DataTable
        Dim aRet As DataTable = Nothing
        aRet = DAO.ExecQry(_DAL.GetMFCode2(0), New Object() {ServiceType, MFCode})
        If aRet.Rows.Count <= 0 Then
            aRet = DAO.ExecQry(_DAL.GetMFCode2(1), New Object() {ServiceType})
        End If
        Return aRet
    End Function
    ''' <summary>
    ''' 是否可以指定該服務別
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CanServiceType(ByVal CustId As Int32, ByVal ServiceType As String) As RIAResult
        Dim aRet As New RIAResult()
        aRet.ResultBoolean = True
        Using rd As DbDataReader = DAO.ExecDtRdr(_DAL.GetCustomer(), New Object() {CustId, ServiceType})
            While rd.Read
                Select Case Int32.Parse("0" & rd.Item("CustStatusCode") & "")
                    Case 1
                    Case 4
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -1
                        aRet.ErrorMessage = Language.CancelCust
                    Case Else
                        Using rd2 As DbDataReader = DAO.ExecDtRdr(_DAL.GetSO042, New Object() {ServiceType})
                            While rd2.Read
                                If (rd2.IsDBNull("AbNormalMtain")) OrElse
                                    (Int32.Parse(rd2.Item("AbNormalMtain").ToString) = 0) Then
                                    If (Int32.Parse("0" & rd.Item("WipCode1")) <> 1) AndAlso
                                        (Int32.Parse("0" & rd.Item("WipCode1")) <> 2) Then
                                        aRet.ResultBoolean = False
                                        aRet.ErrorCode = -2
                                        aRet.ErrorMessage = Language.NotInstallCust
                                    End If
                                End If
                            End While
                        End Using
                End Select
            End While
        End Using
        Return aRet
    End Function
    Public Function IsFixingArea(ByVal CustId As Integer) As RIAResult
        'Dim aSQL = "Select MduId,NodeNo,CircuitNo,substr(AddrSort,0,86) AddrSort,Noe1,Noe2,Noe3,Noe4  " & _
        '                        " From SO014 Where AddrNo in (Select InstAddrNo From SO001 Where CustId = " & CustId & ") "
        Dim aMduId As String = "'X'"
        Dim aNodeNo As String = "'X'"
        Dim aCircuitNo As String = "'X'"
        Dim aAddrSort As String = "'0'"
        Dim aNoe1 As String = "-9"
        Dim aNoe2 As String = "-9"
        Dim aNoe3 As String = "-9"
        Dim aNoe4 As String = "-9"



        Using tb As DataTable = DAO.ExecQry(_DAL.GetIsFixingAreaData, New Object() {CustId})
            If tb.Rows.Count = 0 Then
                Return New RIAResult With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
                Exit Function
            End If
            For Each rw As DataRow In tb.Rows
                If Not DBNull.Value.Equals(rw("MduId")) Then
                    aMduId = aMduId & ",'" & rw("MduId") & "'"
                End If
                If Not DBNull.Value.Equals(rw("NodeNo")) Then
                    aNodeNo = aNodeNo & ",'" & rw("NodeNo") & "'"
                End If
                If Not DBNull.Value.Equals(rw("CircuitNo")) Then
                    aCircuitNo = aCircuitNo & ",'" & rw("CircuitNo") & "'"
                End If
                If Not DBNull.Value.Equals(rw("AddrSort")) Then
                    aAddrSort = rw("AddrSort")
                End If
                If Not DBNull.Value.Equals(rw("Noe1")) Then
                    aNoe1 = rw("Noe1")
                End If
                If Not DBNull.Value.Equals(rw("Noe2")) Then
                    aNoe2 = rw("Noe2")
                End If
                If Not DBNull.Value.Equals(rw("Noe3")) Then
                    aNoe3 = rw("Noe3")
                End If
                If Not DBNull.Value.Equals(rw("Noe4")) Then
                    aNoe4 = rw("Noe4")
                End If
            Next
            Using tbIsFixing As DataTable = DAO.ExecQry(_DAL.IsFixingArea(aMduId, aNodeNo, _
                                                                          aCircuitNo, aAddrSort, aNoe1, aNoe2, aNoe3, aNoe4))
                If tbIsFixing.Rows.Count = 0 Then
                    Return New RIAResult With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}

                Else
                    Return New RIAResult With {.ErrorCode = -99, .ErrorMessage = Language.FixingArea, .ResultBoolean = False}
                End If

                tbIsFixing.Dispose()
            End Using
            tb.Dispose()
        End Using

    End Function
    ''' <summary>
    ''' 可新增
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanAppend(ByVal CustId As Int32, ByVal ServiceType As String, ByVal IsContact As Boolean) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = New RIAResult With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11121")
            If Not String.IsNullOrEmpty(ServiceType) Then
                aRet = ChkCanAddSO008(CustId, ServiceType)
            Else
                Using tbServiceType As DataTable = GetServiceType()
                    For Each rw As DataRow In tbServiceType.Rows
                        aRet = ChkCanAddSO008(CustId, rw.Item("CodeNo"))
                        If aRet.ResultBoolean = True Then
                            If Not IsContact Then
                                Return aRet
                            Else
                                Return IsFixingArea(CustId)
                            End If
                        End If
                    Next
                End Using
                aRet.ResultBoolean = False
                aRet.ErrorMessage = Language.NotUseServiceType
                aRet.ErrorCode = -9
            End If
            'If Not aRet.ResultBoolean Then
            '    Return aRet
            'Else
            'Using rd As DbDataReader = DAO.ExecDtRdr(_DAL.GetCustomer(), New Object() {CustId, ServiceType})
            '    While rd.Read
            '        Select Case Int32.Parse("0" & rd.Item("CustStatusCode") & "")
            '            Case 1
            '            Case 4
            '                aRet.ResultBoolean = False
            '                aRet.ErrorCode = -1
            '                aRet.ErrorMessage = "註銷戶無法產生派工單！"
            '            Case Else
            '                Using rd2 As DbDataReader = DAO.ExecDtRdr(_DAL.GetSO042, New Object() {ServiceType})
            '                    While rd2.Read
            '                        If (rd2.IsDBNull("AbNormalMtain")) OrElse
            '                            (Int32.Parse(rd2.Item("AbNormalMtain").ToString) = 0) Then
            '                            If (Int32.Parse("0" & rd.Item("WipCode1")) <> 1) AndAlso
            '                                (Int32.Parse("0" & rd.Item("WipCode1")) <> 2) Then
            '                                aRet.ResultBoolean = False
            '                                aRet.ErrorCode = -2
            '                                aRet.ErrorMessage = "非裝機中或復機中客戶無法產生派工單 ！"
            '                            End If
            '                        End If
            '                    End While
            '                End Using
            '        End Select
            '    End While
            'End Using
            'End If
        Finally
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If


        End Try
        Return aRet

    End Function
    Public Function GetSysDate() As String
        Return DAO.ExecSclr(_DAL.GetSysDate)
    End Function
    Private Function ChkCanAddSO008(ByVal CustId As Int32, ByVal ServiceType As String) As RIAResult
        Dim aRet As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
        Dim tbSO042 As DataTable = DAO.ExecQry(_DAL.GetSO042, New Object() {ServiceType})
        Try
            Using tbCustomer As DataTable = DAO.ExecQry(_DAL.GetCustomer(), New Object() {CustId, ServiceType})
                For Each rw As DataRow In tbCustomer.Rows
                    Select Case Int32.Parse("0" & rw.Item("CustStatusCode") & "")
                        Case 1
                        Case 4
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = Language.CancelCust
                        Case 6
                            If Integer.Parse("0" & tbSO042.Rows(0).Item("PRMtain").ToString) = 1 Then
                                If Integer.Parse("0" & rw.Item("WipCode2")) = 21 Then
                                    aRet.ResultBoolean = False
                                    aRet.ErrorCode = -99
                                    aRet.ErrorMessage = Language.Mantaining
                                Else
                                    aRet.ResultBoolean = False
                                    aRet.ErrorCode = -99
                                    aRet.ErrorMessage = Language.Demolishing
                                End If
                            End If
                        Case Else

                            For Each rw042 As DataRow In tbSO042.Rows
                                If (DBNull.Value.Equals(rw042.Item("AbNormalMtain"))) OrElse
                                        (Int32.Parse(rw042.Item("AbNormalMtain").ToString) = 0) Then
                                    If (Int32.Parse("0" & rw.Item("WipCode1")) <> 1) AndAlso
                                            (Int32.Parse("0" & rw.Item("WipCode1")) <> 2) Then
                                        aRet.ResultBoolean = False
                                        aRet.ErrorCode = -2
                                        aRet.ErrorMessage = Language.NotInstallCust
                                    End If
                                    If (Integer.Parse("0" & rw.Item("Wipcode1")) = 1) OrElse
                                        (Integer.Parse("0" & rw.Item("WipCode1")) = 2) OrElse
                                        (Integer.Parse("0" & rw.Item("WipCode2")) = 21) Then
                                        aRet.ResultBoolean = False
                                        aRet.ErrorCode = -99
                                        aRet.ErrorMessage = Language.OnMaintain
                                    End If
                                Else
                                    aRet.ResultBoolean = False
                                    aRet.ErrorCode = -99
                                    aRet.ErrorMessage = Language.NotNormal
                                End If
                            Next

                    End Select
                Next
            End Using
        Catch ex As Exception
            aRet.ResultBoolean = False
            aRet.ErrorCode = -1
            aRet.ErrorMessage = ex.ToString
        Finally
            If tbSO042 IsNot Nothing Then
                tbSO042.Dispose()
                tbSO042 = Nothing
            End If

        End Try

        Return aRet
    End Function

    Public Function GetServiceType() As DataTable
        Dim dsServiceType As New DataSet
        Dim dtServiceType As New DataTable
        dtServiceType.Columns.Add("CodeNo", GetType(String))
        dtServiceType.Columns.Add("Description", GetType(String))
        dtServiceType.Columns.Add("ModifyDateChange", GetType(Int32))
        dtServiceType.Columns.Add("MoreDay2", GetType(Int32))

        Using tb As DataTable = DAO.ExecQry(_DAL.GetCD046)
            For Each rw As DataRow In tb.Rows
                Dim rwAdd As DataRow = dtServiceType.NewRow
                rwAdd.Item("CodeNo") = rw.Item("CodeNo")
                rwAdd.Item("Description") = rw.Item("Description")
                Using tbSO042 As DataTable = DAO.ExecQry(_DAL.GetSO042Para, New Object() {rw.Item("CodeNo")})
                    For Each rwSO042 As DataRow In tbSO042.Rows
                        rwAdd.Item("ModifyDateChange") = Int32.Parse(rwSO042.Item("ModifyDateChange"))
                        rwAdd.Item("MoreDay2") = Int32.Parse(rwSO042.Item("MoreDay2"))
                    Next
                End Using
                dtServiceType.Rows.Add(rwAdd)
            Next
        End Using
        dsServiceType.Tables.Add(dtServiceType)
        Return dtServiceType
    End Function
    ''' <summary>
    ''' 可修改
    ''' </summary>
    ''' <param name=" SNo">工單單號</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanEdit(ByVal SNo As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Dim aRet As RIAResult = Nothing
        Try
            Dim rData As DataTable = obj.GetPriv(LoginInfo.EntryId, "SO11122")
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11122")
            If Not aRet.ResultBoolean Then
                Return aRet
            End If
            Dim objValue As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.QueryCanEdit,
                         New Object() {SNo}))
            If objValue = 0 Then
                aRet.ResultBoolean = False
                aRet.ErrorCode = -1
                aRet.ErrorMessage = Language.CloseCanNotEdit
                Return aRet
            End If
            Dim ChkManager As DataRow = rData.AsEnumerable.Where(Function(list) list.Item("Mid") = "SO111221").FirstOrDefault()
            If ChkManager Is Nothing OrElse CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(ChkManager.Item("GroupX")) = 0 Then
                aRet.ErrorCode = 1
            End If

            'If Not DAO.ExecSclr(String.Format("Select ClsTime FROM SO008 WHERE SNO={0}0", _DAL.Sign),
            '             New Object() {SNo}) Is Nothing Then

            'End If
            'If Maintain.Rows.Count <= 0 Then
            '    aRet.ResultBoolean = False
            '    aRet.ErrorCode = -1
            '    aRet.ErrorMessage = "無任何維修單資料！"
            'Else
            '    If Not Maintain.Columns.Contains("ClsTime") Then
            '        aRet.ResultBoolean = False
            '        aRet.ErrorCode = -1
            '        aRet.ErrorMessage = "無日結欄位可判斷！"
            '    Else
            '        If Not Maintain.Rows(0).IsNull("ClsTime") Then
            '            aRet.ResultBoolean = False
            '            aRet.ErrorCode = -1
            '            aRet.ErrorMessage = "已日結不可修改資料！"
            '        End If
            '    End If
            'End If
        Finally
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If

        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 作廢維修單
    ''' </summary>
    ''' <param name="SNo">維修單號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function VoidData(ByVal SNo As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.SaveData(Me.LoginInfo, Me.DAO)
        Dim aRet As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = False}
        Try
            aRet.ResultBoolean = obj.VoidData(BLL.Utility.InvoiceType.Maintain, SNo)
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 可作廢
    ''' </summary>
    ''' <param name=" SNo">維修單號</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanDelete(ByVal SNo As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        'Dim obj As New CableSoft.SO.BLL.Wip.Utility.SaveData(Me.LoginInfo, Me.DAO)
        Dim aRet As RIAResult

        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11123")
            If Not aRet.ResultBoolean Then
                Return aRet
            End If

            Dim objValue As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.QueryCanDelete,
                         New Object() {SNo}))
            If objValue = 0 Then
                aRet.ResultBoolean = False
                aRet.ErrorCode = -1
                aRet.ErrorMessage = Language.CloseCanNotCancel
            End If
            'Dim obj1 As New CableSoft.SO.BLL.Wip.Utility.SaveData
            'obj1.VoidData(BLL.Utility.InvoiceType.Maintain, SNo)
            'CableSoft.SO.BLL.Wip.Utility.SaveData.VoidData()
            'Dim objValue As Object = DAO.ExecSclr(String.Format("Select ClsTime FROM SO008 WHERE SNO={0}0", _DAL.Sign),
            '            New Object() {SNo})

            'If Not DAO.ExecSclr(String.Format("Select ClsTime FROM SO008 WHERE SNO={0}0", _DAL.Sign),
            '             New Object() {SNo}) Is Nothing Then
            '    aRet.ResultBoolean = False
            '    aRet.ErrorCode = -1
            '    aRet.ErrorMessage = "已日結不可作廢資料！"
            'End If


        Finally
            If obj IsNot Nothing Then
                obj.Dispose()
            End If

        End Try
        Return aRet
    End Function

    Public Function CanView() As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO11124")

        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO11124")
    End Function
    ''' <summary>
    ''' 可列印
    ''' </summary>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanPrint() As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO11125")

        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO11125")
    End Function
    ''' <summary>
    ''' 取得所有權限
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetPriv() As DataTable
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
        Dim dt As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1112B")
        Dim dtSO1112A As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1112A")
        Dim dtSO1144 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1144D")
        Dim dtSO1112B As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1112B")
        Dim dtSO1100101 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1100101")
        Dim dtSO1100102 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1100102")
        Dim dtSO1132C As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1132C")
        Dim dtSO1112H As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1112H")
        Dim dtSO11127 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO11127")
        Dim dtSO11126 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO11126")
        Dim dtSO11128 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO11128")
        'Dim dtLike As DataTable = DAO.ExecQry("Select Mid From SO029 Where Mid Like 'SO1112%'")

        Try

            Try

                For Each dr As DataRow In dtSO1112A.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO11128.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO11126.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next

                For Each dr As DataRow In dtSO11127.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO1112H.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO1112B.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO1144.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO1100101.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO1100102.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO1132C.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
            Finally
                'dtSO1144.Dispose()
            End Try

            Return dt.Copy
        Finally
            If dt IsNot Nothing Then
                dt.Dispose()
                dt = Nothing
            End If
            If dtSO11128 IsNot Nothing Then
                dtSO11128.Dispose()
                dtSO11128 = Nothing
            End If
            If dtSO11126 IsNot Nothing Then
                dtSO11126.Dispose()
                dtSO11126 = Nothing
            End If
           
            If dtSO11127 IsNot Nothing Then
                dtSO11127.Dispose()
                dtSO11127 = Nothing
            End If
            If dtSO1112H IsNot Nothing Then
                dtSO1112H.Dispose()
                dtSO1112H = Nothing
            End If
            If dtSO1112B IsNot Nothing Then
                dtSO1112B.Dispose()
                dtSO1112B = Nothing
            End If
            If dtSO1132C IsNot Nothing Then
                dtSO1132C.Dispose()
                dtSO1132C = Nothing
            End If
            If dtSO1100101 IsNot Nothing Then
                dtSO1100101.Dispose()
                dtSO1100101 = Nothing
            End If
            If dtSO1100102 IsNot Nothing Then
                dtSO1100102.Dispose()
                dtSO1100102 = Nothing
            End If
            If dtSO1144 IsNot Nothing Then
                dtSO1144.Dispose()
                dtSO1144 = Nothing
            End If
            If dtSO1112A IsNot Nothing Then
                dtSO1112A.Dispose()
                dtSO1112A = Nothing
            End If
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If

        End Try

    End Function
    ''' <summary>
    ''' 檢核結清資料是否能修改
    ''' </summary>
    ''' <returns>RiaResult</returns>
    ''' <remarks></remarks>
    Public Function GetSO1132Priv() As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Try
            Return obj.ChkPriv(Me.LoginInfo.EntryId, "SO1132C")

        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO1132C")
    End Function
    ''' <summary>
    ''' 取得一般工單資訊
    ''' </summary>
    ''' <param name="InstCode">派工類別</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>Facility,Charge,ChangeFacility</remarks>
    Public Function GetNormalWip(ByVal CustId As Int32,
                                 ByVal ServiceType As String,
                                 ByVal ResvTime As Date,
                                 ByVal InstCode As Integer,
                                 ByVal dtContact As DataTable,
                                 ByVal dsWipData As DataSet, ByVal isRefresh As Boolean
                                 ) As DataSet

        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Dim retDs As DataSet = obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, CustId, ServiceType, Nothing, ResvTime, InstCode, dtContact, dsWipData)
            If isRefresh Then Return retDs
            'ChangeFacility
            If dsWipData.Tables("ChangeFacility").Rows.Count > 0 Then

                For Each rw As DataRow In dsWipData.Tables("ChangeFacility").Rows
                    If retDs.Tables("ChangeFacility").AsEnumerable.Count(Function(rwDouble As DataRow)
                                                                             Dim i As Integer = 0
                                                                             For Each val As Object In rwDouble.ItemArray
                                                                                 If val.ToString() <> rw.Item(i).ToString() Then
                                                                                     Return False
                                                                                 End If
                                                                                 i += 1
                                                                             Next
                                                                             Return True
                                                                         End Function) = 0 Then
                        retDs.Tables("ChangeFacility").Rows.Add(rw.ItemArray)

                    End If

                Next

            End If
            'Facility

            If dsWipData.Tables("Facility").Rows.Count > 0 Then
                For Each rw As DataRow In dsWipData.Tables("Facility").Rows

                    If retDs.Tables("Facility").AsEnumerable.Count(Function(rwDouble As DataRow)
                                                                       Dim i As Integer = 0
                                                                       For Each val As Object In rwDouble.ItemArray
                                                                           If val.ToString() <> rw.Item(i).ToString() Then
                                                                               Return False
                                                                           End If
                                                                           i += 1
                                                                       Next
                                                                       Return True
                                                                   End Function) = 0 Then
                        retDs.Tables("Facility").Rows.Add(rw.ItemArray)

                    End If

                Next
            End If


            'PRFacility
            If dsWipData.Tables("PRFacility").Rows.Count > 0 Then
                For Each rw As DataRow In dsWipData.Tables("PRFacility").Rows
                    If retDs.Tables("PRFacility").AsEnumerable.Count(Function(rwDouble As DataRow)
                                                                         Dim i As Integer = 0
                                                                         For Each val As Object In rwDouble.ItemArray
                                                                             If val.ToString() <> rw.Item(i).ToString() Then
                                                                                 Return False
                                                                             End If
                                                                             i += 1
                                                                         Next
                                                                         Return True
                                                                     End Function) = 0 Then
                        retDs.Tables("PRFacility").Rows.Add(rw.ItemArray)

                    End If

                Next
            End If
            Return retDs
        Finally
            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If

        End Try
    End Function

    Public Function GetMaintainChangeFaci(ByVal CustId As Integer,
                                          ByVal ServiceType As String,
                                          ByVal ResvTime As Date,
                                          ByVal WorkCodeValue As Integer,
                                          ByVal dtContact As DataTable) As DataTable
        '(1)	CableSoft.SO.BLL.Facility.ChangeFaci.GetMaintainFaci
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try

            'Dim ds As DataSet = obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, 1, "X", Date.Now, -1)

            Dim ds As DataSet = obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, CustId, ServiceType, Nothing, ResvTime, WorkCodeValue, dtContact)
            Return ds.Tables(fMaintain_ChangeFacility)
        Catch ex As Exception
            Throw New Exception(Language.GetWipCalculateDataErr & ex.ToString)
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得指定更換設備資訊
    ''' </summary>
    ''' <param name="SNo">工單單號</param>
    ''' <param name="FaciSeqNo">設備流水號</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>Facility,PRFacility,ChangeFacility</remarks>
    Public Function GetReInstChangeFaci(ByVal SNo As String, ByVal FaciSeqNo As String) As DataSet
        'CableSoft.SO.BLL.Facility.ChangeFaci.GetReInstFaci
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Dim ds As DataSet = obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, 1, "X", Date.Now, -1)
            Return ds
        Finally
            obj.Dispose()
        End Try
    End Function
    Public Function GetFalseSNO(ByVal ServiceType As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRiaresult As New RIAResult()
        aRiaresult.ResultBoolean = True
        Try
            aRiaresult.ResultXML = obj.GetFalseSNo(BLL.Utility.InvoiceType.Maintain, ServiceType)
        Finally
            obj.Dispose()
        End Try
        Return aRiaresult
    End Function
    ''' <summary>
    ''' 取得預設預約時間
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="MaintainCode">維修類別</param>
    ''' <param name="AcceptTime">受理時間</param>
    ''' <returns>Date</returns>
    ''' <remarks>受理時間可為Nothing</remarks>
    Public Function GetDefaultResvTime(ByVal ServiceType As String,
                                       ByVal MaintainCode As Int32,
                                       ByVal AcceptTime As String) As RIAResult
        Dim aRet As New RIAResult()

        aRet.ResultBoolean = True
        aRet.ResultXML = Date.Now
        DAO.AutoCloseConn = False
        Using tb As DataTable = DAO.ExecQry(_DAL.GetSO042(), New Object() {ServiceType})
            For Each rw As DataRow In tb.Rows
                Dim aResvTime As Int32 = 0
                If Not DBNull.Value.Equals(rw.Item("GetResvTime")) Then
                    aResvTime = Int32.Parse(rw.Item("GetResvTime").ToString)
                End If
                Select Case aResvTime
                    Case 1
                        Return aRet
                    Case 2
                        aRet.ResultXML = New DateTime(Date.Now.Year, Date.Now.Month,
                                                      Date.Now.Day, Date.Now.Hour, 0, 0, 0, DateTimeKind.Local).AddHours(1).ToString("yyyy/MM/dd HH:mm:ss")
                        Return aRet
                    Case 4
                        Dim aHHMM As Int32 = Int32.Parse(Left(Date.Now.Hour.ToString & Date.Now.Minute.ToString & "0000", 4))
                        Dim strTimePeriod As String = Right("0000" & DAO.ExecSclr(_DAL.GetTimePeriod(), New Object() {ServiceType, aHHMM}), 4)
                        aRet.ResultXML = New DateTime(Date.Now.Year, Date.Now.Month,
                                                      Date.Now.Day, Int32.Parse(Left(strTimePeriod, 2)),
                                                      Int32.Parse(Right(strTimePeriod, 2)), 0, 0, DateTimeKind.Local).ToString("yyyy/MM/dd HH:mm:ss")
                        Return aRet
                    Case Else
                        Dim aDate As Date = Date.Now
                        If Not String.IsNullOrEmpty(AcceptTime) Then
                            If Not Date.TryParse(AcceptTime, aDate) Then
                                aRet.ResultBoolean = False
                                aRet.ErrorCode = -1
                                aRet.ErrorMessage = Language.AcceptDateFmtErr
                                Return aRet
                            End If
                        End If
                        Dim aReserveDay As Int32 = Int32.Parse("0" & DAO.ExecSclr(_DAL.GetReserveDay, New Object() {MaintainCode}))
                        aRet.ResultXML = New DateTime(aDate.Year, aDate.Month,
                                                  aDate.Day + aReserveDay, 9,
                                                 0, 0, 0, DateTimeKind.Local).ToString("yyyy/MM/dd HH:mm:ss")
                        Return aRet
                End Select
            Next

            tb.Dispose()
        End Using
        Return aRet

        'Using rd As DbDataReader = DAO.ExecDtRdr(_DAL.GetSO042(), New Object() {ServiceType})
        '    While rd.Read
        '        Dim aResvTime As Int32 = 0
        '        If Not DBNull.Value.Equals(rd.Item("GetResvTime")) Then
        '            aResvTime = Int32.Parse(rd.Item("GetResvTime").ToString)
        '        End If
        '        Select Case aResvTime
        '            Case 1
        '                Return aRet
        '            Case 2
        '                aRet.ResultXML = New DateTime(Date.Now.Year, Date.Now.Month,
        '                                              Date.Now.Day, Date.Now.Hour, 0, 0, 0, DateTimeKind.Local).AddHours(1).ToString("yyyy/MM/dd HH:mm:ss")
        '                Return aRet
        '            Case 4
        '                Dim aHHMM As Int32 = Int32.Parse(Left(Date.Now.Hour.ToString & Date.Now.Minute.ToString & "0000", 4))
        '                Dim strTimePeriod As String = Right("0000" & DAO.ExecSclr(_DAL.GetTimePeriod(), New Object() {ServiceType, aHHMM}), 4)
        '                aRet.ResultXML = New DateTime(Date.Now.Year, Date.Now.Month,
        '                                              Date.Now.Day, Int32.Parse(Left(strTimePeriod, 2)),
        '                                              Int32.Parse(Right(strTimePeriod, 2)), 0, 0, DateTimeKind.Local).ToString("yyyy/MM/dd HH:mm:ss")
        '                Return aRet
        '            Case Else
        '                Dim aDate As Date = Date.Now
        '                If Not String.IsNullOrEmpty(AcceptTime) Then
        '                    If Not Date.TryParse(AcceptTime, aDate) Then
        '                        aRet.ResultBoolean = False
        '                        aRet.ErrorCode = -1
        '                        aRet.ErrorMessage = Language.AcceptDateFmtErr
        '                        Return aRet
        '                    End If
        '                End If
        '                Dim aReserveDay As Int32 = Int32.Parse("0" & DAO.ExecSclr(_DAL.GetReserveDay, New Object() {MaintainCode}))
        '                aRet.ResultXML = New DateTime(aDate.Year, aDate.Month,
        '                                          aDate.Day + aReserveDay, 9,
        '                                         0, 0, 0, DateTimeKind.Local).ToString("yyyy/MM/dd HH:mm:ss")
        '                Return aRet
        '        End Select
        '    End While
        'End Using
        'Return aRet
    End Function
    Public Function InsSmartCardNo(ByVal Facility As DataTable, ByVal SeqNo As String) As DataTable
        Dim RetDs As New DataSet
        Dim RetTable As DataTable = Facility.Copy
        RetDs.Tables.Add(RetTable)
        If Facility Is Nothing OrElse Facility.Rows.Count = 0 Then
            Return RetTable
        End If
        Dim lstRw As New List(Of DataRow)
        'Dim aSQL As String = String.Format("SELECT COUNT(1)  FROM CD022 " & _
        '                                     " WHERE CODENO = {0}0 AND REFNO = {0}1", _DAL.Sign)
        If String.IsNullOrEmpty(SeqNo) Then
            For Each rw As DataRow In Facility.Rows
                If (Not DBNull.Value.Equals(rw.Item("SmartCardNo"))) Then

                    If (Not DBNull.Value.Equals(rw.Item("FaciCode"))) AndAlso
                        (Int32.Parse(DAO.ExecSclr(_DAL.QueryCD022, New Object() {rw.Item("FaciCode"), 3})) > 0) Then
                        lstRw.Add(rw)
                    End If
                End If
            Next
        Else

            For Each rw As DataRow In Facility.Rows
                If rw.Item("SeqNo").ToString.ToUpper = SeqNo.ToUpper Then
                    If (Not DBNull.Value.Equals(rw.Item("SmartCardNo"))) Then
                        If (Not DBNull.Value.Equals(rw.Item("FaciCode"))) AndAlso
                            (Int32.Parse(DAO.ExecSclr(_DAL.QueryCD022, New Object() {rw.Item("FaciCode"), 3})) > 0) Then
                            lstRw.Add(rw)
                        End If
                    End If
                End If
            Next
        End If
        For Each rwChange As DataRow In lstRw
            For Each rw As DataRow In RetTable.Rows
                If rw.Item("SeqNo") <> rwChange.Item("SeqNo") Then
                    If (DBNull.Value.Equals(rw("FaciSNo"))) AndAlso
                        (Int32.Parse(DAO.ExecSclr(_DAL.QueryCD022, New Object() {rw.Item("FaciCode"), 4})) > 0) Then
                        rw.Item("FaciSNo") = rwChange.Item("SmartCardNo")
                    End If
                End If
            Next
        Next
        lstRw.Clear()
        lstRw = Nothing
        Return RetTable
    End Function
    Public Function ChooseFaciUpdData(FaciSeqNo As String, FaciSNo As String,
                                      WipRefNo As Integer, ReInstAcrossFlag As Boolean,
                                      WipData As DataSet) As Boolean
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
        Try
            Return obj.ChooseFaciUpdData(FaciSeqNo, FaciSNo, BLL.Utility.InvoiceType.Maintain, WipRefNo, ReInstAcrossFlag, WipData)
        Finally
            obj.Dispose()
        End Try
        'Return New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO).ChooseFaciUpdData(FaciSeqNo, FaciSNo, BLL.Utility.InvoiceType.Maintain, WipRefNo, ReInstAcrossFlag, WipData)
    End Function

    Public Function ChkHaveCM(ByVal WipData As DataSet, ByVal Is004D As Boolean, ByVal ServiceType As String) As RIAResult
        Dim strFaciRefNo As String = String.Empty
        Dim strFaciCode As String = String.Empty
        Dim intRefNo As Int32 = 0
        Dim aRiaresult As New RIAResult()
        aRiaresult.ResultBoolean = True
        aRiaresult.ResultXML = "0"
        If Int32.Parse(DAO.ExecSclr(_DAL.QueryMustCallOk, New Object() {ServiceType, Me.LoginInfo.CompCode}).ToString()) = 0 Then
            aRiaresult.ResultXML = "1"
            Return aRiaresult
        End If
        Select Case ServiceType
            Case "I"
                strFaciRefNo = "2,5,7,8"
            Case "D"
                strFaciRefNo = "3"
            Case "P"
                strFaciRefNo = "6"
            Case Else
                strFaciRefNo = "-99"
        End Select
        If WipData.Tables(fMaintain_Wip).Rows.Count <= 0 Then
            aRiaresult.ResultBoolean = True
            aRiaresult.ErrorCode = 0
            aRiaresult.ErrorMessage = String.Empty
            Return aRiaresult
        End If
        If (Not WipData.Tables(fMaintain_Facility).Columns.Contains("CustId")) AndAlso
            (Not WipData.Tables(fMaintain_Facility).Columns.Contains("SeqNo")) Then
            aRiaresult.ResultBoolean = False
            aRiaresult.ErrorCode = -2
            aRiaresult.ErrorMessage = Language.MustCustId
            Return aRiaresult
        End If
        For Each rw As DataRow In WipData.Tables(fMaintain_Facility).Rows
            If (Not rw.IsNull("SeqNo")) AndAlso
                (Not String.IsNullOrEmpty(rw.Item("SeqNo").ToString)) Then
                If Is004D Then
                    strFaciCode = DAO.ExecSclr(_DAL.QueryFaciCode,
                                               New Object() {rw("SeqNo")})
                Else
                    If Not rw.IsNull("FaciCode") Then
                        strFaciCode = rw("FaciCode")
                    Else
                        strFaciCode = String.Empty
                    End If
                End If
                If Not String.IsNullOrEmpty(strFaciCode) Then
                    intRefNo = Int32.Parse(DAO.ExecSclr(_DAL.QueryCD022RefNo,
                                          New Object() {strFaciCode}).ToString)
                    If strFaciRefNo.Split(",").Contains(intRefNo.ToString) Then
                        aRiaresult.ResultXML = "1"
                        Return aRiaresult
                    End If
                End If

            End If
        Next
        Return aRiaresult

    End Function
    Public Function ChkFaciFinishPrivFlag(ByVal WipData As DataSet) As DataTable
        Dim aRetDT As New DataTable(fCommon_FaciFinishPrivFlag)
        Dim intCount As Int32 = 0
        aRetDT.Columns.Add("FaciFinishPrivFlag", GetType(Boolean))
        aRetDT.Columns.Add("HaveSTB", GetType(Boolean))
        aRetDT.Columns.Add("CMFinTimeFlag", GetType(Boolean))
        Dim aRw As DataRow = aRetDT.NewRow
        aRw.BeginEdit()
        aRw("FaciFinishPrivFlag") = True
        aRw("HaveSTB") = False
        aRw.EndEdit()
        aRetDT.Rows.Add(aRw)
        If WipData.Tables(fMaintain_Facility).Rows.Count <= 0 Then
            Return aRetDT
        End If
        Dim intSTBFinTimeFlag As Int32 = Int32.Parse(DAO.ExecSclr(_DAL.QuerySTBFinTimeFlag,
                                                                    New Object() {
                                                                  Me.LoginInfo.CompCode, WipData.Tables(fMaintain_Facility).Rows(0).Item("SERVICETYPE")}).ToString)
        If intSTBFinTimeFlag = 0 Then
            aRetDT.Rows(0).Item("FaciFinishPrivFlag") = True
            aRetDT.Rows(0).Item("HaveSTB") = False
            Return aRetDT
        End If
        For Each rw As DataRow In WipData.Tables(fMaintain_Facility).Rows
            intCount = Int32.Parse(DAO.ExecSclr(_DAL.QueryCD022Count,
                                     New Object() {rw("FaciCode")}).ToString)
            If intCount > 0 Then
                aRetDT.Rows(0).Item("FaciFinishPrivFlag") = False
                aRetDT.Rows(0).Item("HaveSTB") = True
                Exit For
            End If
        Next
        If aRetDT.Rows(0).Item("FaciFinishPrivFlag") Then
            For Each rw As DataRow In WipData.Tables(fMaintain_PRFacility).Rows
                intCount = Int32.Parse(DAO.ExecSclr(_DAL.QueryCD022Count,
                                    New Object() {rw("FaciCode")}).ToString)
                If intCount > 0 Then
                    aRetDT.Rows(0).Item("FaciFinishPrivFlag") = False
                    aRetDT.Rows(0).Item("HaveSTB") = True
                    Exit For
                End If
            Next
        End If
        Return aRetDT
    End Function



#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If _DAL IsNot Nothing Then
                    _DAL.Dispose()
                    _DAL = Nothing
                End If
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                    DAO = Nothing
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
