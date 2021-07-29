Imports System.Data.Common
Imports CableSoft.BLL.Utility
Public Class BillingAPI252
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private Const fCommon_FalseSNo As String = "FalseSNo"
    Private Const fMaintain_Wip As String = "Wip"
    Private Const fMaintain_Facility As String = "Facility"
    Private Const fMaintain_PRFacility As String = "PRFacility"
    Private Const fMaintain_Charge As String = "Charge"
    Private Const fMaintain_ChangeFacility As String = "ChangeFacility"
    Private Const fMaintain_OldWip As String = "OldWip"
    Private Const fCommon_SO001 As String = "SO001"
    Private Const fCommon_SO002 As String = "SO002"
    Private Const fCommon_SO014 As String = "SO014"
    Private Const fCommon_MaintainCode As String = "MaintainCode"
    Private Const fCommon_GroupCode As String = "GroupCode"
    Private Const fCommon_Priv As String = "Priv"
    Private Const fCommon_FaciFinishPrivFlag As String = "FaciFinishPrivFlag"
    Private Const fCommon_AcceptEn As String = "AcceptEn"
    Private _DAL As New BillingAPI252DALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI252Language
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
    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim result As New RIAResult
        Dim dtSNo As New DataTable("SNO")
        Dim _Maintain As New Maintain(Me.LoginInfo, DAO)
        Dim _SaveData As New SaveData(Me.LoginInfo, DAO)
        Dim _ValiData As New Validate(Me.LoginInfo, DAO)
        Dim dsWip As DataSet = Nothing
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        Dim dsResult As New DataSet
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            If cn.State <> ConnectionState.Open Then
                cn.ConnectionString = Me.LoginInfo.ConnectionString
                cn.Open()
            End If
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        Dim dsCommon As DataSet = Nothing
        Try
            dtSNo.Columns.Add(New DataColumn("SNO", GetType(String)))
            Dim rwSNo As DataRow = dtSNo.NewRow
            rwSNo.Item("SNO") = "X"
            dtSNo.Rows.Add(rwSNo)
            dtSNo.AcceptChanges()
            Dim CustId As Integer = Integer.Parse(InData.Tables("SNO").Rows(0).Item("CustId"))
            Dim ServiceType As String = InData.Tables("SNO").Rows(0).Item("ServiceType")
            Dim ServiceCode As String = InData.Tables("SNO").Rows(0).Item("ServiceCode")
            Dim ResvTime As String = InData.Tables("SNO").Rows(0).Item("ResvTime")
            If Not IsDate(ResvTime) Then
                ResvTime = ResvTime.Replace(" ", "").ToString.Replace("/", "").ToString.Replace(":", "")
                ResvTime = Date.ParseExact(ResvTime, "yyyyMMddHHmmss", Nothing).ToString("yyyy/MM/dd HH:mm:ss")
            End If
            '取出要Ins SO008 的基本資料
            dsCommon = QueryAPI252CommData(InData)
            '取出空的Wip DataSet
            '#8731 改用_Maintain.GetNormalWip取出工單資料 By Kin 2021/05/28
            'dsWip = objWipUtility.GetWipDetail(dsCommon.Tables(fCommon_FalseSNo).Rows(0).Item("SNO"),
            '                                   False, BLL.Utility.InvoiceType.Maintain)
            With InData.Tables("SNo").Rows(0)
                dsWip = _Maintain.GetNormalWip(Integer.Parse(.Item("CustId")), .Item("ServiceType").ToString(), _
                                              Date.Parse(ResvTime), Integer.Parse(.Item("ServiceCode")), Nothing, dsWip, True)
            End With
            '測試是否可預約
            'dsCommon.Tables(fCommon_SO001).Rows(0).Item("ServCode").ToString
            result = _ValiData.ChkCanResv(InData.Tables("SNO").Rows(0).Item("WorkServCode"),
                                    ServiceCode,
                                   dsCommon.Tables(fCommon_MaintainCode).Rows(0).Item("GroupNo").ToString,
                                   ServiceType,
                                   Date.Parse(ResvTime),
                                   Date.Now,
                                    IIf(Date.Parse(ResvTime) > Date.Now, Date.Now, ResvTime),
                                   dsCommon.Tables(fCommon_MaintainCode).Rows(0).Item("ReserveDay"),
                                  dsCommon.Tables(fCommon_MaintainCode).Rows(0).Item("WorkUnit"), True)

            If Not result.ResultBoolean Then
                result.ResultBoolean = False
                result.ErrorCode = -155
                If String.IsNullOrEmpty(result.ErrorMessage) Then
                    result.ErrorMessage = Language.NotAllowReServe
                End If
                Return result
            End If
            AddTemMaintain(dsWip, dsCommon, InData)

            result = _SaveData.Save(EditMode.Append, False, dsWip, False)
            result.ResultBoolean = False
            result.ErrorCode = -1
            result.ErrorMessage = "RD Debug"
            If result.ResultBoolean Then
                'dtSNo.Rows(0).Item("AMT") = result.ResultXML
                dtSNo.Rows(0).Item("SNO") = dsWip.Tables(fMaintain_Wip).Rows(0).Item("SNO")
                result.ResultXML = Nothing
                result.ErrorCode = 0
                result.ErrorMessage = Nothing
                result.ResultBoolean = True
            
                dsResult.Tables.Add(dtSNo.Copy)
                dsResult.AcceptChanges()
                result.ResultDataSet = dsResult
                If blnAutoClose Then
                    trans.Commit()
                End If
            Else
                If blnAutoClose Then
                    trans.Rollback()
                End If
                If String.IsNullOrEmpty(result.ErrorMessage) Then
                    result.ErrorMessage = Language.SaveError
                End If
                If result.ErrorCode = 0 Then
                    result.ErrorCode = -99
                End If
                result.ResultBoolean = False
                result.ResultXML = Nothing

            End If
            Return result
        Catch ex As Exception
            result.ErrorCode = -99
            result.ResultBoolean = False
            result.ErrorMessage = result.ToString
            If blnAutoClose Then
                trans.Rollback()
            End If
            Return result
        Finally
            If _Maintain IsNot Nothing Then
                _Maintain.Dispose()
                _Maintain = Nothing
            End If
            If objWipUtility IsNot Nothing Then
                objWipUtility.Dispose()
                objWipUtility = Nothing
            End If
            If _ValiData IsNot Nothing Then
                _ValiData.Dispose()
                _ValiData = Nothing
            End If
            If dsCommon IsNot Nothing Then
                dsCommon.Dispose()
                dsCommon = Nothing
            End If
            If dtSNo IsNot Nothing Then
                dtSNo.Dispose()
                dtSNo = Nothing
            End If
            If dsWip IsNot Nothing Then
                dsWip.Dispose()
                dsWip = Nothing
            End If
            If blnAutoClose Then
                If trans IsNot Nothing Then
                    trans.Dispose()
                    trans = Nothing
                End If
                If cn IsNot Nothing Then
                    cn.Dispose()
                    cn = Nothing
                End If
            End If

        End Try

    End Function
    Private Sub AddTemMaintain(ByRef dsWip As DataSet, ByRef dsCommon As DataSet, ByRef InData As DataSet)
        Dim _ChangFaci As New CableSoft.SO.BLL.Facility.ChangeFaci.ChangeFaci(Me.LoginInfo, DAO)

        Dim dsChangeFaci As DataSet = Nothing
        Dim tbWipOld As DataTable = Nothing
        Try
            Dim rwNew As DataRow = dsWip.Tables(fMaintain_Wip).NewRow
            With rwNew
                .Item("UpdEn") = Me.LoginInfo.EntryName
                If dsCommon.Tables(fCommon_AcceptEn).Rows.Count > 0 Then
                    .Item("AcceptEn") = dsCommon.Tables(fCommon_AcceptEn).Rows(0).Item("EmpNo")
                    .Item("AcceptName") = dsCommon.Tables(fCommon_AcceptEn).Rows(0).Item("EmpName")
                End If
                .Item("AcceptTime") = DateTime.Now
                .Item("Tel1") = dsCommon.Tables(fCommon_SO001).Rows(0).Item("Tel1")
                If Not DBNull.Value.Equals(dsCommon.Tables(fCommon_SO014).Rows(0).Item("SalesCode")) Then
                    .Item("SalesCode") = dsCommon.Tables(fCommon_SO014).Rows(0).Item("SalesCode")
                    .Item("SalesName") = dsCommon.Tables(fCommon_SO014).Rows(0).Item("SalesName")
                End If
                If Not DBNull.Value.Equals(dsCommon.Tables(fCommon_SO014).Rows(0).Item("NodeNo")) Then
                    .Item("NodeNo") = dsCommon.Tables(fCommon_SO014).Rows(0).Item("NodeNo")
                End If
                .Item("AddrNo") = dsCommon.Tables(fCommon_SO014).Rows(0).Item("AddrNo")
                .Item("Address") = dsCommon.Tables(fCommon_SO014).Rows(0).Item("Address")
                If Not DBNull.Value.Equals(dsCommon.Tables(fCommon_SO014).Rows(0).Item("ServCode")) Then
                    .Item("ServCode") = dsCommon.Tables(fCommon_SO014).Rows(0).Item("ServCode")
                End If
                If Not DBNull.Value.Equals(dsCommon.Tables(fCommon_SO014).Rows(0).Item("StrtCode")) Then
                    .Item("StrtCode") = dsCommon.Tables(fCommon_SO014).Rows(0).Item("StrtCode")
                End If
                .Item("ServiceType") = InData.Tables("SNo").Rows(0).Item("ServiceType")
                .Item("WorkServCode") = dsCommon.Tables(fCommon_SO001).Rows(0)("ServCode")

                .Item("CompCode") = InData.Tables("Main").Rows(0).Item("CompCode")
                .Item("SNO") = dsCommon.Tables(fCommon_FalseSNo).Rows(0).Item("SNO")
                .Item("CustId") = dsCommon.Tables(fCommon_SO001).Rows(0).Item("CustId")
                .Item("CustName") = dsCommon.Tables(fCommon_SO001).Rows(0).Item("CustName")
                If Not DBNull.Value.Equals(InData.Tables("SNo").Rows(0).Item("Note")) Then
                    .Item("Note") = InData.Tables("SNo").Rows(0).Item("Note")
                End If
                .Item("ServiceCode") = dsCommon.Tables(fCommon_MaintainCode).Rows(0).Item("CodeNo")
                .Item("ServiceName") = dsCommon.Tables(fCommon_MaintainCode).Rows(0).Item("Description")
                If dsCommon.Tables(fCommon_GroupCode).Rows.Count > 0 Then
                    '#7332 cancel to fill the groupcode datafield by kin 2016/10/28
                    '.Item("GroupCode") = dsCommon.Tables(fCommon_GroupCode).Rows(0).Item("CodeNo")
                    '.Item("GroupName") = dsCommon.Tables(fCommon_GroupCode).Rows(0).Item("Description")
                End If

                .Item("WorkUnit") = dsCommon.Tables(fCommon_MaintainCode).Rows(0).Item("WorkUnit")
                .Item("Priority") = 0
                If Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("Priority")) Then
                    .Item("Priority") = Integer.Parse(InData.Tables("SNO").Rows(0).Item("Priority"))
                End If
                .Item("PrintBillFlag") = 0
                If Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("PrintBillFlag")) Then
                    .Item("PrintBillFlag") = Integer.Parse(InData.Tables("SNO").Rows(0).Item("PrintBillFlag"))
                End If
                Dim ResvTime As String = InData.Tables("SNO").Rows(0).Item("ResvTime")
                If Not IsDate(ResvTime) Then
                    ResvTime = ResvTime.Replace(" ", "").ToString.Replace("/", "").ToString.Replace(":", "")
                    ResvTime = Date.ParseExact(ResvTime, "yyyyMMddHHmmss", Nothing).ToString("yyyy/MM/dd HH:mm:ss")
                End If
                .Item("ResvTime") = Date.Parse(ResvTime)
                If (Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("ResvFlagTime"))) AndAlso
                    (Not String.IsNullOrEmpty(InData.Tables("SNO").Rows(0).Item("ResvFlagTime").ToString)) Then
                    .Item("ResvFlagTime") = InData.Tables("SNO").Rows(0).Item("ResvFlagTime").ToString.Replace(":", "")
                End If
                '#7337  to Add to fill out worker datafield if the datafield of data source exist   by kin 2016/11/08
                If (Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("WorkerEn1"))) AndAlso
                    (Not String.IsNullOrEmpty(InData.Tables("SNO").Rows(0).Item("WorkerEn1").ToString)) Then
                    Dim workerName1 = DAO.ExecSclr(_DAL.QueryWorkerName, New Object() {InData.Tables("SNO").Rows(0).Item("WorkerEn1")})
                    .Item("WorkerEn1") = InData.Tables("SNO").Rows(0).Item("WorkerEn1")                    
                    .Item("WorkerName1") = workerName1
                End If
                '#7899 By Kin 2018/09/11
                If (Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("WorkServCode"))) Then
                    .Item("WorkServCode") = InData.Tables("SNO").Rows(0).Item("WorkServCode")
                End If
                '#8792 該SO009服務別有未結案(signdate is null)的移機單(cd007.refno=3)，且
                '[UseReInstaddr](傳1, 則維修單的地址要帶該SO009.reinstaddrno)

                If Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("UseReInstaddr")) AndAlso _
                    Integer.Parse(InData.Tables("SNO").Rows(0).Item("UseReInstaddr")) = 1 Then
                    Using so009Addr As DataTable = DAO.ExecQry(_DAL.getSO009Reinstaddrno, _
                                                               New Object() {InData.Tables("SNo").Rows(0).Item("ServiceType"), _
                                                                             dsCommon.Tables(fCommon_SO001).Rows(0).Item("CustId")})
                      
                        If so009Addr.Rows.Count > 0 Then
                            .Item("AddrNo") = so009Addr.Rows(0).Item("ReInstAddrNo")
                            .Item("Address") = so009Addr.Rows(0).Item("ReInstAddress")
                            .Item("servcode") = so009Addr.Rows(0).Item("servcode")
                            .Item("StrtCode") = so009Addr.Rows(0).Item("StrtCode")
                            .Item("SalesCode") = so009Addr.Rows(0).Item("SalesCode")
                            .Item("SalesName") = so009Addr.Rows(0).Item("SalesName")
                        End If
                    End Using
                End If

            End With

            dsWip.Tables(fMaintain_Wip).Rows.Add(rwNew)
            tbWipOld = dsWip.Tables(fMaintain_Wip).Copy
            tbWipOld.TableName = fMaintain_OldWip
            dsWip.Tables.Add(tbWipOld)
            dsWip.Tables(fMaintain_Wip).AcceptChanges()
            If Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("CallSeqNo")) Then
                Using dtContact As DataTable = DAO.ExecQry(_DAL.QuerySO006, New Object() {InData.Tables("SNO").Rows(0).Item("CallSeqNo")})
                    dtContact.TableName = "Contact"
                    If dtContact.Rows.Count > 0 Then
                        dsWip.Tables.Add(dtContact.Copy)
                    End If
                    dtContact.Dispose()
                End Using
            End If

            '如果有設備要加入設備資訊
            '#7332 change the way that select faciseqno datafield from facisno datafield by Kin 2016/10/28
            Dim aFaciSNoName As String = "Facisno"
            For i As Integer = 0 To 1
                If i = 1 Then aFaciSNoName = "Facisno2"
                Dim aFaciSeqno As String = Nothing
                If (Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item(aFaciSNoName))) AndAlso
                        (Not String.IsNullOrEmpty(InData.Tables("SNO").Rows(0).Item(aFaciSNoName).ToString)) Then
                    aFaciSeqno = DAO.ExecSclr(_DAL.QueryFaciSeqno, New Object() {
                                                InData.Tables("SNO").Rows(0).Item(aFaciSNoName).ToString,
                                                InData.Tables("SNo").Rows(0).Item("ServiceType"),
                                                 dsCommon.Tables(fCommon_SO001).Rows(0).Item("CustId")
                                              })
                End If


                If (Not DBNull.Value.Equals(aFaciSeqno)) AndAlso
                    (Not String.IsNullOrEmpty(aFaciSeqno)) Then
                    '#7325 do not process other facilities when the soruce of the  DVR is nothing By Kin 2016/10/19
                    Dim filterDVR As Boolean = False
                    If (DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("DVR"))) OrElse
                        (String.IsNullOrEmpty(InData.Tables("SNO").Rows(0).Item("DVR"))) Then
                        filterDVR = True
                    End If
                    If Integer.Parse(DAO.ExecSclr(_DAL.QueryDVRTryMustPair, New Object() {aFaciSeqno})) > 0 Then
                        filterDVR = False
                    End If

                    '0.維修 
                    If Integer.Parse(InData.Tables("SNO").Rows(0).Item("Kind")) = 0 Then

                        dsChangeFaci = _ChangFaci.GetMaintainFaci(dsWip.Tables(fMaintain_Wip).Rows(0).Item("SNO").ToString,
                                                                        aFaciSeqno, filterDVR, dsWip, 0).DataSet


                    Else
                        '1.更換

                        dsChangeFaci = _ChangFaci.GetReInstFaci(dsWip.Tables(fMaintain_Wip).Rows(0).Item("SNO").ToString,
                                                                       aFaciSeqno, dsWip, filterDVR)
                    End If

                    For Each tb As DataTable In dsChangeFaci.Tables
                        For Each rw As DataRow In tb.Rows
                            Dim blnAdd As Boolean = True
                            If tb.Columns.Contains("FaciSNo".ToUpper) Then
                                For Each rw2 As DataRow In dsWip.Tables(tb.TableName).Rows
                                    If (rw2.Item("Facisno").Equals(rw.Item("FaciSNo"))) Then
                                        If DBNull.Value.Equals(rw.Item("Facisno")) Then
                                            blnAdd = True
                                        Else
                                            blnAdd = False
                                        End If

                                    End If
                                Next
                            End If
                            '#8795 same facisno must add
                            If blnAdd OrElse 1 = 1 Then
                                Dim rwNewFaci As DataRow = dsWip.Tables(tb.TableName).NewRow
                                For Each col As DataColumn In dsWip.Tables(tb.TableName).Columns
                                    rwNewFaci.Item(col.ColumnName) = rw.Item(col.ColumnName)
                                Next
                                dsWip.Tables(tb.TableName).Rows.Add(rwNewFaci)
                                dsWip.Tables(tb.TableName).AcceptChanges()
                            End If
                        Next
                    Next
                End If
            Next



        Catch ex As Exception
            Throw ex
        Finally
            If _ChangFaci IsNot Nothing Then
                _ChangFaci.Dispose()
                _ChangFaci = Nothing
            End If
            If dsChangeFaci IsNot Nothing Then
                dsChangeFaci.Dispose()
                dsChangeFaci = Nothing
            End If
        End Try
    End Sub

    Private Function QueryAPI252CommData(ByRef InData As DataSet) As DataSet
        Dim dsResult As New DataSet
        Dim dtAcceptEn As DataTable = Nothing
        If Not DBNull.Value.Equals(InData.Tables("SNO").Rows(0).Item("AcceptEn")) Then
            dtAcceptEn = DAO.ExecQry(_DAL.QueryAcceptEn, New Object() {InData.Tables("SNO").Rows(0).Item("AcceptEn")})
        Else
            dtAcceptEn = DAO.ExecQry(_DAL.QueryAcceptEn, New Object() {Me.LoginInfo.EntryId})
        End If
        '#8706
        If dtAcceptEn IsNot Nothing Then
            If dtAcceptEn.Rows.Count > 0 Then
                Me.LoginInfo.EntryId = dtAcceptEn.Rows(0).Item("EmpNo")
                Me.LoginInfo.EntryName = dtAcceptEn.Rows(0).Item("EmpName")
            End If
        End If
        Dim WipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim _Maintain As New Maintain(Me.LoginInfo, DAO)
        Dim dtWipName As New DataTable("WipTableName")
        dtWipName.Columns.Add("WipName", GetType(String))
        Dim dtFalseSNO As New DataTable(fCommon_FalseSNo)
        dtFalseSNO.Columns.Add("SNO", GetType(String))

        Try
            Dim SNo As String = Nothing
            Dim rwNew As DataRow = dtFalseSNO.NewRow
            rwNew.Item("SNO") = _Maintain.GetFalseSNO(InData.Tables("SNO").Rows(0).Item("ServiceType").ToString).ResultXML
            SNo = rwNew.Item("SNO")
            dtFalseSNO.Rows.Add(rwNew)
            dtFalseSNO.AcceptChanges()

            Using dsWip As DataSet = WipUtility.GetWipDetail(SNo, False, BLL.Utility.InvoiceType.Maintain)
                For Each tbWip As DataTable In dsWip.Tables
                    Dim rwWip As DataRow = dtWipName.NewRow
                    rwWip.Item("WipName") = tbWip.TableName
                    dtWipName.Rows.Add(rwWip)
                    dsResult.Tables.Add(tbWip.Copy)
                Next
                Using tbFaciFinishPrivFlag As DataTable = _Maintain.ChkFaciFinishPrivFlag(dsWip)
                    dsResult.Tables.Add(tbFaciFinishPrivFlag.Copy)
                End Using
                Dim tbHaveCM As New DataTable("HaveCM")
                tbHaveCM.Columns.Add("ResultBoolean", GetType(Boolean))
                tbHaveCM.Columns.Add("ErrorCode", GetType(Integer))
                tbHaveCM.Columns.Add("ErrorMessage", GetType(String))
                tbHaveCM.Columns.Add("ResultXML", GetType(String))
                Dim aRiaResult As RIAResult = _Maintain.ChkHaveCM(dsWip, False, "I")
                Dim rw As DataRow = tbHaveCM.NewRow
                rw.Item("ResultBoolean") = aRiaResult.ResultBoolean
                rw.Item("ErrorCode") = aRiaResult.ErrorCode
                rw.Item("ErrorMessage") = aRiaResult.ErrorMessage
                rw.Item("ResultXML") = aRiaResult.ResultXML
                tbHaveCM.Rows.Add(rw)
                dsResult.Tables.Add(tbHaveCM.Copy)
                tbHaveCM.Dispose()
            End Using

            'Dim dtWorkerEn1 As DataTable = Me.GetWorkerEn(0)
            'dtWorkerEn1.TableName = fCommon_WorkerEn1
            'Dim dtWorkerEn2 As DataTable = dtWorkerEn1.Copy
            'dtWorkerEn2.TableName = fCommon_WorkerEn2
            'Dim dtSignEn As DataTable = Me.GetSignEn
            'dtSignEn.TableName = fCommon_SignEn
            Dim dtPriv As DataTable = _Maintain.GetPriv()
            dtPriv.TableName = fCommon_Priv
            dsResult.Tables.Add(dtPriv.Copy)
            'dsResult.Tables.Add(dtSignEn.Copy)
            'dsResult.Tables.Add(dtWorkerEn1.Copy)
            'dsResult.Tables.Add(dtWorkerEn2.Copy)
            'dsResult.Tables.Add(dtServiceType.Copy)

            dtAcceptEn.TableName = fCommon_AcceptEn
            Dim dtSO001 As DataTable = _Maintain.GetSO001(Integer.Parse(InData.Tables("SNO").Rows(0).Item("CustId").ToString))
            dtSO001.TableName = fCommon_SO001
            Dim dtCustomer As DataTable = _Maintain.GetCustomer(Integer.Parse(InData.Tables("SNO").Rows(0).Item("CustId").ToString),
                                                         InData.Tables("SNO").Rows(0).Item("ServiceType").ToString)
            dtCustomer.TableName = fCommon_SO002

            Dim dtMaintainCode As DataTable = Me.QueryServiceCode(
                Integer.Parse(InData.Tables("SNO").Rows(0).Item("ServiceCode").ToString))
            dtMaintainCode.TableName = fCommon_MaintainCode
            '取出SO014
            Dim AddrNo As Int32 = Int32.Parse(dtSO001.Rows(0).Item("InstAddrNO"))
            If (Not DBNull.Value.Equals(dtCustomer.Rows(0).Item("WipCode3"))) AndAlso (Integer.Parse(dtCustomer.Rows(0).Item("WipCode3").ToString) = 11) Then
                Using dtReInstAddrNo As DataTable = _Maintain.GetReInstAddrNo(Integer.Parse(InData.Tables("SNO").Rows(0).Item("CustId").ToString),
                                                                                                                InData.Tables("SNO").Rows(0).Item("ServiceType").ToString)
                    If (dtReInstAddrNo.Rows.Count > 0) AndAlso (Not DBNull.Value.Equals(dtReInstAddrNo.Rows(0).Item("ReInstAddrNo"))) Then
                        AddrNo = dtReInstAddrNo.Rows(0).Item("ReInstAddrNo")
                    End If
                    dtReInstAddrNo.Dispose()
                End Using
            End If
            '全部都用目前地址，先不用管是否移機
            AddrNo = Int32.Parse(dtSO001.Rows(0).Item("InstAddrNO"))
            Dim dtSO014 As DataTable = _Maintain.GetSO014(AddrNo)
            dtSO014.TableName = fCommon_SO014        
            Dim dtGroupCode As DataTable = _Maintain.GetGroupCode(dtSO001.Rows(0).Item("ServCode"))
            dtGroupCode.TableName = fCommon_GroupCode

            dsResult.Tables.Add(dtWipName)
            dsResult.Tables.Add(dtCustomer.Copy)
            dsResult.Tables.Add(dtFalseSNO.Copy)
            dsResult.Tables.Add(dtMaintainCode.Copy)

            dsResult.Tables.Add(dtSO001.Copy)
            dsResult.Tables.Add(dtGroupCode.Copy)
            dsResult.Tables.Add(dtAcceptEn.Copy)
            dsResult.Tables.Add(dtSO014.Copy)
            dsResult.AcceptChanges()
            dtCustomer.Dispose()
            dtCustomer = Nothing
            dtMaintainCode.Dispose()
            dtMaintainCode = Nothing
            dtAcceptEn.Dispose()
            dtAcceptEn = Nothing
            dtSO001.Dispose()
            dtSO001 = Nothing
            dtSO014.Dispose()
            dtSO014 = Nothing

        Catch ex As Exception
            Throw
        Finally
            If WipUtility IsNot Nothing Then
                WipUtility.Dispose()
                WipUtility = Nothing
            End If
            If _Maintain IsNot Nothing Then
                _Maintain.Dispose()
                _Maintain = Nothing
            End If
        End Try
        Return dsResult
    End Function
    Private Function QueryServiceCode(ByVal ServiceCode As Integer) As DataTable        
        Using tbResult As DataTable = DAO.ExecQry(_DAL.QuertyServiceCode, New Object() {ServiceCode})
            Return tbResult.Copy
            tbResult.Dispose()
        End Using
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
                    DAO = Nothing
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
