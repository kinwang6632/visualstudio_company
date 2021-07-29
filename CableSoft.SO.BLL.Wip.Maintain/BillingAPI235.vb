Imports System.Data.Common
Imports CableSoft.BLL.Utility

Public Class BillingAPI235
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI

    Private Const fMaintain_Wip As String = "Wip"
    Private Const fMaintain_Facility As String = "Facility"
    Private Const fMaintain_PRFacility As String = "PRFacility"
    Private Const fMaintain_Charge As String = "Charge"
    Private Const fMaintain_ChangeFacility As String = "ChangeFacility"
    Private Const fMaintain_OldWip As String = "OldWip"
    Private _DAL As New BillingAPI235DALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI235Language
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
    Function Execute(SeqNo As Integer, InData As DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim result As New RIAResult
        'Me.LoginInfo.EntryName = InData.Tables("Main").Rows(0).Item("Upden")
        Dim _Maintain As New Maintain(Me.LoginInfo, DAO)
        Dim _VailData As New Validate(Me.LoginInfo, DAO)
        Dim _SaveData As New SaveData(Me.LoginInfo, DAO)
        Dim dsOriginalWip As DataSet = _Maintain.GetMaintainData(InData.Tables("Main").Rows(0).Item("SNO"))
        Dim dsWip As DataSet = Nothing
        Dim tbOldWip As DataTable = Nothing
        Dim dtCD006 As DataTable = Nothing
        Dim dtAMT As New DataTable("AMT")
        Dim dsResult As New DataSet
        Dim ResvTime As String = InData.Tables("Main").Rows(0).Item("ResvTime")        
        Dim NoteContent As String = String.Empty
        Dim NoteType As Integer = -1

        Dim Amt As String = "0"
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        '#8731 重新帶入工單資料 
        If InData.Tables("Main").Columns.Contains("HandleEn") Then
            Me.LoginInfo.EntryId = InData.Tables("Main").Rows(0).Item("HandleEn")
            Using tbCM003 As DataTable = DAO.ExecQry(_DAL.GetEmpName, New Object() {Me.LoginInfo.EntryId})
                Me.LoginInfo.EntryName = tbCM003.Rows(0).Item("EmpName")
            End Using
        End If

        With dsOriginalWip.Tables(fMaintain_Wip).Rows(0)
            Using obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
                dsWip = obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, Integer.Parse(.Item("CustId")), .Item("ServiceType").ToString(),
                                                 .Item("SNO").ToString, Date.Parse(ResvTime), Integer.Parse(.Item("ServiceCode")),
                                                 Nothing, Nothing)
            End Using
            'dsWip = _Maintain.GetNormalWip(Integer.Parse(.Item("CustId")), .Item("ServiceType").ToString(), _
            '                              Date.Parse(ResvTime), Integer.Parse(.Item("ServiceCode")), Nothing, dsWip, True)
            'Using obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
            '    Dim retDs As DataSet = obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, CustId, ServiceType, Nothing, ResvTime, InstCode, dtContact, dsWipData)
            'End Using

        End With

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
        If Not IsDate(ResvTime) Then
            ResvTime = ResvTime.Replace(" ", "").ToString.Replace("/", "").ToString.Replace(":", "")
            ResvTime = Date.ParseExact(ResvTime, "yyyyMMddHHmmss", Nothing).ToString("yyyy/MM/dd HH:mm:ss")
        End If
        dtAMT.Columns.Add(New DataColumn("Amount", GetType(String)))

        Dim rwNew As DataRow = dtAMT.NewRow
        rwNew.Item("Amount") = Amt
        dtAMT.Rows.Add(rwNew)
        dtAMT.AcceptChanges()
        '---------------------------------------------------------------------------------------------------------
        '#7354 According to notetype to decide the content of note by kin 2016/12/08

        If InData.Tables("Main").Columns.Contains("Note") Then
            If (Not DBNull.Value.Equals(InData.Tables("Main").Rows(0).Item("Note"))) Then
                NoteContent = InData.Tables("Main").Rows(0).Item("Note")
            End If
            If (Not DBNull.Value.Equals(InData.Tables("Main").Rows(0).Item("NoteType")) AndAlso
                 Not String.IsNullOrEmpty(InData.Tables("Main").Rows(0).Item("NoteType").ToString)) Then
                NoteType = Integer.Parse(InData.Tables("Main").Rows(0).Item("NoteType").ToString)
            End If
        End If
        '#7899 Add WorkServCode to maintain by kin 2018/11/22
        If InData.Tables("Main").Columns.Contains("WorkServCode") Then
            If (Not DBNull.Value.Equals(InData.Tables("Main").Rows(0).Item("WorkServCode")) AndAlso
                Not String.IsNullOrEmpty(InData.Tables("Main").Rows(0).Item("WorkServCode").ToString)) Then                
                dsWip.Tables(fMaintain_Wip).Rows(0).Item("WorkServCode") = InData.Tables("Main").Rows(0).Item("WorkServCode")
            End If
        End If
        Select Case NoteType
            Case 0
                dsWip.Tables(fMaintain_Wip).Rows(0).Item("Note") = NoteContent
            Case 1
                dsWip.Tables(fMaintain_Wip).Rows(0).Item("Note") = NoteContent & dsWip.Tables(fMaintain_Wip).Rows(0).Item("Note")
            Case 2
                dsWip.Tables(fMaintain_Wip).Rows(0).Item("Note") = dsWip.Tables(fMaintain_Wip).Rows(0).Item("Note") & NoteContent
        End Select
        '-----------------------------------------------------------------------------------------------------------
        Try
            If dsWip.Tables(fMaintain_Wip).Rows.Count = 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -101
                result.ErrorMessage = Language.NotFoundSNo
                Return result
            Else
                tbOldWip = dsWip.Tables(fMaintain_Wip).Copy
                tbOldWip.TableName = fMaintain_OldWip
                dsWip.Tables.Add(tbOldWip)
            End If
            dtCD006 = DAO.ExecQry(_DAL.GetMaitainCode,
                                  New Object() {dsWip.Tables(fMaintain_Wip).Rows(0).Item("ServiceCode").ToString})
            result = _VailData.ChkDataOk(EditMode.Edit, dsWip, False)
            If Not result.ResultBoolean Then
                Return result
            End If
            'dsWip.Tables(fMaintain_Wip).Rows(0).Item("ServCode").ToString()
            result = _VailData.ChkCanResv(InData.Tables("Main").Rows(0).Item("WorkServCode"),
                                    dsWip.Tables(fMaintain_Wip).Rows(0).Item("ServiceCode").ToString,
                                    dtCD006.Rows(0).Item("GroupNo").ToString,
                                    dsWip.Tables(fMaintain_Wip).Rows(0).Item("ServiceType").ToString,
                                    Date.Parse(ResvTime),
                                    dsWip.Tables(fMaintain_Wip).Rows(0).Item("AcceptTime"),
                                    dsWip.Tables(fMaintain_Wip).Rows(0).Item("ResvTime"),
                                    dtCD006.Rows(0).Item("ReserveDay"),
                                    dtCD006.Rows(0).Item("WorkUnit"), True)
            If result.ResultBoolean = False Then
                Return result
            End If
            dsWip.Tables(fMaintain_Wip).Rows(0).Item("ResvTime") = Date.Parse(ResvTime)
            dsWip.Tables(fMaintain_Wip).AcceptChanges()
            result = _SaveData.Save(EditMode.Edit, False, dsWip, False)
            dsWip = _Maintain.GetMaintainData(InData.Tables("Main").Rows(0).Item("SNO"))
            For Each rwCharge As DataRow In dsWip.Tables(fMaintain_Charge).Rows
                If Not DBNull.Value.Equals(rwCharge.Item("ShouldAmt")) Then
                    If (Not DBNull.Value.Equals(rwCharge.Item("CancelFlag"))) AndAlso
                        (Integer.Parse(rwCharge.Item("CancelFlag")) = 1) Then
                    Else
                        Amt = Integer.Parse(Amt) + Integer.Parse(rwCharge.Item("ShouldAmt").ToString)
                    End If
                End If
            Next
            dtAMT.Rows(0).Item("Amount") = Amt
            dsResult.Tables.Add(dtAMT)
            dsResult.AcceptChanges()
            If blnAutoClose Then
                trans.Commit()
            End If
            result.ErrorCode = 0
            result.ErrorMessage = Nothing
            result.ResultXML = Nothing
            result.ResultDataSet = dsResult
            result.ResultBoolean = True
            Return result
        Catch ex As Exception
            result.ResultBoolean = False
            result.ErrorCode = -99
            result.ErrorMessage = ex.ToString
            Return result
            ' trans.Rollback()
        Finally
            If Language IsNot Nothing Then
                Language.Dispose()
                Language = Nothing
            End If
            If _DAL IsNot Nothing Then
                _DAL.Dispose()
                _DAL = Nothing
            End If
            If _Maintain IsNot Nothing Then
                _Maintain.Dispose()
                _Maintain = Nothing
            End If
            If _SaveData IsNot Nothing Then
                _SaveData.Dispose()
                _SaveData = Nothing
            End If
            If tbOldWip IsNot Nothing Then
                tbOldWip.Dispose()
                tbOldWip = Nothing
            End If
            If dtCD006 IsNot Nothing Then
                dtCD006.Dispose()
                dtCD006 = Nothing
            End If

            If _VailData IsNot Nothing Then
                _VailData.Dispose()
                _VailData = Nothing
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
