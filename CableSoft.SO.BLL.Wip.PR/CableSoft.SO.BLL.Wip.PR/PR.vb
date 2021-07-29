Option Compare Binary
Option Infer On
Option Explicit On

Imports CableSoft.BLL.Utility
Imports SOUtilClass = CableSoft.SO.BLL.Utility

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports System.Data
Imports System.Data.Common
Imports CableSoft.Utility.DataAccess

Public Class PR
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New PRDAL(Me.LoginInfo.Provider)

    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
    ''' <summary>
    ''' 取得可選停拆機類別(GetPRCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetPRCode(ByVal ServiceType As String) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRCode, ServiceType, False)
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選停拆機類別(GetPRCodeByContactRefNo)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="ContactRefno">互動管理參考號</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetPRCodeByContactRefNo(ByVal ServiceType As String, ByVal ContactRefno As Integer) As DataTable
        Dim RefNo As String = String.Empty
        '(1)	ContactRefNo = 4:移機 則 派工參考號過濾3
        '(2)	ContactRefNo = 5:拆機 則 派工參考號過濾2,5,6,8
        '(3)	ContactRefNo = 6:停機則 派工參考號過濾1,11
        '(4)	ContactRefNo = 11:拆設備則 派工參考號過濾6,8
        '(5)	ContactRefNo = 27,28:同/跨區移機則 派工參考號過濾2,6,8
        '(6)	ContactRefNo = 30:關機 則 派工參考號過濾7
        '(7)	ContactRefNo = 32,33:關機 則 派工參考號過濾10
        '(8)	ContactRefNo = 30:關機 則 派工參考號過濾7
        '(9)	ContactRefNo = 39:暫停頻道 則 派工參考號過濾15
        '(10)	ContactRefNo = 42:暫停頻道 則 派工參考號過濾14
        Select Case ContactRefno
            Case 4
                RefNo = "3"
            Case 5
                RefNo = "2,5,6,8"
            Case 6
                RefNo = "1,11"
            Case 11
                RefNo = "6,8"
            Case 27, 28
                RefNo = "2,6,8"
            Case 30
                RefNo = "7"
            Case 32, 33
                RefNo = "10"
            Case 39
                RefNo = "15"
            Case 42
                RefNo = "14"
            Case Else
                RefNo = ""
        End Select
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRCodeByContactRefNo, New Object() {ServiceType, RefNo})
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選停拆移機原因(GetPRReasonCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetPRReasonCode(ByVal ServiceType As String) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRReasonCode, ServiceType, False)
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選停拆移機原因(GetPRReasonDescCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="PRReasonCode">停拆移機原因</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetPRReasonDescCode(ByVal ServiceType As String, ByVal PRReasonCode As Integer) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRReasonDescCode, New Object() {ServiceType, PRReasonCode})
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選工程組別(GetGroupCode)
    ''' </summary>
    ''' <param name="ServCode">服務區</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetGroupCode(ByVal ServCode As String) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetGroupCode, ServCode, False)
        If dt.Rows.Count <= 0 Then dt = DAO.ExecQry("Select * From CD003 Where StopFlag = 0")
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選工作人員(GetWorkerEn)
    ''' </summary>
    ''' <param name="Type">工程人員種類 (0:工程人員1,1:工程人員2)</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetWorkerEn(ByVal Type As Integer) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetWorkerEn)
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選退單原因(GetReturnCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetReturnCode(ByVal ServiceType As String) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetReturnCode, ServiceType, False)
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選退單原因分類(GetReturnDescCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetReturnDescCode(ByVal ServiceType As String) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetReturnDescCode, ServiceType, False)
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選簽收人員(GetSignEn)
    ''' </summary>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetSignEn() As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetSignEn)
        Return dt
    End Function
    ''' <summary>
    ''' 取得可選服務滿意度(GetSatiCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetSatiCode(ByVal ServiceType As String) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetSatiCode, ServiceType, False)
        Return dt
    End Function

    Public Function CanAppend(ByVal CustId As Integer, ByVal ServiceType As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11131")

            If Not aRet.ResultBoolean Then
                Return aRet
            Else
                Using rd As DbDataReader = DAO.ExecDtRdr(_DAL.GetCustomer, New Object() {CustId, ServiceType})
                    While rd.Read
                        Select Case Int32.Parse("0" & rd.Item("CustStatusCode") & "")
                            Case 1
                            Case 4
                                aRet.ResultBoolean = False
                                aRet.ErrorCode = -4
                                aRet.ErrorMessage = "註銷戶無法產生派工單！"
                            Case 5
                                Using rd2 As DbDataReader = DAO.ExecDtRdr(_DAL.GetSO042, New Object() {ServiceType})
                                    While rd2.Read
                                        If (rd2.IsDBNull("AbnormalFaci ")) OrElse
                                            (Int32.Parse(rd2.Item("AbnormalFaci ").ToString) = 0) Then
                                            aRet.ResultBoolean = False
                                            aRet.ErrorCode = -5
                                            aRet.ErrorMessage = "可傳派工類別: 15"
                                        Else
                                            aRet.ResultBoolean = False
                                            aRet.ErrorCode = -5
                                            aRet.ErrorMessage = "可派的派工類別: 6,8,10,15"
                                        End If
                                    End While
                                End Using
                            Case Else
                                aRet.ResultBoolean = False
                                aRet.ErrorCode = -99
                                aRet.ErrorMessage = "可派的派工參考號: 6,7,9,10,15"
                        End Select
                    End While
                End Using
            End If
        Finally
            obj.Dispose()
        End Try
        Return aRet

    End Function
    ''' <summary>
    ''' 可修改
    ''' </summary>
    ''' <param name="Maintain">維修單資料</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanEdit(ByVal Maintain As DataTable) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11132’")
            If Not aRet.ResultBoolean Then
                Return aRet
            End If
            If Maintain.Rows.Count <= 0 Then
                aRet.ResultBoolean = False
                aRet.ErrorCode = -1
                aRet.ErrorMessage = "無任何維修單資料！"
            Else
                If Not Maintain.Columns.Contains("ClsTime") Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = "無日結欄位可判斷！"
                Else
                    If Not Maintain.Rows(0).IsNull("ClsTime") Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -1
                        aRet.ErrorMessage = "已日結不可修改資料！"
                    End If
                End If
            End If
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function

    ''' <summary>
    ''' 可作廢
    ''' </summary>
    ''' <param name="Maintain">維修單資料</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanDelete(ByVal Maintain As DataTable) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11133’’")
            If Not aRet.ResultBoolean Then
                Return aRet
            End If
            If Maintain.Rows.Count <= 0 Then
                aRet.ResultBoolean = False
                aRet.ErrorCode = -1
                aRet.ErrorMessage = "無任何維修單資料！"
            Else
                If Not Maintain.Columns.Contains("ClsTime") Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = "無日結欄位可判斷！"
                Else
                    If Not Maintain.Rows(0).IsNull("ClsTime") Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -1
                        aRet.ErrorMessage = "已日結不可作廢資料！"
                    End If
                End If
            End If
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 可列印
    ''' </summary>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanPrint() As RIAResult
        Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo).ChkPriv(Me.LoginInfo.EntryId, "SO11135")
    End Function
    ''' <summary>
    ''' 取得所有權限
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetPriv() As DataTable
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo)
        Try
            Dim dt As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1113")
            Return dt
        Finally
            obj.Dispose()
        End Try
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
                                 ByVal InstCode As Int32) As DataSet
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return (obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, CustId, ServiceType, ResvTime, InstCode))
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得轉換派工類別(GetChangePRCode)
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="PRRefNo">派工類別參考號</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetChangePRCode(ByVal CustId As Integer, ByVal ServiceType As String, ByVal PRRefNo As Integer) As DataTable
        Dim dtRtn As DataTable = Nothing
        Select Case PRRefNo
            Case 2, 5, 6
                Dim Refno As String = String.Empty
                Select Case ServiceType.ToUpper
                    Case "C"
                        Refno = "10"
                    Case "D"
                        Refno = "3"
                    Case "I"
                        Refno = "2,5,7,8"
                    Case "P"
                        Refno = "6"
                End Select
                Dim Count004 As Int16 = DAO.ExecNqry(_DAL.GetChangePRCode, New Object() {CustId, Refno})
                If Count004 = 0 Then
                    dtRtn = DAO.ExecQry(_DAL.GetChangePRCode2, New Object() {"2,5", ServiceType})
                Else
                    dtRtn = DAO.ExecQry(_DAL.GetChangePRCode2, New Object() {"6", ServiceType})
                End If
        End Select
        Return dtRtn
    End Function





#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                DAO.Dispose()
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
