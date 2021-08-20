Imports System.Data.Common
Imports CableSoft.Utility.DataAccess
Imports CableSoft.BLL.Utility
Public Class BillingAPI260
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI


    Private _DAL As New BillingAPI260DAL(Me.LoginInfo.Provider)
    'Private _DAL As New BillingAPI253DALMultiDB(Me.LoginInfo.Provider)
    Private ServiceType As String = Nothing
    Private Lang As New CableSoft.BLL.Language.SO61.BillingAPI260Language
    Private SOUtil As CableSoft.SO.BLL.Utility.Utility = Nothing
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
    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        InData.Tables("SNo").Columns.Add(New DataColumn("DTVPRCode", GetType(Integer)))
        InData.Tables("SNo").Columns.Add(New DataColumn("DTVReasonCode", GetType(Integer)))
        InData.Tables("SNo").Columns.Add(New DataColumn("CMPRCode", GetType(Integer)))
        InData.Tables("SNo").Columns.Add(New DataColumn("CMReasonCode", GetType(Integer)))
        InData.Tables("SNo").Rows(0).Item("CMPRCode") = Integer.Parse(InData.Tables("SNo").Rows(0).Item("PRCode"))
        InData.Tables("SNo").Rows(0).Item("CMReasonCode") = Integer.Parse(InData.Tables("SNo").Rows(0).Item("ReasonCode"))

        Try
            Using bll253 As New BillingAPI253(Me.LoginInfo, Me.DAO)
                bll253.Ref3ServiceType = "I"
                Dim result As RIAResult = bll253.Execute(SeqNo, InData)                
                Return result
            End Using
        Catch ex As Exception
            Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = ex.ToString}
        End Try




    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
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
