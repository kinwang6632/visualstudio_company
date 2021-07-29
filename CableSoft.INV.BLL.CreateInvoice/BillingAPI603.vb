Imports System.Data.Common
Imports CableSoft.BLL.BillingAPI
Imports CableSoft.BLL.Utility
Public Class BillingAPI603
    Inherits BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI

    Private _DAL As New BillingAPI603DALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.BillingAPI603Language
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

    Public Function Execute(SeqNo As Integer, InData As DataSet) As RIAResult Implements IBillingAPI.Execute


        Dim aInvID As Object = InData.Tables("Inv").Rows(0).Item("InvID")
        Dim aCompCode As Integer = Integer.Parse(InData.Tables("Main").Rows(0).Item("Compcode"))
        Dim aAllowanceNo As String = InData.Tables("Inv").Rows(0).Item("AllowanceNo").ToString
        Dim aObsoleteId As String = InData.Tables("Inv").Rows(0).Item("ObsoleteId").ToString()
        Dim aCaller As String = InData.Tables("Main").Rows(0).Item("Caller").ToString

        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = "NO"}

        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim blnAutoClose As Boolean = False
        '#8706
        aCaller = InData.Tables("Inv").Rows(0).Item("Upden").ToString
        Me.LoginInfo.EntryName = aCaller
        'aCaller = Me.LoginInfo.EntryName
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
        Try
            DAO.ExecNqry(_DAL.updInv014, New Object() {aObsoleteId, aObsoleteId, aCaller, aCompCode, aAllowanceNo})
            If blnAutoClose Then
                If result.ResultBoolean Then
                    trans.Commit()
                Else
                    trans.Rollback()
                End If
            End If
            result.ResultBoolean = True
            result.ErrorCode = 1
            result.ErrorMessage = String.Format(Language.retMsg, aInvID, aAllowanceNo)
            Return result
        Catch ex As Exception
            If blnAutoClose Then
                trans.Rollback()
            End If
            result.ResultBoolean = False
            result.ErrorCode = -999
            result.ErrorMessage = ex.ToString
            Return result
        Finally

        End Try
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下方的 Finalize()。
            ' TODO: 將大型欄位設為 null。
        End If
        disposedValue = True
    End Sub

    ' TODO: 只有當上方的 Dispose(disposing As Boolean) 具有要釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' Visual Basic 加入這個程式碼的目的，在於能正確地實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更這個程式碼。請將清除程式碼放在上方的 Dispose(disposing As Boolean) 中。
        Dispose(True)
        ' TODO: 覆寫上列 Finalize() 時，取消下行的註解狀態。
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
