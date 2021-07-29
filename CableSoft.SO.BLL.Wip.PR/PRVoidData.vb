Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess

Public Class PRVoidData
    Inherits BLLBasic
    Implements IDisposable
    Private _PRDAL As New PRDAL(Me.LoginInfo.Provider)

    Public Sub New()
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As DbConnection)
        MyBase.New(LoginInfo, DBConnection)
        
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As DAO)
        MyBase.New(LoginInfo, DAO)
        
    End Sub
    ''' <summary>
    ''' 作廢停拆移單
    ''' </summary>
    ''' <param name="SNo">停拆移單號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function VoidData(ByVal SNo As String, ByVal Custid As Integer, ByVal ServiceType As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.SaveData(Me.LoginInfo, Me.DAO)
        Dim aRet As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = False}
        Try
            aRet.ResultBoolean = obj.VoidData(BLL.Utility.InvoiceType.PR, SNo)
            If aRet.ResultBoolean Then
                Using SOWipUtil As New CableSoft.SO.BLL.Utility.Wip(LoginInfo, DAO)
                    Dim RetCode As Int16 = 0
                    Dim P_RETMSG As String = ""
                    RetCode = SOWipUtil.SF_ADJSTATUS1(Nothing, Custid, 1, 0, LoginInfo.CompCode, ServiceType, P_RETMSG)
                    '更新客戶狀態(SF_ADJSTATUS1)
                    If RetCode < 0 Then
                        Throw New Exception(String.Format("Wip.SF_ADJSTATUS1-ReturnCode:{0},ReturnMessage:{1}", RetCode, P_RETMSG))
                    End If
                End Using
            End If
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If
            Try
                If _PRDAL IsNot Nothing Then
                    _PRDAL.Dispose()
                End If
                If MyBase.MustDispose AndAlso DAO IsNot Nothing Then
                    DAO.Dispose()
                End If
            Catch ex As Exception
            End Try
            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    Protected Overrides Sub Finalize()
        ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
