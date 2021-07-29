Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess

Imports Lang = CableSoft.SO.BLL.Order.CutDate.CutDateLanguage

Public Class CutDate
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New CutDateDALMultiDB(LoginInfo.Provider)
    'Private _DAO As New DAO(LoginInfo.Provider, LoginInfo.ConnectionString)
    Private _GetSchema As Boolean
    Private _Disposed As Boolean ' 偵測多餘的呼叫

    Public Sub New()
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo)
        MyBase.New(LoginInfo)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub

    Protected Overrides Sub Finalize()
        Try
            Dispose(False)
        Finally
            MyBase.Finalize()
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Sub Dispose(ByVal disposing As Boolean) ' IDisposable
        If Not _Disposed Then
            If disposing Then
                ' TODO: 釋放其他狀態 (Managed 物件)。
                Try
                    If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                        DAO.Dispose()
                    End If
                    If _DAL IsNot Nothing Then
                        _DAL.Dispose()
                    End If
                Catch ex As Exception
                    Throw ex
                End Try
            End If
            ' TODO: 釋放您自己的狀態 (Unmanaged 物件) 或將大型欄位設定為 null。
        End If
        _Disposed = True
    End Sub

    ''' <summary>
    ''' 取得查詢可用下收切齊資料
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <returns>回傳查詢可用下收切齊資料</returns>
    ''' <remarks></remarks>
    Public Function QueryCanCutDate(ByVal CustId As Integer, ByVal ResvTime As Date) As DataTable
        QueryCanCutDate = Nothing
        Try
            If CustId > 0 Then
                Return DAO.ExecQry(_DAL.QueryCanCutDate, New Object() {CustId, ResvTime, CustId, ResvTime})
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

End Class
