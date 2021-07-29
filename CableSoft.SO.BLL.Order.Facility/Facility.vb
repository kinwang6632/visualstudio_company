Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess

'Imports Lang = CableSoft.SO.BLL.Order.Facility.FacilityLanguage

Public Class Facility
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New FacilityDALMultiDB(LoginInfo.Provider)
    'Private _DAO As New DAO(LoginInfo.Provider, LoginInfo.ConnectionString)
    Private _GetSchema As Boolean
    Private _Disposed As Boolean ' 偵測多餘的呼叫
    Private Language As New CableSoft.BLL.Language.SO61.FacilityLanguage
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
                    If Language IsNot Nothing Then
                        Language.Dispose()
                        Language = Nothing
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
    ''' 取得服務類別
    ''' </summary>
    ''' <returns>回傳服務類別</returns>
    ''' <remarks></remarks>
    Public Function GetServiceTypeCode() As DataTable
        GetServiceTypeCode = Nothing
        Try
            Dim qry As String = _DAL.GetServiceTypeCode
            Return DAO.ExecQry(qry, _GetSchema)
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function CanEdit(ByVal Facility As DataSet) As RIAResult
        Dim aRet As New RIAResult() With {.ResultBoolean = True}
        If Not DBNull.Value.Equals(Facility.Tables("Facility").Rows(0).Item("BPCode")) Then
            aRet.ResultBoolean = False
            aRet.ErrorMessage = Language.CanNotEdit
            aRet.ErrorCode = -1
        End If
        Return aRet
    End Function
    Public Function CanDelete(ByVal Facility As DataSet) As RIAResult
        Dim aRet As New RIAResult() With {.ResultBoolean = True}
        If Not DBNull.Value.Equals(Facility.Tables("Facility").Rows(0).Item("BPCode")) Then
            aRet.ResultBoolean = False
            aRet.ErrorMessage = Language.CanNotDel
            aRet.ErrorCode = -1
        End If
        Return aRet
    End Function
    Public Function ChkDataOk(ByVal Facility As DataSet) As RIAResult
        Dim aRet As New RIAResult() With {.ResultBoolean = True}
        If DBNull.Value.Equals(Facility.Tables("Facility").Rows(0).Item("ServiceType")) Then
            aRet.ResultBoolean = False
            aRet.ErrorMessage = Language.MustServiceType
            aRet.ErrorCode = -1
            Return aRet
        End If
        If DBNull.Value.Equals(Facility.Tables("Facility").Rows(0).Item("FaciName")) Then
            aRet.ResultBoolean = False
            aRet.ErrorMessage = Language.MustFaciName
            aRet.ErrorCode = -2
            Return aRet
        End If
        If DBNull.Value.Equals(Facility.Tables("Facility").Rows(0).Item("BuyCode")) Then
            aRet.ResultBoolean = False
            aRet.ErrorMessage = Language.MustBuyType
            aRet.ErrorCode = -3
            Return aRet
        End If
        If DBNull.Value.Equals(Facility.Tables("Facility").Rows(0).Item("CodeNo")) Then
            aRet.ResultBoolean = False
            aRet.ErrorMessage = Language.MustWorkType
            aRet.ErrorCode = -4
            Return aRet
        End If
        Return aRet
    End Function
    Public Function GetFaciSeqNo() As String
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
        Try
            Return obj.GetFaciSeqNo(True)
        Catch ex As Exception
            Throw ex
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得可選設備項目
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>回傳可選設備項目</returns>
    ''' <remarks></remarks>
    Public Function QueryFaciCode(ByVal ServiceType As String) As DataTable
        QueryFaciCode = Nothing
        Try
            If ServiceType IsNot Nothing Then
                Return DAO.ExecQry(_DAL.QueryFaciCode,
                                   ServiceType,
                                   "Facilities",
                                   "Facility",
                                   _GetSchema)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 取得可選買賣方式
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>回傳可選買賣方式</returns>
    ''' <remarks></remarks>
    Public Function QueryBuyCode(ByVal ServiceType As String) As DataTable
        QueryBuyCode = Nothing
        Try
            If ServiceType IsNot Nothing Then
                Return DAO.ExecQry(_DAL.QueryBuyCode,
                                   ServiceType,
                                   _GetSchema)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 取得可選派工類別
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>回傳可選派工類別</returns>
    ''' <remarks></remarks>
    Public Function QueryWipCode(ByVal ServiceType As String) As DataTable
        QueryWipCode = Nothing
        Try
            If ServiceType IsNot Nothing Then
                Return DAO.ExecQry(_DAL.QueryWipCode,
                                   ServiceType,
                                   _GetSchema)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

End Class
