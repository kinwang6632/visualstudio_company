
Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess

Imports Lang = CableSoft.SO.BLL.Order.Deposit.DepositLanguage

Public Class Deposit
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New DepositDALMultiDB(LoginInfo.Provider)
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
                Catch ex As Exception
                    Throw ex
                End Try
            End If
            ' TODO: 釋放您自己的狀態 (Unmanaged 物件) 或將大型欄位設定為 null。
        End If
        _Disposed = True
    End Sub

    ''' <summary>
    ''' 查詢可選付款種類
    ''' </summary>
    ''' <param name="BPCode">組合產品代碼</param>
    ''' <param name="CitemCode">收費項目代碼</param>
    ''' <returns>回傳查詢可選付款種類</returns>
    ''' <remarks></remarks>
    Public Function QueryPTCode(ByVal BPCode As String,
                                ByVal CitemCode As Integer) As DataTable
        QueryPTCode = Nothing
        Try

            Return DAO.ExecQry(_DAL.QueryPTCode,
                               New Object() {BPCode, CitemCode},
                               _GetSchema)

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 選擇付款種類
    ''' </summary>
    ''' <param name="BPCode">組合產品代碼</param>
    ''' <param name="CitemCode">收費項目代碼</param>
    ''' <param name="PTCode">付款種類代碼</param>
    ''' <returns>回傳選擇付款種類</returns>
    ''' <remarks></remarks>
    Public Function ChoosePTCode(ByVal BPCode As String,
                                 ByVal CitemCode As Integer,
                                 ByVal PTCode As Integer) As DataTable
        ChoosePTCode = Nothing
        Try

            Return DAO.ExecQry(_DAL.ChoosePTCode,
                               New Object() {BPCode, CitemCode, PTCode},
                               _GetSchema)

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 查詢可選收費方式
    ''' </summary>
    ''' <returns>回傳查詢可選收費方式</returns>
    ''' <remarks></remarks>
    Public Function QueryCMCode() As DataTable
        QueryCMCode = Nothing
        Try
            Return DAO.ExecQry(_DAL.QueryCMCode,_GetSchema)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 執行保證金設定
    ''' </summary>
    ''' <param name="Order">訂單單號</param>
    ''' <param name="ContName">本票開票人</param>
    ''' <param name="CheckNo">本票票號</param>
    ''' <param name="CMCode">收費方式代碼</param>
    ''' <param name="CMName">收費方式</param>
    ''' <param name="PTCode">付款種類代碼</param>
    ''' <param name="PTName">付款種類</param>
    ''' <returns>回傳處理結果</returns>
    ''' <remarks></remarks>
    Public Function Execute(ByVal Order As Object,
                            ByVal ContName As String,
                            ByVal CheckNo As String,
                            ByVal CMCode As Integer,
                            ByVal CMName As String,
                            ByVal PTCode As Integer,
                            ByVal PTName As String) As RIAResult

        Dim r As New RIAResult() With {.ResultBoolean = False, .ErrorCode = -1}

        Execute = r

        Try
            r.ResultBoolean = True
            r.ErrorCode = 0
            Return r

        Catch ex As Exception
            Throw
        End Try

        '1.	目的: 更新訂購明細保證金資料。
        '2.說明()
        '(1) 呼叫ChoosePTCode取得收費項目及金額。
        '(2) 回填欄位：
        'Charge.CMCode = <CMCode>
        'Charge.CMName = <CMName>
        'Charge.Citemcode = <ChoosePTCode.Codeno>
        'Charge.CitemName = <ChoosePTCode.Description>
        'Charge.Amount　= <ChoosePTCode.Amount>
        'Charge.PTCode = <PTCode>
        'Charge.PTName = <PTName>

        '(3) 依畫面欄位回填訂購設備明細：
        '條件：Charge.Faciseqno=Facility.Faciseqno
        '回填欄位：Facility. Deposit = <ChoosePTCode.Amount>

    End Function


End Class
