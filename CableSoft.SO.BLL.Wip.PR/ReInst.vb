Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess

Public Class ReInst
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New ReInstDAL(Me.LoginInfo.Provider)

    Public Sub New()
    End Sub

    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub

    Public Function GetNewAddressData(ByVal ID As String, ByVal AddrNo As String) As DataTable
        Try
            Dim strSQL As String = String.Empty
            Dim strService As String = String.Empty
            Dim strAddQry As String = String.Empty

            Using dtCD046 As DataTable = DAO.ExecQry("Select CodeNo,Description From CD046 Order By Ord")
                For Each drCD046 As DataRow In dtCD046.Rows
                    strService = String.Format("{0},Max(Decode(B.ServiceType,'{1}',B.CustStatusName,'')) {1}CustStatusCode", strService, drCD046("CodeNo"))
                Next
            End Using
            strAddQry = String.Format(" A.ID='{0}' and C.AddrNO <>'{1}'", ID, AddrNo)
            strSQL = "Select A.*,B.* From (Select C.AddrNo,A.CustId,A.CustName,A.Tel1,A.Tel2,A.Tel3,C.CompCode" & strService & " From SO001 A,SO002 B,SO014 C" & _
                     " Where A.CustId = B.CustId(+) And A.CompCode = B.CompCode(+) And A.InstAddrNo(+)=C.AddrNo And A.CompCode(+)=C.CompCode And " & _
                    strAddQry & " Group By C.AddrNo,A.CustId,A.CustName,A.Tel1,A.Tel2,A.Tel3,C.CompCode) A ,SO014 B Where A.AddrNo = B.AddrNo And A.CompCode = B.CompCode Order By B.AddrSort"
            Return DAO.ExecQry(strSQL)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetCD046(ByVal CodeNo As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.CD046, CodeNo, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetSO002(ByVal AddrNo As Int32) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.SO002, AddrNo, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetCD005(ByVal ServiceType As String) As String
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.CD005, ServiceType, False)
            Dim CanUseCode As String = String.Empty
            For Each dr As DataRow In dt.Rows
                CanUseCode = String.Format("{0},{1}", CanUseCode, dr("CodeNO"))
            Next
            If CanUseCode.Length > 0 Then CanUseCode = CanUseCode.Substring(1)
            Return CanUseCode
        Catch ex As Exception
            Throw ex
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
            Try
                If _DAL IsNot Nothing Then
                    _DAL.Dispose()
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
