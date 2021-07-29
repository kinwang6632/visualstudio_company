Imports System.Data.Common
Imports CableSoft.BLL.Utility
Public Class CPEMAC
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New CPEMACDALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.CPEMACLanguage
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
    Public Function GetCPEMAC(ByVal CustId As Integer, ByVal FaciSeqNo As String) As DataSet
        Return DAO.ExecQry(_DAL.GetCPEMAC,
                    New Object() {CustId, FaciSeqNo}).DataSet
    End Function
    Public Function ChkDataOK(ByVal EditMode As EditMode,
                             ByVal dtCPEMAC As DataTable,
                             ByVal CPEMAC As String, ByVal IPAddress As String,
                             ByVal FaciSeqNo As String, FixIPcount As Integer) As RIAResult
        Dim aRet As New RIAResult()
        Dim intCount As Byte = 0
        aRet.ResultBoolean = True
        If String.IsNullOrEmpty(CPEMAC) Then
            aRet.ErrorCode = -1
            aRet.ErrorMessage = Language.MustCPEField
            aRet.ResultBoolean = False
            Return aRet
        End If
        If Not String.IsNullOrEmpty(IPAddress) Then
            Dim aErrMsg As String = ChkHaveIPAddress(IPAddress, Me.LoginInfo.CompCode, "CPE".ToUpper)
            If Not String.IsNullOrEmpty(aErrMsg) Then
                aRet.ErrorCode = -2
                aRet.ErrorMessage = aErrMsg
                aRet.ResultBoolean = False
                Return aRet
            End If
            aErrMsg = chkIPAddressDup(IPAddress, FaciSeqNo)
            If Not String.IsNullOrEmpty(aErrMsg) Then
                aRet.ErrorCode = -3
                aRet.ErrorMessage = aErrMsg
                aRet.ResultBoolean = False
                Return aRet
            End If
        End If
        If CPEMAC.Replace(":", String.Empty).Replace("-", String.Empty).ToString.Length <> 12 Then
            aRet.ErrorCode = -4
            aRet.ErrorMessage = Language.CPELenError
            aRet.ResultBoolean = False
            Return aRet
        End If
        If EditMode = CableSoft.BLL.Utility.EditMode.Append Then intCount = 1

        

        'Dim Count As Integer = dtCPEMAC.AsEnumerable.Count(Function(rw As DataRow)
        '                                                       Return DBNull.Value.Equals(rw.Item("StopDate"))
        '                                                   End Function)


        'If FixIPcount <> Count + intCount Then
        '    aRet.ResultBoolean = True
        '    aRet.ErrorCode = -99
        '    aRet.ResultXML = String.Format("固定IP數{0}個, CPE MAC 設定{1}個, 是否要存檔??",
        '                                     FixIPcount, Count + intCount)
        'End If
        
        Return aRet
    End Function
    Private Function chkIPAddressDup(ByVal IPAddress As String, ByVal FaciSeqNo As String) As String
        Dim dt As DataTable = DAO.ExecQry(_DAL.chkIPAddressDup, New Object() {IPAddress, FaciSeqNo})
        If dt.Rows.Count > 0 Then
            Return String.Format(Language.IPAddressDouble,
                                 dt.Rows(0).Item("CustId").ToString, dt.Rows(0).Item("CPEMAC").ToString)
        End If
        Return Nothing
    End Function
    Private Function ChkHaveIPAddress(ByVal IPAddress As String,
                                      ByVal CompCode As Integer,
                                      ByVal strIPType As String) As String

        Dim dt As DataTable = DAO.ExecQry(_DAL.ChkHaveIPAddress, New Object() {IPAddress, CompCode, strIPType})
        If dt.Rows.Count = 0 Then
            Return Language.IPNotInList
        End If
        If Int32.Parse(dt.Rows(0).Item("UseFlag")) = 1 Then
            Return Language.IPIsUse
        End If
        Return Nothing
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                End If
                If Language IsNot Nothing Then
                    Language.Dispose()
                    Language = Nothing
                End If
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
