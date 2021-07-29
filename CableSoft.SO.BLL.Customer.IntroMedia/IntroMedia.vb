Imports System.Data.Common
Imports CableSoft.BLL.Utility
Public Class IntroMedia
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New IntroMediaDALMultiDB(Me.LoginInfo.Provider)
    Private Const FCurrectTableName As String = "Intro"
    Private Const FPKField As String = "RefNo"

    Private result As New RIAResult()
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

    Public Function CanView() As CableSoft.BLL.Utility.RIAResult
        Return New RIAResult() With {.ErrorCode = 0, .ErrorMessage = String.Empty, .ResultBoolean = True}
    End Function
    Public Function GetIntroId(ByVal MediaRefNo As Integer, ByVal IntroId As String) As DataTable

        If MediaRefNo = 1 AndAlso String.IsNullOrEmpty(IntroId) Then
            IntroId = "X"
        End If

        Select Case MediaRefNo
            Case 1
                Return DAO.ExecQry(_DAL.GetIntroId(MediaRefNo), New Object() {IntroId})
            Case Else
                Return DAO.ExecQry(_DAL.GetIntroId(MediaRefNo))
        End Select


    End Function
    Public Function GetIntroData(ByVal MediaRefNo As Integer, ByVal Search1 As String, ByVal Search2 As String) As DataTable
        Dim aWhere As String = String.Empty
        aWhere = _DAL.GetWhere(MediaRefNo, Search1, Search2)
        Return DAO.ExecQry(_DAL.GetIntroData(MediaRefNo) & aWhere)
        'Select Case MediaRefNo
        '    Case 1

        '    Case Else
        '        Return DAO.ExecQry(_DAL.GetIntroId(MediaRefNo))
        'End Select
    End Function
    Public Function keyCodeSearch(ByVal MediaRefNo As Integer, ByVal searchWord As String) As DataTable
        Dim aSQL As String = Nothing
        'Select Case MediaRefNo
        '    Case 1
        '        aSQL = String.Format("Select CustName as Description ,CustId as CodeNo From SO001 Where CustId = {0}", searchWord)
        '    Case 2
        '        aSQL = String.Format("SELECT EmpName as Description, EmpNo as CodeNo FROM CM003 WHERE EmpNo = '{0}'", searchWord)
        '    Case 3
        '        aSQL = String.Format("SELECT NameP as Description, IntroID as CodeNo FROM SO013 WHERE IntroID = '{0}'", searchWord)
        '    Case Else
        '        aSQL = String.Format("Select CustName as Description ,CustId as CodeNo From SO001 Where CustId = {0}", searchWord)
        'End Select
        aSQL = _DAL.GetkeyCodeSearchSQL(MediaRefNo, searchWord)
        Return DAO.ExecQry(aSQL)
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
