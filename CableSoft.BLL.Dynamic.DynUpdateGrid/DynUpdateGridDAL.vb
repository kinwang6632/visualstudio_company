Imports CableSoft.BLL.Utility
Public Class DynUpdateGridDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function GetCompCode(ByVal GroupId As String, ByVal strCD039 As String, ByVal strSO026 As String) As String
        If GroupId = "0" Then
            Return "Select A.CodeNo ,A.Description From " & strCD039 & " A Order By CodeNo"
        End If
        Return String.Format("Select distinct A.CodeNo ,A.Description " &
                             " From " & strCD039 & " A," & strSO026 & " B  " &
                             " Where Instr(',' ||B.CompStr|| ',' , ',' ||A.CodeNo|| ',') > 0 " &
                             " And UserId = {0}0 Order By CodeNO", Sign)
    End Function
    Friend Function chkAuthority(ByVal GroupField As String) As String
        Return String.Format("Select count(*) From SO029 Where Mid = {0}0 And  Group" & GroupField & "= 1", Sign)
    End Function
    Friend Function QueryMaster()
        Return String.Format("SELECT * FROM SO1111A WHERE SysProgramId = {0}0 " &
                              " AND 1=1 ", Sign)
    End Function
    Friend Function QueryCD052()
        Return "Select * From CD052"
    End Function
    Friend Function QueryDetail()
        Return String.Format("SELECT * FROM SO1111B WHERE ProgramId = {0}0  " &
                              " AND 1=1", Sign)
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
