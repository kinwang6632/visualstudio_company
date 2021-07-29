Imports CableSoft.BLL.Utility
Public Class ThirdDiscountDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryThirdDiscount() As String
        Dim aRet As String = String.Format("Select SO003B.RowID,SO003B.*, " &
                                           " A.FACISNO FACISNOA,B.FACISNO FACISNOB " &
                                           " From SO003B,SO004 A,SO004 B " &
                             " Where SO003B.CustId = {0}0 " &
                             " AND SO003B.FACISEQNO = A.SEQNO(+) " &
                             " AND SO003B.FACISEQNOB = B.SEQNO(+) ", Sign)


        'Dim aRet As String = String.Format("Select SO003B.*,A.FACISNO FACISNOA " &
        '                                   " From SO003B,SO004 A " &
        '                     " Where SO003B.CustId = {0}0 And SO003B.FaciSeqNo = {0}1 " &
        '                     " AND SO003B.FACISEQNO = A.SEQNO(+) ", Sign)

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
