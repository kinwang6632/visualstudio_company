Imports CableSoft.BLL.Utility
Public Class BillingAPI603DAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overridable Function updInv014() As String
        Dim result As String = Nothing
        result = String.Format(" UPDATE INV014 SET ISOBSOLETE = 'Y', " &
                   " OBSOLETEID = {0}0, " &
                  " OBSOLETEREASON = (SELECT DESCRIPTION FROM INV006 " &
                                                        " WHERE IDENTIFYID1 = '1' AND IDENTIFYID2=0 AND ITEMID= {0}1 " &
                                                        " AND 1=1 )," &
                 " UPTTIME = SYSDATE,UPTEN = {0}2 " &
                 " WHERE IDENTIFYID1 = '1'  AND IDENTIFYID2 = 0 " &
                 " AND COMPID = {0}3  AND ALLOWANCENO = {0}4", Sign)
        Return result
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
