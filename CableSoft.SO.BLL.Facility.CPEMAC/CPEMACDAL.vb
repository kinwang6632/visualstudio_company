Imports CableSoft.BLL.Utility
Public Class CPEMACDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Public Function ChkHaveIPAddress() As String
        Dim aRet As String = String.Format("Select Nvl(UseFlag,0) UseFlag  " &
                                         " From SO048 " &
                                         " Where IPAddress = {0}0 And CompCode = {0}1 And IPNature = {0}2", Sign)
        Return aRet
    End Function
    Public Function chkIPAddressDup() As String

        Dim aRet As String = String.Format("Select B.CustID,B.CPEMAC " &
                                           " From SO004 A,SO004C B " &
                                           " Where A.SeqNo = B.SeqNo " &
                                           " And A.PRDate Is null And B.IPAddress ={0}0 " &
                                           " And B.StopDate is null " &
                                           " And A.SeqNo <> {0}1 ", Sign)
        Return aRet
    End Function
    Friend Function GetCPEMAC() As String
        Return String.Format("SELECT * FROM SO004C " &
                                         " WHERE CUSTID={0}0 " &
                                         " AND SEQNO = {0}1", Sign)
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
