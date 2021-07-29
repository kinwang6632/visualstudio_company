Public Class ReInstDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub

    Friend Function CD046() As String
        Return String.Format("Select * from CD046 where CodeNO={0}0", Sign)
    End Function

    Friend Function SO002() As String
        Return String.Format("Select A.CustStatusCode,A.ServiceType From SO002 A,SO001 B Where A.CustId = B.CustId And A.CompCode = B.CompCode And B.InstAddrNo = {0}0", Sign)
    End Function

    Friend Function CD005() As String
        Return String.Format("Select * From CD005 Where RefNO in (1,5) And ReInstAcrossFlag > 0 And ServiceType = {0}0", Sign)
    End Function
    Friend Function QryNewAddressData(ByVal strService As String, ByVal strAddQry As String) As String
        Return "Select A.*,B.* From (" &
                    "Select E.AddrNo,A.CustId,A.CustName,F.Conttel Tel1,A.Tel2,F.ContMobile Tel3,E.CompCode" & strService &
                    "  From SO001 A left join SO002  B on A.CustId = B.CustId And A.CompCode = B.CompCode " &
                    " right join  SO014 E on A.InstAddrNo=E.AddrNo And A.CompCode=E.CompCode right join  SO137 F on A.ID=F.ID  " &
                    "  Where 1 =1 And A.CustID Is Not Null " & strAddQry &
                    "  Group By E.AddrNo,A.CustId,A.CustName,F.Conttel,A.Tel2,F.ContMobile,E.CompCode) A ,SO014 B Where A.AddrNo = B.AddrNo And A.CompCode = B.CompCode Order By B.AddrSort"
    End Function
    Friend Function QueryCD046() As String
        Return "Select CodeNo,Description From CD046 Order By Ord"
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
