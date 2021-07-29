Public Class BillingAPI255DAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)

    End Sub
    Friend Function GetWorkCode() As String
        Dim strSQL As String
        strSQL = String.Format("Select GroupNO,Resvdatebefore,WorkUnit From CD007 Where CodeNo = {0}0", Sign)
        Return strSQL
    End Function
    Friend Function QueryCD049Description() As String
        Return String.Format("Select Description From CD049 where CodeNo = {0}0 And Nvl(Stopflag,0) =0", Sign)
    End Function
    Friend Function getMediaRefNo() As String
        Return String.Format("Select Nvl(RefNo,0) RefNo From CD009 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QueryCD009Description() As String
        Return String.Format("Select Description From CD009 where CodeNo = {0}0 And Nvl(Stopflag,0) =0", Sign)
    End Function
    Friend Function GetCustomerData() As String
        Dim strSQL As String
        strSQL = String.Format("Select B.CustName,A.CustStatusCode,A.WipCode3,B.InstAddrNo From SO002 A,SO001 B Where A.CustId = B.CustId And A.CustId = {0}0 And A.ServiceType = {0}1", Sign)
        Return strSQL
    End Function
    Friend Overridable Function GetContactDetailData() As String
        Dim strSQL As String
        strSQL = String.Format("Select * From (Select * From SO006A Where SeqNo = {0}0 Order By AutoSerialNo) Where Rownum = 1", Sign)
        Return strSQL
    End Function
    Friend Function GetContactData() As String
        Dim strSQL As String
        strSQL = String.Format("Select * From SO006 Where SeqNo = {0}0", Sign)
        Return strSQL
    End Function
    Friend Function GetEmpName() As String
        Dim strSQL As String
        strSQL = String.Format("Select EmpName From CM003 Where EmpNo = {0}0", Sign)
        Return strSQL
    End Function
    Friend Function GetReasonDescName() As String
        Dim strSQL As String
        strSQL = String.Format("Select Description From CD014A Where CodeNo = {0}0", Sign)
        Return strSQL
    End Function
    Friend Function GetReInstCode() As String
        Dim strSQL As String
        'strSQL = String.Format("Select CodeNo,Description From CD005 Where RefNo in (1,2,5) And ServiceType = {0}0 And ReInstAcrossFlag = 2 And StopFlag = 0 Order By Decode(RefNo,1,0,5,1,2)", Sign)
        strSQL = String.Format("Select CodeNo,Description From CD005 Where RefNo in (1,2,5) And ServiceType = {0}0 And ReInstAcrossFlag = 2 And StopFlag = 0 Order By (case refno  " &
          " when 1 then 0   when 5 then 1 else 2 end)", Sign)
        Return strSQL
    End Function
    Friend Function GetSO137() As String
        Dim strSQL As String
        strSQL = String.Format("Select * From SO137 Where ID = {0}0", Sign)
        Return strSQL
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean

    ' 偵測多餘的呼叫

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
