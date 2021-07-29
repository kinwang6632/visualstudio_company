Public Class BillingAPI258DAL
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
    Friend Function GetWorkCodeForCitemCode(ByVal inPRCode As Integer) As String
        Dim strSQL As String
        Dim strWhere As String = String.Empty
        '#7666 2018.02.12 by Corey 需求增加傳入PRCODE 則指定該PRCODE。
        If inPRCode > 0 Then strWhere = String.Format("and CodeNo={0}", inPRCode)
        strSQL = String.Format("Select CodeNo,Description,ServiceType From CD007 Where REFNO=7 and nvl(StopFlag,0)=0 {0}", strWhere)
        Return strSQL
    End Function
    Friend Function GetCustomerData() As String
        Dim strSQL As String
        strSQL = String.Format("Select A.CustStatusCode,A.WipCode3,B.InstAddrNo From SO002 A,SO001 B Where A.CustId = B.CustId And A.CustId = {0}0 And A.ServiceType = {0}1", Sign)
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
    Friend Function GetAddress() As String
        Dim strSQL As String
        strSQL = String.Format("Select AddrNo,Address From SO014 Where AddrNo = {0}0", Sign)
        Return strSQL
    End Function

    Friend Function GetChargeCitemCode(ByVal intType As Integer, strCitemCode As String) As String
        'intType: 0=SO003 1=SO033
        Dim strTable As String = "SO003"
        If intType = 1 Then strTable = "SO033"
        Return String.Format("Select CitemCode From {1} Where Custid = {0}0 and Faciseqno={0}1 and CitemCode in ({2})", Sign, strTable, strCitemCode)
    End Function

    Friend Function FacilityData(ByVal FaciRefNo As String) As String
        Return String.Format("Select * From SO004 A,CD022 B Where A.CustID={0}0 and A.ServiceType={0}1 and SEQNO ={0}2" &
                             " and A.FaciCode=B.CodeNo and B.RefNo In ({1})", Sign, FaciRefNo)
    End Function

    Friend Function CheckPrDouble() As String
        Return String.Format("Select * from SO009 Where Custid={0}0 and ServiceType={0}1 and PrCode={0}2 and ReturnCode Is Null and FinTime Is Null ", Sign)
    End Function

    Friend Function CheckPrChangeFacility() As String
        Return String.Format("Select * from SO004D where Custid={0}0 and SNO={0}1 ", Sign)
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
