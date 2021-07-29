Imports CableSoft.BLL.Utility

Public Class CreditIBonDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)

    End Sub

    '單據編號取收費資料
    Friend Function GetRealCharge(ByVal Custid As Integer, ByVal MediaBillNo As String, ByVal intCrossCustCombine As Integer) As String
        Dim strSQL As String = Nothing
        If intCrossCustCombine > 0 Then
            strSQL = String.Format("Select A.*  From SO033 A Where A.MediaBillNo='{0}'", MediaBillNo)
        Else
            strSQL = String.Format("Select A.*  From SO033 A Where A.Custid={0} And A.MediaBillNo='{1}'", Custid, MediaBillNo)
        End If
        Return strSQL
    End Function
    '判斷SO033是否有該媒体單號
    Friend Function GetChargeCnt(ByVal MediaBillNo As String) As String
        Dim strSQL As String = Nothing
        strSQL = String.Format("Select count(*)  From SO033 A Where A.MediaBillNo='{0}'", MediaBillNo)
        Return strSQL
    End Function
    '如果有媒体單號再判斷媒体單號是否屬於該客編
    Friend Function GetChargeCustid(ByVal MediaBillNo As String, ByVal Custid As Integer) As String
        Dim strSQL As String = Nothing
        strSQL = String.Format("Select count(*)  From SO033 A Where A.MediaBillNo='{0}' AND CUSTID <> {1}", MediaBillNo, Custid)
        Return strSQL
    End Function
    '是否金額不符   '990127 #5499 調整金額加總時要過濾作廢
    Friend Function GetChargeAmount(ByVal MediaBillNo As String, ByVal Custid As Integer, ByVal intCrossCustCombine As Integer) As String
        Dim strSQL As String = Nothing
        If intCrossCustCombine > 0 Then
            strSQL = String.Format("Select nvl(SUM(A.ShouldAmt),0)  From SO033 A Where A.MediaBillNo='{0}' And A.CancelFlag=0", MediaBillNo)
        Else
            strSQL = String.Format("Select nvl(SUM(A.ShouldAmt),0)  From SO033 A Where A.Custid={0} And A.MediaBillNo='{1}' And A.CancelFlag=0", Custid, MediaBillNo)
        End If
        Return strSQL
    End Function
    '是否作廢   '990127 #5499 調整若整張全都作廢才能算作廢,單一筆不算(總筆數=總作廢筆數)
    Friend Function GetChargeCancel(ByVal MediaBillNo As String, ByVal Custid As Integer, ByVal intCrossCustCombine As Integer) As String
        Dim strSQL As String = Nothing
        If intCrossCustCombine > 0 Then
            strSQL = String.Format("Select count(*)  From SO033 A Where A.MediaBillNo='{0}' And A.CancelFlag=1", MediaBillNo)
        Else
            strSQL = String.Format("Select count(*)  From SO033 A Where A.Custid={0} And A.MediaBillNo='{1}' And A.CancelFlag=1", Custid, MediaBillNo)
        End If
        Return strSQL
    End Function
    Friend Function GetChargeCancel2(ByVal MediaBillNo As String, ByVal Custid As Integer, ByVal intCrossCustCombine As Integer) As String
        Dim strSQL As String = Nothing
        If intCrossCustCombine > 0 Then
            strSQL = String.Format("Select count(*)  From SO033 A Where A.MediaBillNo='{0}' ", MediaBillNo)
        Else
            strSQL = String.Format("Select count(*)  From SO033 A Where A.Custid={0} And A.MediaBillNo='{1}' ", Custid, MediaBillNo)
        End If
        Return strSQL
    End Function
    '是否已收   '990224 #5499 測試報告,因作廢會填RealDate,要增加條件過濾  '990503 #5641 調整判斷已收的條件 '990519 #5564 調整規格,已收的判斷,CD013.RefNo=3,7 or PayOk=1
    Friend Function GetChargePayOK(ByVal MediaBillNo As String, ByVal Custid As Integer, ByVal intCrossCustCombine As Integer) As String
        Dim strSQL As String = Nothing
        If intCrossCustCombine > 0 Then
            strSQL = String.Format("Select count(*)  From SO033 A Where A.MediaBillNo='{0}' And A.CancelFlag=0 And (A.UCCode is Null or A.UCCode in (Select CodeNo From CD013 Where RefNo in(3,7) or PayOK=1))", MediaBillNo)
        Else
            strSQL = String.Format("Select count(*)  From SO033 A Where A.Custid={0} And A.MediaBillNo='{1}' And A.CancelFlag=0 And (A.UCCode is Null or A.UCCode in (Select CodeNo From CD013 Where RefNo in(3,7) or PayOK=1))", Custid, MediaBillNo)
        End If
        Return strSQL
    End Function
    '取得櫃台已收UCCode
    Friend Function GetUCCode() As String
        Dim strSQL As String = Nothing
        'strSQL = String.Format("Select CodeNo,Description From CD013 " &
        '                                            "Where RefNo=7 And Nvl(StopFlag,0)=0  And (ServiceType is null or ServiceType='{0}') Order by CodeNo", ServiceType)
        strSQL = "Select CodeNo,Description From CD013 Where RefNo=7 And Nvl(StopFlag,0)=0  Order by CodeNo"
        Return strSQL
    End Function
    '更新SO033收費設定資料
    Friend Function UpdRealCharge(ByVal UCCode As String, ByVal UCName As String, ByVal strUpdEn As String, ByVal strUpdTime As String,
                                                                 ByVal MediaBillNo As String) As String
        Dim strSQL As String = String.Format("Update SO033 Set UCCode={0},UCName='{1}' " & _
                                                                                ",UpdEn= '{2}',UpdTime='{3}'  " & _
                                                                                " Where MediaBillNo='{4}' And Cancelflag=0", Integer.Parse(UCCode), UCName, strUpdEn, strUpdTime, MediaBillNo)
        Return strSQL
    End Function
    '更新客戶來源資料 
    Friend Function UpdMSData(ByVal ShopID As String, ByVal TableName As String, ByVal strErrorMsg As String, ByVal MediaBillNo As String) As String
        'Dim strSQL As String = String.Format("Update {0}1 Set Status={0}2,StatusDate=to_date({0}4,'yyyy/MM/dd hh24:mi:ss')  " & _
        '                                                                        " Where MediaBillNo={0}3 And ShopID={0}0 ", Sign)
        Dim strSQL As String = String.Format("Update {2} Set Status='{3}',StatusDate={0}0  " & _
                                                                        " Where MediaBillNo='{4}' And ShopID='{1}' ", Sign, ShopID, TableName, strErrorMsg, MediaBillNo)

        Return strSQL
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
