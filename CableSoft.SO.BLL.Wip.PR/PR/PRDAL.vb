Imports SaveLanguage = CableSoft.BLL.Language.SO61.WipPRLanguage
Public Class PRDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub

    Friend Function GetServiceType(ByVal CanUseServiceType As String) As String
        Dim strWhere As String = String.Empty
        If Not String.IsNullOrEmpty(CanUseServiceType) Then strWhere = String.Format("Where CodeNo in ('{0}')", CanUseServiceType.Replace(",", "','"))
        Return String.Format("Select CodeNo,Description,DependService From CD046 {0} Order by CodeNo", strWhere)
    End Function

    Friend Function GetALLPRCode()
        Return String.Format("Select * From CD007 Order by CodeNo")
    End Function
    
    Friend Function GetPRCode(ByVal CanUseRefNo As String, ByVal CanNotUseRefNo As String, ByVal WipCodeValueStr As String,
                              ByVal blnReInstFilter As Boolean, ByVal ReInstAcrossFlag As Boolean, ByVal filterMoveRefNo As Boolean) As String
        Dim strWhere As String = String.Empty
        '過濾可使用的參考號
        If Not String.IsNullOrEmpty(CanUseRefNo) Then
            strWhere = String.Format(" {0} and RefNo in ({1})", strWhere, CanUseRefNo)
        End If
        '過濾不可使用的參考號
        If Not String.IsNullOrEmpty(CanNotUseRefNo) Then
            strWhere = String.Format(" {0} and RefNo Not in ({1})", strWhere, CanNotUseRefNo)
        End If
        '過濾可使用個工單號碼
        If Not String.IsNullOrEmpty(WipCodeValueStr) Then
            strWhere = String.Format(" {0} and CodeNo in ({1})", strWhere, WipCodeValueStr)
        End If
        '過濾是否要移機跨區
        If blnReInstFilter Then
            If ReInstAcrossFlag Then
                strWhere = String.Format(" {0} and ReInstAcrossFlag >0", strWhere)
            Else
                strWhere = String.Format(" {0} and ReInstAcrossFlag =0", strWhere)
            End If
        End If
        '#8257 2019.04.02 by Corey 增加過濾MoveRefNo。功能:GetCanMoveServiceType使用
        If filterMoveRefNo Then
            strWhere = String.Format("{0} And MoveRefNo Is Null", strWhere)
        End If

        Return String.Format("Select * From CD007 Where (ServiceType ={0}0 Or ServiceType is null) And StopFlag = 0 {1} Order by CodeNo", Sign, strWhere)
    End Function

    Friend Function GetPRCodeByContactRefNo() As String
        Return String.Format("Select CodeNo,Description,RefNo,WorkUnit,GroupNo From CD007 Where (ServiceType ={0}0 Or ServiceType is null) And StopFlag = 0 And RefNo = {0}1  Order by CodeNo", Sign)
    End Function

    Friend Function GetPRReasonCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD014 Where (ServiceType = {0}0 Or ServiceType is null) And StopFlag = 0 Order by CodeNo", Sign)
    End Function

    Friend Function GetPRReasonDescCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD014A Where (ServiceType = {0}0 Or ServiceType is null) And StopFlag = 0 And CodeNo in (Select ReasonDescCode From CD014B Where ReasonCode = {0}1) Order by CodeNo", Sign)
    End Function

    Friend Function GetGroupCode() As String
        Return String.Format("Select A.* From CD003 A Where Exists (Select 1 From CD002CM003 B Where A.CodeNo = B.EmpNo And ServCode = {0}0 And Type = 3) And StopFlag = 0 Order by CodeNo", Sign)
    End Function
    Friend Function GetCD002() As String
        Return String.Format("Select CODENO,DESCRIPTION FROM CD002 WHERE CODENO = {0}0 AND COMPCODE = {0}1 ", Sign)
    End Function
    Friend Function GetWorkerEn() As String
        Return "Select EmpNo CodeNo,EmpName Description,RefNo From CM003 Where StopFlag = 0 Order by EmpNo"
    End Function

    Friend Function GetReturnCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD015 Where StopFlag = 0 And (ServiceType is null or ServiceType ={0}0) Order by CodeNo", Sign)
    End Function

    Friend Function GetReturnDescCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD072 Where StopFlag = 0 And (ServiceType is null or ServiceType like ({0}0)) Order by CodeNo", Sign)
    End Function

    Friend Function GetSignEn() As String
        Return "Select EmpNo CodeNo,EmpName Description,RefNo From CM003 Where StopFlag = 0 Order by EmpNo"
    End Function

    Friend Function GetSatiCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD026 Where StopFlag = 0 And (ServiceType is null or ServiceType ={0}0) Order by CodeNo", Sign)
    End Function

    Friend Function GetCustomer(ByVal ServiceType As String) As String
        Dim ServiceSQL As String
        If String.IsNullOrEmpty(ServiceType) Or ServiceType = "X" Then
            ServiceSQL = ""
        Else
            ServiceSQL = String.Format(" And A.ServiceType = '{0}'", ServiceType)
        End If
        Return String.Format("SELECT A.*,B.ServArea,B.ClassName1,B.InstAddress,B.Tel1,Nvl(B.Balance,0) Balance From SO002 A,SO001 B " & _
                            " Where A.CustId = B.CustId And A.CustId = {0}0 {1}", Sign, ServiceSQL)
    End Function

    Friend Function GetSO042() As String
        Return String.Format("SELECT * FROM SO042 WHERE SERVICETYPE={0}0 ", Sign)
    End Function

    Friend Function GetChangePRCode() As String
        Return String.Format("Select Count(*) From SO004 Where CustId ={0}0 And PRDate is null And FaciCode in (Select CodeNo From CD022 Where RefNo in ({0}1)", Sign)
    End Function

    Friend Function GetCD007() As String
        Return String.Format("Select * From CD007 Where CodeNo = {0}0 Order by CodeNo", Sign)
    End Function

    Friend Function GetWipData(Optional ByVal Field As String = "*") As String
        Return String.Format("Select {1} From SO009 WHERE SNO={0}0 ", Sign, Field)
    End Function

    Friend Overridable Function GetSO001() As String
        Return String.Format("Select A.RowID,A.* From SO001 A WHERE A.Custid={0}0 ", Sign)
    End Function

    Friend Overridable Function GetSO002(ByVal ServiceType As String) As String
        Dim strwhere As String = String.Empty
        If Not String.IsNullOrEmpty(ServiceType) Then
            strwhere = String.Format("and A.ServiceType='{0}'", ServiceType)
        End If
        Return String.Format("Select A.RowID,A.* From SO002 A WHERE A.Custid={0}0 {1}", Sign, strwhere)
    End Function

    Friend Function GetCustRtnCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD096 Where StopFlag = 0 And (ServiceType is null or  ServiceType ={0}0) Order by CodeNo", Sign)
        'Return String.Format("Select CodeNo,Description,RefNo From CD072 Where StopFlag = 0 And (ServiceType is null or  instr(ServiceType, {0}0) > 0) Order by CodeNo", Sign)
    End Function

    Friend Function GetPRInterdepend() As String
        Return String.Format("Select * From SO009 WHERE Custid= {0}0 and SNO={0}1 ", Sign)
    End Function

    Friend Function chkPrChangeFacility() As String
        Return String.Format("Select Count(*) From SO004D Where SeqNo ={0}0 And Kind ='{1}'", Sign, SaveLanguage.SO004DKind)
    End Function

    Friend Function chkWipPRMainSNO() As String
        Return String.Format("Select Count(*) From SO009 Where MainSNo = {0}0 And CustId = {0}1 And ServiceType = {0}2 And ReturnCode is not null", Sign)
    End Function

    Friend Function chkReInstAcross(ByVal ReInstOwner As String) As String
        Return String.Format("Select NSNOStatus From {1}SO313 Where OCustId = {0}0 And OSNo = {0}1 And OCompCode = {0}2", Sign, ReInstOwner)
    End Function

    Friend Function GetSO041() As String
        Return "Select * From SO041"
    End Function

    Friend Function chkWipPRFinTime() As String
        Return String.Format("Select Count(*) From SO009 Where MainSNo = {0}0 And CustId = {0}1 And ServiceType = {0}2 And FinTime is not null", Sign)
    End Function

    Friend Function GetAddressData() As String
        Dim strSQL As String = ""
        Dim strField As String = "AddrNo,Address,StrtCode,StrtName,ServCode,ServName," &
                                "AreaCode,AreaName,ClctAreaCode,ClctAreaName," &
                                "MDUId,MDUName,NodeNo,CircuitNo,SalesCode,SalesName"
        strSQL = String.Format("Select {1} From SO014 A Where Exists (Select 1 From SO001 B Where A.AddrNo = B.InstAddrNo And A.CompCode = B.CompCode And B.CustId = {0}0 And B.CompCode = {0}1)", Sign, strField)
        Return strSQL
    End Function

    Friend Function GetFaciSeqNoData() As String
        Return String.Format("Select * From SO004 Where Custid={0}0 and SeqNo={0}1 ", Sign)
    End Function
    Friend Function getServiceIdByCitemCode() As String
        Return String.Format("Select ServiceId from SO003C Where 1=1 and ((PrDate Is Null and InstDate Is Not Null) Or (InstDate > PrDate)) " & _
                              "And ProductCode IN (Select ProductCode From CD019 Where CodeNo = {0}0) " & _
                              " And Custid = {0}1 And FaciSeqNo = {0}2", Sign)
    End Function
    Friend Function Get003CData() As String
        Return String.Format("Select ServiceId From SO003C Where ServiceId Is Not Null and Custid={0}0 and FaciSeqNo={0}1 and ((PrDate Is Null and InstDate Is Not Null) Or (InstDate > PrDate))", Sign)
    End Function

    Friend Function GetCanMoveServiceType(ByVal strCanChooseFaciRefno As String) As String
        '#8481 add condition to filter that bill has been generated by kin 2019/08/21
        '沒有拆除日，沒有取回日，沒有拆除單號，有安裝日,若該設備雖然有拆除單號，但拆除單號的設備指定若為更換狀態，仍需產生移機單
        Return String.Format("Select A.ServiceType,0 PRCODE,'' PRName,0 REASONCODE,'' REASONNAME From SO004 A ,CD022 B " &
                             "Where A.CustId={0}0 and A.PRDate Is Null And A.GetDAte Is Null and A.ServiceType<>{0}1 " &
                             " AND A.INSTDATE IS NOT NULL " &
                             " AND (A.PRSNO IS NULL OR EXISTS (SELECT 1 FROM SO004D C WHERE C.KIND='更換' AND C.SNO=A.PRSNO AND A.CUSTID=C.CUSTID)) " &
                             "and A.ServiceType In (Select ServiceType From SO002 Where Custid={0}0 And CustStatusCode In (1,2))" &
                             "and A.FaciCode = B.CodeNo and Nvl(B.StopFlag,0)=0 and B.RefNo in ({1}) " &
                             "Group by A.ServiceType Order by A.ServiceType ", Sign, strCanChooseFaciRefno)
    End Function

    Friend Function GetChangeFacilityPinCode(ByVal InSeqNo As String) As String
        Return String.Format("Select A.* From SO004 A ,CD022 B Where A.CustId={0}0 and SeqNo in ({1}) and PinCode Is Not Null and A.PRDate Is Null And A.GetDAte Is Null and A.FaciCode = B.CodeNo and Nvl(B.StopFlag,0)=0 and B.RefNo in (2,3,5,6,7,8,10)", Sign, InSeqNo)
    End Function

    Friend Function GetDeclarantData() As String
        Return String.Format("Select * From SO137 Where ID={0}0", Sign)
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
