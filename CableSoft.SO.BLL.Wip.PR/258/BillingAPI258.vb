Imports System.Data.Common
Imports CableSoft.Utility.DataAccess
Imports CableSoft.BLL.Utility
'新增拆機工單（不進結清）－258
'	使用時機：新增拆機或暫停頻道工單資料於營運系統。
'	使用限制：不進行費用結清作業，直接產生停拆機工單。
Public Class BillingAPI258
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    '    Private _DAL As New BillingAPI258DALMultiDB(Me.LoginInfo.Provider)
    Private _DAL As New BillingAPI258DAL(Me.LoginInfo.Provider)
    Private Lang As New CableSoft.BLL.Language.SO61.BillingAPI258Language

    '回應碼	回應狀態	回應訊息
    ' 0     成功
    '-1	    失敗	{參數}資料有誤!!
    '-304	失敗	客編不存在
    '-163	失敗	停拆機原因不存在
    '-111	失敗	該 {機上盒、智慧卡…} 序號:{0}不存在,請查證!!
    '-118	失敗	設備已拆或拆機中
    '-119	失敗	該設備安裝中
    '-157	失敗	受理人員不存在
    '-171	失敗	停拆機原因細項不存在(拆機原因代碼不存在)
    '-104	失敗	互動單號不存在

    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        'TAG	筆數	參數	    名稱	        必要	說明
        'Main	單	    APIID	    命令識別碼	    V	    258
        'Main	單	    Compcode	公司別	        V	    預設公司別
        'Main	單	    Caller	    呼叫來源	    V	    自定名稱 Ex. IVR,CSR,WEB…
        'Main	單	    Seqno	    來源識別碼	    V	    自定編碼，編碼規則：17碼，YYYYMMDDHHMMSS+3碼流水號
        'SNO	單	    CustId	    客戶編號	    V	
        'SNO	單	    ReasonCode	停拆機原因代碼	V	
        'SNO	單	    ResvTime	預約時間	    V	    YYYY/MM/DD HH24:MI:SS 
        'SNO	單	    AcceptEn	受理人員代號	V	
        'SNO	單	    Note	    備註		
        'SNO	單	    ReasonDescCode	停拆機原因細項		
        'SNO	單	    CallSeqNo	互動單號		
        'SNO	單	    Faciseqno	設備流水號		V       
        'SNO	單	    PRCode	    派工類別	    V	    '#7666 需求增加指定派工類別
        Dim result As RIAResult = Nothing
        Dim ServiceType As String = Nothing
        Dim PRCode As Int32 = 0
        Dim WipRefNo As Int32 = 0
        Dim ReInstAcrossFlag As Boolean = False
        Dim strRetSNO As String = String.Empty '紀錄產生工單號碼
        '#8706
        Me.LoginInfo.EntryId = InData.Tables("SNo").Rows(0).Item("AcceptEn")
        Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {InData.Tables("SNo").Rows(0).Item("AcceptEn")})
            Me.LoginInfo.EntryId = InData.Tables("SNo").Rows(0).Item("AcceptEn")
            Me.LoginInfo.EntryName = EmpName.Rows(0).Item("EmpName")
        End Using
        '檢核是否可派工
        result = CheckCanPR(InData, ServiceType, WipRefNo, ReInstAcrossFlag, PRCode, "")
        If result.ResultBoolean = False Then
            Return result
        End If
        result = WipDataSave(InData, ServiceType, WipRefNo, ReInstAcrossFlag, PRCode, "")
        If result.ResultBoolean Then
            strRetSNO = strRetSNO & "," & result.ResultXML
        Else
            Return result
        End If
        
        '回傳資料
        If Not String.IsNullOrEmpty(strRetSNO) Then
            If strRetSNO.Substring(0, 1) = "," Then strRetSNO = strRetSNO.Substring(1)
        End If
        result.ResultDataSet = GetReturnData(strRetSNO)
        Return result
    End Function

    Private Function WipDataSave(InData As DataSet, ByVal ServiceType As String,
                                 ByVal WipRefNo As Int32, ByVal ReInstAcrossFlag As Boolean,
                                 ByVal PRCode As Int32, ByVal CanChooseCitemCode As String) As RIAResult
        Dim result As RIAResult = Nothing
        Dim WipData As DataSet = Nothing
        Dim ResvTime As DateTime = DateTime.Parse(InData.Tables("SNo").Rows(0).Item("ResvTime"))

        Using PR As New PR(LoginInfo, DAO)
            WipData = PR.GetPRData(Nothing, InData.Tables("Sno").Rows(0).Item("CustId"), ServiceType)
            result = GetWipData(InData, WipData, ServiceType, WipRefNo, ReInstAcrossFlag, PRCode, CanChooseCitemCode)
            If Not result.ResultBoolean Then
                Throw New Exception("GetWipData")
            End If
            '檢核預約時間是否可以改約
            result = ChkCanResv(WipData, ResvTime, ServiceType)
            If result.ResultBoolean = False Then
                result.ErrorCode = -155
                Return result
            End If
        End Using
        WipData.Tables("Wip").Rows(0).Item("ResvTime") = ResvTime

        '檢核移拆工單是否可存檔
        Using Vali As New Validate(LoginInfo, DAO)
            result = Vali.ChkDataOk(EditMode.Append, WipData)
            If result.ResultBoolean = False Then
                result = ChangeErrorCode(result)
                Return result
            End If
        End Using
        '工單存檔
        'API 測試報告,SO004D.Chooseserviceid 沒有填 補填進去By Kin 2019/11/18 For Debby
        Using objBll As New PR(Me.LoginInfo, DAO)
            Dim strChooseserviceid As String = objBll.setSO004DChooseserviceid(WipData.Tables("Wip").Rows(0).Item("CustId"), _
                    InData.Tables("SNO").Rows(0).Item("FACISEQNO").ToString)
            If Not String.IsNullOrEmpty(strChooseserviceid) Then
                For i As Integer = 0 To WipData.Tables("ChangeFacility").Rows.Count - 1
                    WipData.Tables("ChangeFacility").Rows(i).Item("chooseserviceid") = strChooseserviceid
                    WipData.Tables("ChangeFacility").AcceptChanges()
                Next
            End If
            Using bll As New SaveData(LoginInfo, DAO)
                result = bll.Save(EditMode.Append, False, WipData, False)
                If result.ResultBoolean = False Then
                    result = ChangeErrorCode(result)
                    Return result
                End If
            End Using
        End Using
       
        WipData.Dispose()
        Return New RIAResult With {.ResultBoolean = True, .ResultXML = WipData.Tables("Wip").Rows(0)("SNO")}
    End Function

    Private Function CheckCanPR(InWipData As DataSet, ByRef ServiceType As String,
                                ByRef WipRefNo As Int32, ByVal ReInstAcrossFlag As Boolean,
                                ByRef intPRCode As Integer, ByRef CitemCode As String) As RIAResult
        Using Vali As New Validate(LoginInfo, DAO)
            Dim InRow As DataRow = InWipData.Tables("SNo").Rows(0)
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                If Not String.IsNullOrEmpty(InRow("ReasonDescCode").ToString) Then
                    Using Code As DataTable = DAO.ExecQry(_DAL.GetReasonDescName(), New Object() {InRow("ReasonDescCode")})
                        If Code.Rows.Count = 0 Then
                            Return New RIAResult With {.ErrorCode = -171, .ErrorMessage = Lang.noPRDetailReasno}
                        End If
                    End Using
                End If
                Dim inPRCode As Integer = 0 '#7666 2018.02.12 by Corey 增加判斷是否有傳入PRCODE，才針對REFNO=7的派工類別代碼?
                If InWipData.Tables("SNO").Columns.Contains("PRCODE") Then
                    If Not InRow.IsNull("PRCode") Then inPRCode = InRow("PRCode")
                End If
                Using dtCD007 As DataTable = DAO.ExecQry(_DAL.GetWorkCodeForCitemCode(inPRCode))
                    If dtCD007.Rows.Count > 0 Then
                        intPRCode = dtCD007.Rows(0)("CodeNo")
                        ServiceType = dtCD007.Rows(0)("ServiceType")
                    Else
                        Return New RIAResult With {.ErrorCode = -159, .ErrorMessage = Lang.NoPRClass}
                    End If
                End Using
                If Not String.IsNullOrEmpty(InRow("FaciSeqno").ToString) Then
                    Dim aryFaciSeqno() As String = InRow("FaciSeqno").ToString.Split(",")
                    For Each strSEQNO As String In aryFaciSeqno
                        Dim dtPROther As DataTable = DAO.ExecQry(_DAL.CheckPrDouble, New Object() {InRow("Custid"), ServiceType, intPRCode})
                        If dtPROther.Rows.Count > 0 Then
                            For Each drPROther As DataRow In dtPROther.Rows
                                Dim dtPrChangeFaci As DataTable = DAO.ExecQry(_DAL.CheckPrChangeFacility, New Object() {InRow("Custid"), drPROther("SNO")})
                                For Each drPrChangeFaci As DataRow In dtPrChangeFaci.Rows
                                    If drPrChangeFaci("SEQNO") = strSEQNO Then
                                        Return New RIAResult With {.ErrorCode = -999, .ErrorMessage = Lang.OrderExists}
                                    End If
                                Next
                            Next
                        End If
                    Next
                End If
                Using PRCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, intPRCode.ToString, False)
                    ServiceType = PRCode.Rows(0).Item("ServiceType")
                    WipRefNo = Int32.Parse("0" & PRCode.Rows(0).Item("RefNo").ToString)
                    ReInstAcrossFlag = IIf(Int32.Parse("0" & PRCode.Rows(0).Item("ReInstAcrossFlag").ToString) = 0, False, True)
                    Using Customer As DataTable = DAO.ExecQry(_DAL.GetCustomerData(), New Object() {InRow.Item("CustId"), ServiceType})
                        Return Vali.CheckCanPR(PRCode.Rows(0).Item("CodeNo"), CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(PRCode.Rows(0).Item("RefNo")), CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(PRCode.Rows(0).Item("Interdepend")), Customer.Rows(0).Item("CustStatusCode"), Customer.Rows(0).Item("WipCode3"), InRow.Item("CustId"), ServiceType, Customer.Rows(0).Item("InstAddrNo"), Nothing)
                    End Using
                End Using
            End Using
        End Using
    End Function

    Private Function ChkCanResv(WipData As System.Data.DataSet, ResvTime As DateTime, ByVal ServiceType As String) As RIAResult
        Using bll As New Validate(LoginInfo, DAO)
            Dim result As RIAResult = Nothing
            '檢核預約時間是否可以改約
            Dim WipRow As DataRow = WipData.Tables("Wip").Rows(0)
            Using WorkCode As DataTable = DAO.ExecQry(_DAL.GetWorkCode(), New Object() {WipRow.Item("PRCode")})
                Dim MCode As Integer = CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(WorkCode.Rows(0).Item("GroupNo"))
                Dim Resvdatebefore As Integer = CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(WorkCode.Rows(0).Item("Resvdatebefore"))
                Dim WorkUnit As Decimal = CableSoft.BLL.Utility.Utility.ConvertDBNullToDecimal(WorkCode.Rows(0).Item("WorkUnit"))
                '2016.09.20 傳參順序錯誤調正
                'result = bll.ChkCanResv(WipRow.Item("PRCode"), WipRow.Item("WorkServCode"), MCode, ServiceType, ResvTime, WipRow.Item("AcceptTime"), WipRow.Item("ResvTime"), Resvdatebefore, WorkUnit, True)
                result = bll.ChkCanResv(WipRow.Item("WorkServCode"), WipRow.Item("PRCode"), MCode, ServiceType, ResvTime, WipRow.Item("AcceptTime"), WipRow.Item("ResvTime"), Resvdatebefore, WorkUnit, True)
                If result.ResultBoolean = False Then
                    Return result
                End If
            End Using
            Return New RIAResult With {.ResultBoolean = True}
        End Using
    End Function

    Private Function ChangeErrorCode(ByRef Changeresult As RIAResult) As RIAResult
        Select Case Changeresult.ErrorCode
            Case -11005
                Changeresult.ErrorCode = -173
                Changeresult.ErrorMessage = Lang.FullPoint
        End Select
        Return Changeresult
    End Function

    Private Function GetContactData(InWipData As DataSet, ByRef WipData As DataSet) As Boolean
        If String.IsNullOrEmpty(InWipData.Tables("SNo").Rows(0)("CallSeqNo").ToString) = False AndAlso String.IsNullOrEmpty(InWipData.Tables("SNo").Rows(0).Item("CallSeqNo")) = False Then
            Dim Contact As DataTable = DAO.ExecQry(_DAL.GetContactDetailData(), New Object() {InWipData.Tables("SNo").Rows(0).Item("CallSeqNo")})
            If Contact.Rows.Count = 0 Then
                Contact = DAO.ExecQry(_DAL.GetContactData(), New Object() {InWipData.Tables("SNo").Rows(0).Item("CallSeqNo")})
            End If
            Contact.TableName = "Contact"
            WipData.Tables.Add(Contact.Copy())
            Contact.Dispose()
        End If
        Return True
    End Function

    Private Function GetWipData(InWipData As DataSet, ByRef WipData As DataSet,
                                ByVal ServiceType As String, ByVal WipRefNo As Int32,
                                ByVal ReInstAcrossFlag As Boolean, ByVal PRCode As Int32,
                                ByVal CanChooseCitemCode As String) As RIAResult
        '更新互動資料
        If Not GetContactData(InWipData, WipData) Then
            Throw New Exception("GetWipData")
        End If
        '取得工單相關資料
        Using PR As New PR(LoginInfo, DAO)
            Dim InDataRow As DataRow = InWipData.Tables("Sno").Rows(0)
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                Using dtPRCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, PRCode.ToString(), False)
                    Dim SNo As String = SOUtil.GetFalseSNo(BLL.Utility.InvoiceType.PR, ServiceType)
                    Dim Contact As DataTable = Nothing
                    If WipData.Tables.Contains("Contact") Then
                        Contact = WipData.Tables("Contact")
                    End If
                    Using RetWip As DataSet = PR.GetNormalCalculateData(InDataRow.Item("CustId"), ServiceType, PRCode, InDataRow.Item("ResvTime"), SNo, True, WipData)
                        For Each RetTable As DataTable In RetWip.Tables
                            If WipData.Tables.Contains(RetTable.TableName) Then
                                WipData.Tables.Remove(RetTable.TableName)
                            End If
                            WipData.Tables.Add(RetTable.Copy())
                        Next
                    End Using

                    For Each tFaciSeqNo As String In InDataRow.Item("FaciSeqNo").ToString.Split(",")
                        Dim FaciSeqNo As String = tFaciSeqNo
                        '將指定後的設備相關資料填寫到Wipdata內
                        If WipData.Tables("ChangeFacility").AsEnumerable.Where(Function(list) list.IsNull("SeqNo") = False AndAlso list.Item("SeqNo") = FaciSeqNo).Count = 0 Then
                            ''指定異動設備
                            Using bll As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, Me.DAO)
                                Dim dtDataTable As DataTable = bll.GetCanChangeKind(SO.BLL.Utility.InvoiceType.PR, WipRefNo, ReInstAcrossFlag)
                                If dtDataTable.Rows.Count > 0 Then
                                    Dim strFaciRefNo As String = String.Empty
                                    Select Case ServiceType
                                        Case "I"
                                            strFaciRefNo = "2,5,7,8"
                                        Case "D"
                                            strFaciRefNo = "3"
                                        Case "P"
                                            strFaciRefNo = "6"
                                    End Select
                                    Dim FacilityData As DataTable = DAO.ExecQry(_DAL.FacilityData(strFaciRefNo), New Object() {WipData.Tables("Wip").Rows(0)("Custid"), ServiceType, FaciSeqNo})
                                    If FacilityData.Rows.Count > 0 Then
                                        If Not String.IsNullOrEmpty(CanChooseCitemCode) Then
                                            If CanChooseCitemCode.Substring(0, 1) = "," Then CanChooseCitemCode = CanChooseCitemCode.Substring(1)
                                        End If
                                        If Not bll.GetChangeFacility(dtDataTable.Rows(0)(1), WipData.Tables("Wip").Rows(0), Nothing, FacilityData.Rows(0), CanChooseCitemCode, WipData.Tables("ChangeFacility")) Then
                                            Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -900, .ErrorMessage = Lang.GetChangeFacilityErr}
                                        End If
                                    End If
                                End If
                            End Using
                        End If
                    Next
                End Using
            End Using
        End Using
        '異動拆機工單相關欄位
        If Not UpdateWipPRHead(InWipData, WipData) Then
            Throw New Exception("UpdateWipHead")
        End If

        Return New RIAResult With {.ResultBoolean = True}
    End Function

    Private Function UpdateWipPRHead(InWipData As DataSet, ByRef WipData As DataSet) As Boolean
        Dim PRRow As DataRow = WipData.Tables("Wip").Rows(0)
        Dim InDataRow As DataRow = InWipData.Tables("Sno").Rows(0)
        With PRRow
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                .Item("ReasonCode") = InDataRow.Item("ReasonCode")
                Using Code As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.ReasonCode, .Item("ReasonCode").ToString(), True)
                    .Item("ReasonName") = Code.Rows(0).Item("Description")
                End Using
                If Not String.IsNullOrEmpty(InDataRow("ReasonDescCode").ToString) Then
                    .Item("ReasonDescCode") = InDataRow.Item("ReasonDescCode")
                    Using Code As DataTable = DAO.ExecQry(_DAL.GetReasonDescName(), New Object() {.Item("ReasonDescCode")})
                        .Item("ReasonDescName") = Code.Rows(0).Item("Description")
                    End Using
                End If
            End Using
            .Item("AcceptEn") = InDataRow.Item("AcceptEn")
            Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {.Item("AcceptEn")})
                .Item("AcceptName") = EmpName.Rows(0).Item("EmpName")
            End Using
            .Item("AcceptTime") = DateTime.Now
            '.Item("ResvFlagTime") = InDataRow.Item("ResvFlagTime")
            .Item("ModifyFlag") = 1
            .Item("PrintBillFlag") = 0
            .Item("Note") = InDataRow.Item("Note")
            .Item("UpdEn") = LoginInfo.EntryName
            .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(DateTime.Now)
            .Item("NewUpdTime") = DateTime.Now
        End With
        WipData.Tables("Wip").Rows(0).AcceptChanges()
        Return True
    End Function

    Private Function GetReturnData(SNO As String) As DataSet
        Dim RetData As New DataSet With {.DataSetName = "DataSet"}
        Dim RetTable As New DataTable With {.TableName = "SNo"}
        RetTable.Columns.Add(New DataColumn With {.ColumnName = "SNo", .DataType = GetType(String)})
        RetTable.Rows.Add(RetTable.NewRow())
        RetTable.Rows(0).Item("SNo") = SNO
        RetData.Tables.Add(RetTable)
        Return RetData
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean
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
    ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If
            Try
                If _DAL IsNot Nothing Then
                    _DAL.Dispose()
                End If
                If MyBase.MustDispose AndAlso DAO IsNot Nothing Then
                    DAO.Dispose()
                End If
                If Lang IsNot Nothing Then
                    Lang.Dispose()
                    Lang = Nothing
                End If
            Catch ex As Exception
            End Try
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
'填入設定檔資料
'Insert into TBCSH.SO1114A
'   (APIID, APINAME, COMMENTS, DLLNAME, CLASSNAME)
' Values
'   ('258', '新增CM軟關工單（不進結清）', '使用時機：新增CM軟關工單資料於營運系統。
'使用限制：不進行費用結清作業，直接產生停拆機工單。', 'CableSoft.SO.BLL.Wip.PR.dll', 'CableSoft.SO.BLL.Wip.PR.BillingAPI258');


'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'APIID', '命令識別碼', 0, 1, 
''Main', -1, 1);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'Compcode', '公司別', 1, 1, 
''Main', -1, 2);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'Caller', '呼叫來源', 0, 1, 
''Main', -1, 3);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'Seqno', '來源識別碼', 0, 1, 
''Main', -1, 4);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'CustId', '客戶編號', 0, 1, 
''SNo', -1, 5);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'ReasonCode', '停拆機原因代碼', 0, 1, 
''SNo', -1, 6);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'ResvTime', '預約時間', 3, 1, 
''SNo', -1, 7);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'AcceptEn', '受理人員代號', 0, 1, 
''SNo', -1, 8);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'Note', '備註', 0, 0, 
''SNo', -1, 9);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'ReasonDescCode', '停拆機原因細項', 0, 0, 
''SNo', -1, 10);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'CallSeqNo', '互動單號', 0, 0, 
''SNo', -1, 11);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, 
'    DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('258', 'FaciSeqNo', '設備流水號', 0, 1, 
''SNo', -1, 12);



'================================================================================================================================
'SO1114B.CustId 預約時間 ChkSQLQuery設定內容
'Select -304 ErrorCode,'客編不存在' ErrorMsg From Dual
'   Where (Select Count(*) From SO001 Where CustId = '[CustId]') = 0
'Union All
'Select -163 ErrorCode,'停拆機原因不存在' ErrorMsg From Dual
'   Where (Select Count(*) From CD014 Where CodeNo = '[ReasonCode]' And StopFlag = 0 and ServiceType='I') = 0
'Union All
'Select -157 ErrorCode,'受理人員不存在' ErrorMsg From Dual
'   Where (Select Count(*) From CM003 Where EmpNo = '[AcceptEn]' And StopFlag = 0) = 0
'Union All
'Select -169 ErrorCode,'客戶狀態為註銷不允許新增' ErrorMsg From SO002 A
'   Where A.CustId = [CustId] And A.CustStatusCode = 4 
'Union All
'Select -111 ErrorCode,'該 {機上盒、智慧卡…} 序號:{[Faciseqno]}不存在,請查證!!' ErrorMsg 
'   From Dual
'   Where (Select Count(*) From SO004 Where SeqNo In ([Faciseqno])) = 0
'Union All
'Select -118 ErrorCode,'設備已拆或拆除中' ErrorMsg
'   From SO004
'   Where SeqNo = '[Faciseqno]' And PrDate Is Not Null
'Union All
'Select -118 ErrorCode,'設備已拆或拆除中' ErrorMsg
'   From SO004
'   Where SeqNo = '[Faciseqno]' And GetDate Is Not Null
'Union All
'Select -119 ErrorCode,'該設備安裝中' ErrorMsg
'   From SO004
'   Where SeqNo = '[FaciSeqNo]' And InstDate Is Null
'Union All
'Select -200 ErrorCode,'收費項目不存在' ErrorMsg
'   From CD019
'   Where CodeNO = '[CitemCode]' And CodeNo Is Null and ServiceType='I'

'--Union All
'--SELECT -104 ErrorCode,'互動單號不存在' ErrorMsg From Dual
'--   Where (Select Count(*) From 
'--   (Select 1 From SO006A Where CustId =[custid] And SeqNo = [CallSeqNo] and ProcResultNo is null 
'--     Union All 
'--    Select 1 From SO006 Where CustId = [custid] And SeqNo = [CallSeqNo] and ProcResultNo is null)) = 0 

'================================================================================================================================
'測試傳入參數
'{
' "Main": [{
'  "APIID": "258",
' "Compcode": "3",
'  "Caller": "API-Corey",
'  "Seqno": "201401150628114539" }],
' "SNO": [{
'  "CustId": "600040",
'  "ReasonCode": "204",
'  "ResvTime": "2016/09/01 15:00:00",
'  "AcceptEn": "1606941112",
'  "Note": "去時請電聯",
'  "ReasonDescCode": "",
'  "CallSeqNo": "",
'  "Faciseqno": "201207170216227" }]
'}

