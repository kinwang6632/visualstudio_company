Imports System.Data.Common
Imports CableSoft.Utility.DataAccess
Imports CableSoft.BLL.Utility
'4.43	新增拆機工單（無結清）－259
'	使用時機：拆頻道或拆機流程不進入結清機制，新增拆機工單資料於營運系統。
'	使用限制：
'   1.催收作業需求，整批派拆機單，不需透過來電與客戶互動，所以不需CallSeqNo(互動單號)。
'   2.拆CATV服務別時，因CATV無設備所以”設備流水號”填入該客編即可。
Public Class BillingAPI259
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private _DAL As New BillingAPI259DAL(Me.LoginInfo.Provider)
    'Private _DAL As New BillingAPI259DALMultiDB(Me.LoginInfo.Provider)
    Private Lang As CableSoft.BLL.Language.SO61.BillingAPI259Language

    '回應碼	回應狀態	回應訊息
    '0	    成功	
    '-1	    失敗	    {參數}資料有誤!!
    '-304	失敗	    客編不存在
    '-163	失敗	    停拆機原因不存在
    '-157	失敗	    受理人員不存在
    '-171	失敗	    停拆機原因細項不存在(拆機原因代碼不存在)
    '-104	失敗	    互動單號不存在
    '-159	失敗	    停拆機類別不存在
    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        '參數抬頭	筆數	參數	        名稱	        必要	說明
        'Main	    單	    APIID	        命令識別碼	    V	    259
        'Main	    單	    Compcode	    公司別	        V	    預設公司別
        'Main	    單	    Caller	        呼叫來源	    V	    自定名稱 Ex. IVR,CSR,WEB…
        'Main	    單	    Seqno	        來源識別碼	    V	    自定編碼，編碼規則：17碼，YYYYMMDDHHMMSS+3碼流水號
        'SNO	    單	    CustId	        客戶編號	    V	
        'SNO	    單	    PRCode	        停拆機類別代碼	V	
        'SNO	    單	    ReasonCode	    停拆機原因代碼	V	    Follow SO009.ReasonCode
        'SNO	    單	    ReasonDescCode	停拆機原因細項		
        'SNO	    單	    ResvTime	    預約時間	    V	    YYYY/MM/DD HH24:MI:SS 
        'SNO	    單	    Faciseqno	    設備流水號	    V	    若是CATV帶入該客編值即可。
        'SNO	    單	    Note	        備註		
        'SNO	    單	    GroupCode	    工程組別		
        'SNO	    單	    WorkerEn1	    工作人員1		
        'SNO	    單	    AcceptEn	    受理人員代號	V	
        'SNO	    單	    CallSeqNo	    互動單號		

        Dim result As RIAResult = Nothing
        Dim ServiceType As String = Nothing
        Dim strRetSNO As String = String.Empty '紀錄產生工單號碼
        '#8706
        Me.LoginInfo.EntryId = InData.Tables("SNo").Rows(0).Item("AcceptEn")
        Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {InData.Tables("SNo").Rows(0).Item("AcceptEn")})
            Me.LoginInfo.EntryId = InData.Tables("SNo").Rows(0).Item("AcceptEn")
            Me.LoginInfo.EntryName = EmpName.Rows(0).Item("EmpName")
        End Using
        '檢核是否可派工
        result = CheckCanPR(InData, ServiceType)
        If result.ResultBoolean = False Then
            Return result
        End If

        result = WipDataSave(InData, ServiceType)
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

    Private Function WipDataSave(InData As DataSet, ByVal ServiceType As String) As RIAResult
        Dim result As RIAResult = Nothing
        Dim WipData As DataSet = Nothing
        Dim ResvTime As DateTime = DateTime.Parse(InData.Tables("SNo").Rows(0).Item("ResvTime"))
        Using PR As New PR(LoginInfo, DAO)
            WipData = PR.GetPRData(Nothing, InData.Tables("Sno").Rows(0).Item("CustId"), ServiceType)
            result = GetWipData(InData, WipData, ServiceType)
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
        Using bll As New SaveData(LoginInfo, DAO)
            result = bll.Save(EditMode.Append, False, WipData, False)
            If result.ResultBoolean = False Then
                result = ChangeErrorCode(result)
                Return result
            End If
        End Using
        WipData.Dispose()
        Return New RIAResult With {.ResultBoolean = True, .ResultXML = WipData.Tables("Wip").Rows(0)("SNO")}
    End Function

    Private Function CheckCanPR(InWipData As DataSet, ByRef ServiceType As String) As RIAResult
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
                Using PRCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, InRow.Item("PRCode").ToString, False)
                    Dim strRefNo As String = PRCode.Rows(0).Item("RefNo").ToString
                    ServiceType = PRCode.Rows(0).Item("ServiceType")
                    If Not String.IsNullOrEmpty(strRefNo) Then
                        '2016.09.30 因為SO1114B檢核，已經限定 REFNO=2 and REFNO IS NULL 才可以執行該API，所以這邊只需要判斷REFNO是否有值。
                        '有值表示 該派工類別參考號為2
                        Using Customer As DataTable = DAO.ExecQry(_DAL.GetCustomerData(), New Object() {InRow.Item("CustId"), ServiceType})
                            If Int32.Parse("0" & Customer.Rows(0)("CustStatusCode").ToString) = 6 Then
                                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = Lang.OrderExists}
                            End If
                            Return Vali.CheckCanPR(PRCode.Rows(0).Item("CodeNo"), CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(PRCode.Rows(0).Item("RefNo")), CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(PRCode.Rows(0).Item("Interdepend")), Customer.Rows(0).Item("CustStatusCode"), Customer.Rows(0).Item("WipCode3"), InRow.Item("CustId"), ServiceType, Customer.Rows(0).Item("InstAddrNo"), Nothing)
                        End Using
                    Else
                        '2016.09.30 因為客戶要增加新的流程，CATV拆機 需要增加CATV軟關的工單，原因是要催收紀錄用的。所以完工退單都不應該影響客戶狀態。
                        '如果派工類別工單沒有參考號 則需要判斷該派工類別是否重複派工。
                        Using dtPROther As DataTable = DAO.ExecQry(_DAL.CheckPrDouble, New Object() {InRow("Custid"), ServiceType, InRow("PRCODE")})
                            If dtPROther.Rows.Count > 0 Then
                                Return New RIAResult With {.ErrorCode = -999, .ErrorMessage = Lang.OrderExists}
                            End If
                        End Using
                        Return New RIAResult With {.ResultBoolean = True}
                    End If
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
                                ByVal ServiceType As String) As RIAResult
        '更新互動資料
        If Not GetContactData(InWipData, WipData) Then
            Throw New Exception("GetWipData")
        End If
        '取得工單相關資料
        Using PR As New PR(LoginInfo, DAO)
            Dim InDataRow As DataRow = InWipData.Tables("Sno").Rows(0)
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                Using PRCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, InDataRow.Item("PRCode").ToString(), False)
                    Dim SNo As String = SOUtil.GetFalseSNo(BLL.Utility.InvoiceType.PR, ServiceType)
                    Dim Contact As DataTable = Nothing
                    If WipData.Tables.Contains("Contact") Then
                        Contact = WipData.Tables("Contact")
                    End If
                    Using RetWip As DataSet = PR.GetNormalCalculateData(InDataRow.Item("CustId"), ServiceType, InDataRow.Item("PRCode"), InDataRow.Item("ResvTime"), SNo, True, WipData)
                        For Each RetTable As DataTable In RetWip.Tables
                            If WipData.Tables.Contains(RetTable.TableName) Then
                                WipData.Tables.Remove(RetTable.TableName)
                            End If
                            WipData.Tables.Add(RetTable.Copy())
                        Next
                    End Using
                    If InWipData.Tables("SNO").Columns.Contains("FaciSeqNo") Then
                        For Each tFaciSeqNo As String In InDataRow.Item("FaciSeqNo").ToString.Split(",")
                            Dim FaciSeqNo As String = tFaciSeqNo
                            '將指定後的設備相關資料填寫到Wipdata內
                            If WipData.Tables("ChangeFacility").AsEnumerable.Where(Function(list) list.IsNull("SeqNo") = False AndAlso list.Item("SeqNo") = FaciSeqNo).Count = 0 Then
                                '指定異動設備
                                Using bll As New CableSoft.SO.BLL.Facility.ChangeFaci.ChangeFaci(LoginInfo, DAO)
                                    '2018.04.23 by Corey 原本API-259是抄257的，原本是GetMovePRFaci "移拆"
                                    '           PM_Jammy告知 #7685 要的是 "拆除"，所以改呼叫 GetPRFaci
                                    Using RetData As DataSet = bll.GetPRFaci(SNo, FaciSeqNo)
                                        For Each Table As String In New String() {"Facility", "PRFacility", "ChangeFacility"}
                                            For Each Row As DataRow In RetData.Tables(Table).Rows
                                                WipData.Tables(Table).Rows.Add(CableSoft.BLL.Utility.Utility.CopyDataRow(Row, WipData.Tables(Table).NewRow()))
                                            Next
                                        Next
                                    End Using
                                End Using
                            End If
                        Next
                    End If
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
                .Item("ReasonCode") = InDataRow.Item("ReasonCode") '停拆機原因代碼(必要)
                Using Code As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.ReasonCode, .Item("ReasonCode").ToString(), True)
                    .Item("ReasonName") = Code.Rows(0).Item("Description")
                End Using
                If Not String.IsNullOrEmpty(InDataRow("ReasonDescCode").ToString) Then
                    .Item("ReasonDescCode") = InDataRow.Item("ReasonDescCode") '停拆機原因細項
                    Using Code As DataTable = DAO.ExecQry(_DAL.GetReasonDescName(), New Object() {.Item("ReasonDescCode")})
                        .Item("ReasonDescName") = Code.Rows(0).Item("Description")
                    End Using
                End If
            End Using
            .Item("AcceptEn") = InDataRow.Item("AcceptEn") '受理人員代號(必要)
            Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {.Item("AcceptEn")})
                .Item("AcceptName") = EmpName.Rows(0).Item("EmpName")
            End Using
            If Not String.IsNullOrEmpty(InDataRow("Note").ToString) Then
                .Item("Note") = InDataRow.Item("Note") '備註
            End If
            If Not String.IsNullOrEmpty(InDataRow("GroupCode").ToString) Then
                .Item("GroupCode") = InDataRow.Item("GroupCode") '工程組別
                Using Code As DataTable = DAO.ExecQry(_DAL.GetGroupCode(), New Object() {.Item("GroupCode")})
                    .Item("GroupName") = Code.Rows(0).Item("Description")
                End Using
            End If
            If Not String.IsNullOrEmpty(InDataRow("WorkerEn1").ToString) Then
                .Item("WorkerEn1") = InDataRow.Item("WorkerEn1") '工作人員1
                Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {.Item("WorkerEn1")})
                    .Item("WorkerName1") = EmpName.Rows(0).Item("EmpName")
                End Using
            End If
            .Item("AcceptTime") = DateTime.Now
            '.Item("ResvFlagTime") = InDataRow.Item("ResvFlagTime")
            .Item("ModifyFlag") = 1
            .Item("PrintBillFlag") = 0
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
