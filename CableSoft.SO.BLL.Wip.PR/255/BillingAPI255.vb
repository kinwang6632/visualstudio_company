Imports System.Data.Common
Imports CableSoft.Utility.DataAccess
Imports CableSoft.BLL.Utility
'新增同區移機工單－255
'	使用時機：產生同區移機工單於營運系統。
'	使用限制：CM,DTV同區移機，需先新增新址客編資料。
Public Class BillingAPI255
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private _DAL As New BillingAPI255DAL(Me.LoginInfo.Provider)
    'Private _DAL As New BillingAPI255DALMultiDB(Me.LoginInfo.Provider)
    Private ServiceType As String = Nothing
    Private Lang As New CableSoft.BLL.Language.SO61.BillingAPI255Language
    Private IntroName As String = Nothing
    '回應碼	回應狀態	回應訊息
    '0	    成功	
    '-1	    失敗	    {參數}資料有誤!!
    '-304	失敗	    客編不存在
    '-159	失敗	    停拆機類別不存在
    '-163	失敗	    停拆機原因不存在
    '-157	失敗	    受理人員不存在
    '-171	失敗	    停拆機原因細項不存在
    '-104	失敗	    互動單號不存在
    '-172	失敗	    新地址客編不存在
    '-173	失敗	    件數已滿不允許派工

    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim result As RIAResult = Nothing
        Dim ResvTime As DateTime = DateTime.Parse(InData.Tables("SNo").Rows(0).Item("ResvTime"))

        'SNO	單	CustId	        客戶編號            V	
        'SNO	單	PRCode	        停拆機類別代碼      V	取停拆移機代碼之參考號碼為2,5,6且CD007.ReInstAcrossFlag=2
        'SNO	單	ReasonCode	    停拆機原因代碼	    V	
        'SNO	單	ResvTime	    預約時間	        V	YYYY/MM/DD HH24:MI:SS
        'SNO	單	AcceptEn	    受理人員代號	    V	
        'SNO	單	Note	        備註		
        'SNO	單	ReasonDescCode	停拆機原因細項		
        'SNO	單	CallSeqNo	    互動單號	        V	
        'SNO	單	NewCustId	    新址客編	        V	
        'SNO	單	FaciSeqNo	    設備流水號	        V	多筆用逗號隔開
        'SNO	單	WorkServCode    實際派工服務區代碼  V '#7899 增加
        Dim WipData As DataSet = Nothing
        Dim InstWipData As DataSet = Nothing
        Dim MediaCode As String = Nothing
        Dim IntroId As String = Nothing
        '#8706
        Me.LoginInfo.EntryId = InData.Tables("SNo").Rows(0).Item("AcceptEn")
        Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {InData.Tables("SNo").Rows(0).Item("AcceptEn")})
            Me.LoginInfo.EntryId = InData.Tables("SNo").Rows(0).Item("AcceptEn")
            Me.LoginInfo.EntryName = EmpName.Rows(0).Item("EmpName")
        End Using
        '檢核是否可派工
        result = CheckCanPR(InData)        
        If result.ResultBoolean = False Then
            Return result
        End If
        '#8715
        If Not DBNull.Value.Equals(InData.Tables("SNo").Rows(0).Item("MediaCode")) Then
            MediaCode = InData.Tables("SNo").Rows(0).Item("MediaCode").ToString
        End If
        If Not DBNull.Value.Equals(InData.Tables("SNo").Rows(0).Item("IntroId")) Then
            IntroId = InData.Tables("SNo").Rows(0).Item("IntroId").ToString
        End If
        '#8767 change to sql to check by kin 2021/06/15
        'result = checkExistisIntroid(MediaCode, IntroId)
        'If result.ResultBoolean = False Then
        'Return result
        'End If
        Using PR As New PR(LoginInfo, DAO)
            WipData = PR.GetPRData(Nothing, InData.Tables("Sno").Rows(0).Item("CustId"), ServiceType)
            If Not GetWipData(InData, WipData, InstWipData) Then
                Throw New Exception("GetWipData")
            End If
            '檢核預約時間是否可以改約
            result = ChkCanResv(WipData, ResvTime)
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
        '檢核移入工單是否可存檔
        Using Vali As New CableSoft.SO.BLL.Wip.Install.Validate(LoginInfo, DAO)
            result = Vali.ChkDataOk(EditMode.Append, InstWipData)
            If result.ResultBoolean = False Then
                result = ChangeErrorCode(result)
                Return result
            End If
        End Using
        '工單存檔
        Using bll As New SaveData(LoginInfo, DAO)
            result = bll.Save(EditMode.Append, False, WipData, InstWipData, False)
            If result.ResultBoolean = False Then
                result = ChangeErrorCode(result)
                Return result
            End If
        End Using
        '回傳資料
        result.ResultDataSet = GetReturnData(WipData, InstWipData)
        WipData.Dispose()
        Return result
    End Function
    Private Function checkExistisIntroid(ByVal MediaCode As String, ByVal IntroId As String) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = False, .ErrorCode = -166, .ErrorMessage = Lang.notExistsIntroName}
        If String.IsNullOrEmpty(MediaCode) Then
            result.ErrorMessage = Nothing
            result.ErrorCode = 0
            result.ResultBoolean = True
            Return result
        End If
      
        Try
            Dim intMediaRefNo As Integer = Integer.Parse(DAO.ExecSclr(_DAL.getMediaRefNo, New Object() {MediaCode}))
            If intMediaRefNo = 0 Then
                If Not String.IsNullOrEmpty(IntroId) Then
                    result.ErrorCode = -1
                    result.ErrorMessage = String.Format(Lang.IntroidMustNull, MediaCode)
                    Return result
                End If
            Else
                If String.IsNullOrEmpty(IntroId) Then
                    result.ErrorCode = -1
                    result.ErrorMessage = String.Format(Lang.IntroidMust, MediaCode)
                    Return result
                End If
            End If
            Dim aIntroID As Object = IntroId
            If intMediaRefNo <> 2 AndAlso intMediaRefNo <> 3 Then
                If Not IsNumeric(IntroId) Then
                    aIntroID = -1
                End If
            End If
            Using o As New CableSoft.SO.BLL.Customer.IntroMedia.IntroMedia(Me.LoginInfo, Me.DAO)
                Using t As DataTable = o.keyCodeSearch(intMediaRefNo, aIntroID)
                    If t.Rows.Count > 0 Then
                        IntroName = t.Rows(0).Item("Description")
                        result.ErrorMessage = Nothing
                        result.ErrorCode = 0
                        result.ResultBoolean = True                    
                    End If
                End Using
            End Using
        Catch ex As Exception
            Throw ex
        End Try
        Return result
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
    Private Function GetWipData(InWipData As DataSet, ByRef WipData As DataSet, ByRef InstWipData As DataSet) As Boolean
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
                    '指定移出變更
                    Using bll As New CableSoft.SO.BLL.Facility.ChangeFaci.ChangeFaci(LoginInfo, DAO)
                        For Each tFaciSeqNo As String In InDataRow.Item("FaciSeqNo").ToString.Split(",")
                            Dim FaciSeqNo As String = tFaciSeqNo
                            If WipData.Tables("ChangeFacility").AsEnumerable.Where(Function(list) list.IsNull("SeqNo") = False AndAlso list.Item("SeqNo") = FaciSeqNo).Count = 0 Then
                                Using RetData As DataSet = bll.GetMovePRFaci(SNo, FaciSeqNo)
                                    For Each Table As String In New String() {"Facility", "PRFacility", "ChangeFacility"}
                                        For Each Row As DataRow In RetData.Tables(Table).Rows
                                            WipData.Tables(Table).Rows.Add(CableSoft.BLL.Utility.Utility.CopyDataRow(Row, WipData.Tables(Table).NewRow()))
                                        Next
                                    Next
                                End Using
                            End If
                        Next
                    End Using
                End Using
            End Using
        End Using
        '異動拆機工單相關欄位
        If Not UpdateWipPRHead(InWipData, WipData) Then
            Throw New Exception("UpdateWipHead")
        End If
        '新增移入單資料
        If Not GetInstWipData(InWipData, WipData, InstWipData) Then
            Throw New Exception("GetInstWipData")
        End If
        '異動裝機工單相關欄位
        If Not UpdateWipInstHead(InWipData, WipData, InstWipData) Then
            Throw New Exception("UpdateWipHead")
        End If
        Return True
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
            '#7899 2018.11.22 by Corey 需求增加前端呼叫增加欄位WorkServCode，並填寫在工單內WorkServCode。
            If InWipData.Tables("Sno").Columns.Contains("WorkServCode") Then
                If Not String.IsNullOrEmpty(InDataRow("WorkServCode").ToString) Then
                    .Item("WorkServCode") = InDataRow.Item("WorkServCode")
                End If
            End If
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
    Private Function GetInstWipData(InWipData As DataSet, WipData As DataSet, ByRef InstWipData As DataSet) As Boolean
        Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
            Using Inst As New CableSoft.SO.BLL.Wip.Install.Install(LoginInfo, DAO)
                Dim SNo As String = SOUtil.GetFalseSNo(BLL.Utility.InvoiceType.Install, ServiceType)
                Dim InDataRow As DataRow = InWipData.Tables("SNo").Rows(0)
                Dim InstCode As DataTable = Nothing
                Using InstCodes As DataTable = DAO.ExecQry(_DAL.GetReInstCode(), New Object() {ServiceType})
                    Using Valid As New CableSoft.SO.BLL.Wip.Install.Validate(LoginInfo, DAO)
                        For Each Row As DataRow In InstCodes.Rows
                            Dim result As RIAResult = Valid.CheckCanInstall(InDataRow.Item("NewCustId"), ServiceType, Row.Item("CodeNo"))
                            If result.ResultBoolean = True Then
                                InstCode = InstCodes.Clone
                                InstCode.Rows.Add(CableSoft.BLL.Utility.Utility.CopyDataRow(Row, InstCode.NewRow()))
                                Exit For
                            End If
                        Next
                    End Using
                End Using
                Dim RetData As DataSet = Inst.GetInstallData(SNo)
                Using TempData As DataSet = Inst.GetNormalCalculateData(InDataRow.Item("NewCustId"), ServiceType, InstCode.Rows(0).Item("CodeNo"), InDataRow.Item("ResvTime"), SNo, True, Nothing, WipData, WipData.Tables("Contact"))
                    For Each RetTable As DataTable In TempData.Tables
                        If RetData.Tables.Contains(RetTable.TableName) Then
                            RetData.Tables.Remove(RetTable.TableName)
                        End If
                        RetData.Tables.Add(RetTable.Copy())
                    Next
                End Using
                InstWipData = RetData
                '#8715
                For i As Integer = 0 To InstWipData.Tables("Wip").Rows.Count - 1
                    If Not DBNull.Value.Equals(InDataRow.Item("BulletinCode")) Then
                        InstWipData.Tables("Wip").Rows(i).Item("BulletinCode") = InDataRow.Item("BulletinCode")
                        InstWipData.Tables("Wip").Rows(i).Item("BulletinName") = DAO.ExecSclr(_DAL.QueryCD049Description, New Object() {InDataRow.Item("BulletinCode")})
                    End If
                    If Not DBNull.Value.Equals(InDataRow.Item("MediaCode")) Then
                        InstWipData.Tables("Wip").Rows(i).Item("MediaCode") = InDataRow.Item("MediaCode")
                        InstWipData.Tables("Wip").Rows(i).Item("MediaName") = DAO.ExecSclr(_DAL.QueryCD009Description, New Object() {InDataRow.Item("MediaCode")})
                    End If
                    If Not DBNull.Value.Equals(InDataRow.Item("IntroId")) Then
                        InstWipData.Tables("Wip").Rows(i).Item("IntroId") = InDataRow.Item("IntroId")
                        If Not String.IsNullOrEmpty(IntroName) Then
                            InstWipData.Tables("Wip").Rows(i).Item("IntroName") = IntroName
                        End If                                           
                    End If

                Next
                InstWipData.Tables("Wip").Rows(0).AcceptChanges()
            End Using
        End Using
        Return True
    End Function
    Private Function UpdateWipInstHead(InWipData As DataSet, WipData As DataSet, ByRef InstWipData As DataSet) As Boolean
        Dim InstRow As DataRow = InstWipData.Tables("Wip").Rows(0)
        Dim InDataRow As DataRow = InWipData.Tables("Sno").Rows(0)
        With InstRow
            Using dtCust As DataTable = DAO.ExecQry(_DAL.GetCustomerData, New Object() {InstRow("Custid"), InstRow("ServiceType")})
                If dtCust.Rows.Count > 0 Then
                    .Item("CustName") = dtCust.Rows(0)("CustName")
                End If
            End Using
            If InstWipData.Tables.Contains("Facility") Then
                Dim dtFaci As DataTable = InstWipData.Tables("Facility")
                If dtFaci.Rows.Count > 0 Then
                    .Item("ID") = dtFaci.Rows(0)("ID")
                End If
            End If
            If Not String.IsNullOrEmpty(InstRow("ID").ToString) Then
                Using dtSO137 As DataTable = DAO.ExecQry(_DAL.GetSO137, New Object() {InstRow("ID")})
                    If dtSO137.Rows.Count > 0 Then
                        .Item("ContName") = dtSO137.Rows(0)("DeclarantName")
                        .Item("ContTel") = dtSO137.Rows(0)("ContTel")
                        .Item("Contmobile") = dtSO137.Rows(0)("Contmobile")
                    End If
                End Using
            End If

            .Item("AcceptEn") = InDataRow.Item("AcceptEn")
            Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {.Item("AcceptEn")})
                .Item("AcceptName") = EmpName.Rows(0).Item("EmpName")
            End Using
            .Item("AcceptTime") = DateTime.Now
            .Item("ModifyFlag") = 1
            .Item("PrintBillFlag") = 0
            .Item("Note") = InDataRow.Item("Note")
            .Item("UpdEn") = LoginInfo.EntryName
            .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(DateTime.Now)
            .Item("NewUpdTime") = DateTime.Now
        End With
        InstWipData.Tables("Wip").Rows(0).AcceptChanges()
        Return True
    End Function
    Private Function CheckCanPR(InWipData As DataSet) As RIAResult
        Using Vali As New Validate(LoginInfo, DAO)
            Dim InRow As DataRow = InWipData.Tables("SNo").Rows(0)
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                If Not String.IsNullOrEmpty(InRow("ReasonDescCode").ToString) Then
                    Using Code As DataTable = DAO.ExecQry(_DAL.GetReasonDescName(), New Object() {InRow("ReasonDescCode")})
                        If Code.Rows.Count = 0 Then
                            Return New RIAResult With {.ErrorCode = -171, .ErrorMessage = Lang.NoPRDetailReasno}
                        End If
                    End Using
                End If
                Using PRCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, InRow.Item("PRCode").ToString, False)
                    ServiceType = PRCode.Rows(0).Item("ServiceType")
                    Using Customer As DataTable = DAO.ExecQry(_DAL.GetCustomerData(), New Object() {InRow.Item("CustId"), ServiceType})
                        Return Vali.CheckCanPR(PRCode.Rows(0).Item("CodeNo"), CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(PRCode.Rows(0).Item("RefNo")), CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(PRCode.Rows(0).Item("Interdepend")), Customer.Rows(0).Item("CustStatusCode"), Customer.Rows(0).Item("WipCode3"), InRow.Item("CustId"), ServiceType, Customer.Rows(0).Item("InstAddrNo"), Nothing)
                    End Using
                End Using
            End Using
        End Using
    End Function
    Private Function ChkCanResv(WipData As System.Data.DataSet, ResvTime As DateTime) As RIAResult
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
    Private Function GetReturnData(WipData As DataSet, InstWipData As DataSet) As DataSet
        Dim RetData As New DataSet With {.DataSetName = "DataSet"}
        Dim RetTable As New DataTable With {.TableName = "SNo"}
        RetTable.Columns.Add(New DataColumn With {.ColumnName = "SNo", .DataType = GetType(String)})
        RetTable.Rows.Add(RetTable.NewRow())
        RetTable.Rows(0).Item("SNo") = String.Format("{0},{1}", WipData.Tables("Wip").Rows(0).Item("SNo"), InstWipData.Tables("Wip").Rows(0).Item("SNo"))
        RetData.Tables.Add(RetTable)
        Return RetData
    End Function

    Private Function ChangeErrorCode(ByRef Changeresult As RIAResult) As RIAResult
        Select Case Changeresult.ErrorCode
            Case -11005
                Changeresult.ErrorCode = -173
                If String.IsNullOrEmpty(Changeresult.ErrorMessage) Then
                    Changeresult.ErrorMessage = Lang.FullPoint
                End If
        End Select
        Return Changeresult
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
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'APIID', '命令識別碼', 0, 1, 'Main', -1, 1);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'Compcode', '公司別', 1, 1, 'Main', -1, 2);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'Caller', '呼叫來源', 0, 1, 'Main', -1, 3);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'Seqno', '來源識別碼', 0, 1, 'Main', -1, 4);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'CustId', '客戶編號', 0, 1, 'SNo', -1, 5);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'PRCode', '停拆機類別代碼', 0, 1, 'SNo', -1, 6);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'ReasonCode', '停拆機原因代碼', 0, 1, 'SNo', -1, 7);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'ResvTime', '預約時間', 3, 1, 'SNo', -1, 8);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'AcceptEn', '受理人員代號', 0, 1, 'SNo', -1, 9);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'Note', '訂單備註', 0, 0, 'SNo', -1, 10);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'ReasonDescCode', '停拆機原因細項', 0, 0, 'SNo', -1, 11);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'CallSeqNo', '互動單號', 0, 1, 'SNo', -1, 12);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'NewCustId', '新址客編', 0, 1, 'SNo', -1, 13);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('255', 'FaciSeqNo', '設備流水號', 0, 1, 'SNo', -1, 14);
'================================================================================================================================
'SO1114B.CustId 預約時間 ChkSQLQuery設定內容
'Select -304 ErrorCode,'客編不存在' ErrorMsg From Dual
'Where (Select Count(*) From SO001 Where CustId = '[CustId]') = 0
'Union All
'Select -159 ErrorCode,'停拆機類別不存在' ErrorMsg From Dual
'Where (Select Count(*) From CD007 Where CodeNo = '[PRCode]' And StopFlag = 0) = 0
'Union All
'Select -163 ErrorCode,'停拆機原因不存在' ErrorMsg From Dual
'Where (Select Count(*) From CD014 Where CodeNo = '[ReasonCode]' And StopFlag = 0) = 0
'Union All
'Select -157 ErrorCode,'受理人員不存在' ErrorMsg From Dual
'Where (Select Count(*) From CM003 Where EmpNo = '[AcceptEn]' And StopFlag = 0) = 0
'Union All
'Select -171 ErrorCode,'停拆機原因細項不存在' ErrorMsg From Dual
'Where (Select Count(*) From CD014A Where CodeNo = '[ReasonDescCode]' And StopFlag = 0) = 0 And '[ReasonDescCode]' is not null
'Union All
'Select -104 ErrorCode,'互動單號不存在' ErrorMsg From Dual
'Where (Select Count(*) From 
'(Select 1 From SO006A Where CustId = [CustId] And SeqNo = '[CallSeqNo]' 
'Union All Select 1 From SO006 Where CustId = [CustId] And SeqNo = '[CallSeqNo]')) = 0
'Union All
'Select -172 ErrorCode,'新客編不存在' ErrorMsg From Dual
'Where (Select Count(*) From SO001 Where CustId = '[NewCustId]') = 0
'Union All
'Select -169 ErrorCode,'客戶狀態為註銷不允許新增' ErrorMsg From SO002 A
'Where A.CustId = [CustId] And A.CustStatusCode = 4 