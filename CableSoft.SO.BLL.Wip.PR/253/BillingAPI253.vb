Imports System.Data.Common
Imports CableSoft.Utility.DataAccess
Imports CableSoft.BLL.Utility
'新增CATV移機工單(無結清)－253
'	使用時機：新增移機工單資料於營運系統。
'	使用限制：全產品移機(CATV移機)
Public Class BillingAPI253
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private _DAL As New BillingAPI253DAL(Me.LoginInfo.Provider)
    'Private _DAL As New BillingAPI253DALMultiDB(Me.LoginInfo.Provider)
    Private ServiceType As String = Nothing
    Private Lang As New CableSoft.BLL.Language.SO61.BillingAPI253Language
    Private SOUtil As CableSoft.SO.BLL.Utility.Utility = Nothing
    Private _Ref3ServiceType As String = "C"
    '回應碼	回應狀態	回應訊息
    '0   	成功	
    '-1	    失敗	    {參數}資料有誤!!
    '-304	失敗      	客編不存在
    '-159	失敗      	停拆機類別不存在
    '-402	失敗      	新地址編號不存在
    '-157	失敗      	受理人員不存在
    '-403	失敗      	新收費地址不存在
    '-404	失敗      	新郵寄地址不存在
    '-163	失敗      	停拆機原因不存在
    '-104	失敗	    互動單號不存在
    '-155	失敗	    己排件數大於預設件數，不允許預約
    Private Sub CreateMoveFaciTable(ByRef tb As DataTable)
        If tb Is Nothing Then
            tb = New DataTable
            tb.Columns.Add("SERVICETYPE", GetType(String))
            tb.Columns.Add("PRCODE", GetType(String))
            tb.Columns.Add("PRNAME", GetType(String))
            tb.Columns.Add("REASONCODE", GetType(String))
            tb.Columns.Add("REASONNAME", GetType(String))
            tb.TableName = "Wip"
        End If

    End Sub

    Public Property Ref3ServiceType() As String
        Get
            Return _Ref3ServiceType
        End Get
        Set(ByVal value As String)
            _Ref3ServiceType = value
        End Set
    End Property
    Private Function isRelativeFaci(ByVal intCustId As Integer, ByVal InData As System.Data.DataSet, ByVal strMainWorkCode As String) As RIAResult
        Dim strCalcFaciRefNo As String = Nothing
        Dim result As RIAResult = New RIAResult With {.ResultBoolean = True}

        Dim InterDependRefNo As String = Nothing
        For i As Integer = 0 To 1
            Select Case i
                Case 0
                    strCalcFaciRefNo = CableSoft.SO.BLL.Utility.Utility.GetServiceCanChooseRefNo(DAO, "D", False, True)
                    Using MainWorkCode As DataTable = SOUtil.GetCode(SO.BLL.Utility.CodeType.PRCode, "InterDependRefNo", "CodeNo = " & strMainWorkCode)
                        If MainWorkCode.Rows.Count > 0 Then
                            If MainWorkCode.Rows(0).IsNull("InterDependRefNo") = False Then
                                InterDependRefNo = MainWorkCode.Rows(0).Item("InterDependRefNo")
                            End If
                        End If
                    End Using
                    If String.IsNullOrEmpty(InterDependRefNo) Then InterDependRefNo = "-1"
                    Using tb As DataTable = DAO.ExecQry(_DAL.GetMoveFaciData(InterDependRefNo, strCalcFaciRefNo), New Object() {intCustId, "D"})
                        If tb.Rows.Count > 0 Then
                            If DBNull.Value.Equals(InData.Tables("Sno").Rows(0).Item("DTVPRCode")) Then

                                result.ResultBoolean = False
                                result.ErrorCode = -1
                                result.ErrorMessage = Lang.mustDTVPRCode
                                Return result
                            End If
                            If DBNull.Value.Equals(InData.Tables("Sno").Rows(0).Item("DTVReasonCode")) Then
                                result.ResultBoolean = False
                                result.ErrorCode = -1
                                result.ErrorMessage = Lang.mustDTVReasonCode
                                Return result
                            End If

                        End If
                    End Using
                Case 1
                    strCalcFaciRefNo = CableSoft.SO.BLL.Utility.Utility.GetServiceCanChooseRefNo(DAO, "I", False, True)
                    Using MainWorkCode As DataTable = SOUtil.GetCode(SO.BLL.Utility.CodeType.PRCode, "InterDependRefNo", "CodeNo = " & strMainWorkCode)
                        If MainWorkCode.Rows.Count > 0 Then
                            If MainWorkCode.Rows(0).IsNull("InterDependRefNo") = False Then
                                InterDependRefNo = MainWorkCode.Rows(0).Item("InterDependRefNo")
                            End If
                        End If
                    End Using
                    If String.IsNullOrEmpty(InterDependRefNo) Then InterDependRefNo = "-1"
                    Using tb As DataTable = DAO.ExecQry(_DAL.GetMoveFaciData(InterDependRefNo, strCalcFaciRefNo), New Object() {intCustId, "I"})
                        If tb.Rows.Count > 0 Then
                            If DBNull.Value.Equals(InData.Tables("Sno").Rows(0).Item("CMPRCode")) Then

                                result.ResultBoolean = False
                                result.ErrorCode = -1
                                result.ErrorMessage = Lang.mustCMPRCode
                                Return result
                            End If
                            If DBNull.Value.Equals(InData.Tables("Sno").Rows(0).Item("CMReasonCode")) Then
                                result.ResultBoolean = False
                                result.ErrorCode = -1
                                result.ErrorMessage = Lang.mustCMReasonCode
                                Return result
                            End If

                        End If
                    End Using
            End Select
        Next
        Return result


    End Function


    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim result As RIAResult = Nothing
        Dim ResvTime As DateTime = DateTime.Parse(InData.Tables("SNo").Rows(0).Item("ResvTime"))

        'SNO	單	CustId	        客戶編號	        V	
        'SNO	單	PRCode	        停拆機類別代碼	    V	取停拆移機代碼之參考號碼為3
        'SNO	單	ReasonCode	    停拆機原因代碼	    V	
        'SNO	單	ReInstAddrNo	新址編號	        V	
        'SNO	單	ResvTime	    預約時間	        V	YYYY/MM/DD HH24:MI:SS
        'SNO	單	AcceptEn	    受理人員代號	    V	
        'SNO	單	NewTel1	        新電話(1)		
        'SNO	單	NewChargeAddrNo	新收費地址編號		
        'SNO	單	NewMailAddrNo	新郵寄地址編號		
        'SNO	單	Note	        備註		
        'SNO	單	NewTel2	        新電話(2)		
        'SNO	單	NewTel3	        新電話(3)		
        'SNO	單	ReasonDescCode	停拆機原因細項		
        'SNO	單	CallSeqNo	    互動單號	        V	
        'SNO	單	WorkServCode    實際派工服務區代碼  V '#7899 增加
        Dim WipData As DataSet = Nothing
        Dim MoveFaciData As DataSet = Nothing
        Dim tbMoveFaciData As DataTable = Nothing

        '#8706
        Using EmpName As DataTable = DAO.ExecQry(_DAL.GetEmpName(), New Object() {InData.Tables("SNo").Rows(0).Item("AcceptEn")})
            Me.LoginInfo.EntryId = InData.Tables("SNo").Rows(0).Item("AcceptEn")
            Me.LoginInfo.EntryName = EmpName.Rows(0).Item("EmpName")
        End Using
        SOUtil = New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
        result = isRelativeFaci(Integer.Parse(InData.Tables("Sno").Rows(0).Item("CustId")), _
                                        InData, InData.Tables("Sno").Rows(0).Item("PRCode"))
        If result.ResultBoolean = False Then
            Return result
        End If
        '檢核是否可派工
        result = CheckCanPR(InData)
        If result.ResultBoolean = False Then
            Return result
        End If
        Using PR As New PR(LoginInfo, DAO)
            WipData = PR.GetPRData(Nothing, InData.Tables("Sno").Rows(0).Item("CustId"), ServiceType)
            If Not GetWipData(InData, WipData) Then
                Throw New Exception("GetWipData")
            End If
            '檢核預約時間是否可以改約
            result = ChkCanResv(WipData, ResvTime)
            If result.ResultBoolean = False Then
                result.ErrorCode = -155
                Return result
            End If
        End Using
        If InData.Tables("Sno").Columns.Contains("DTVPRCode") Then
            If Not DBNull.Value.Equals(InData.Tables("Sno").Rows(0).Item("DTVPRCode")) Then
                CreateMoveFaciTable(tbMoveFaciData)
                Dim rw As DataRow = tbMoveFaciData.NewRow
                rw.Item("ServiceType") = "D"
                rw.Item("PRCODE") = InData.Tables("Sno").Rows(0).Item("DTVPRCode")
                rw.Item("PRNAME") = DAO.ExecSclr(_DAL.getCD007ByCode, New Object() _
                                                        {Integer.Parse(InData.Tables("Sno").Rows(0).Item("DTVPRCode"))}).ToString
                rw.Item("REASONCODE") = InData.Tables("Sno").Rows(0).Item("DTVReasonCode")
                rw.Item("REASONNAME") = DAO.ExecSclr(_DAL.getCD014ByCode, New Object() _
                                                        {Integer.Parse(InData.Tables("Sno").Rows(0).Item("DTVReasonCode"))}).ToString
                tbMoveFaciData.Rows.Add(rw)
            End If
        End If
        If InData.Tables("Sno").Columns.Contains("CMPRCode") Then
            If Not DBNull.Value.Equals(InData.Tables("Sno").Rows(0).Item("CMPRCode")) Then
                CreateMoveFaciTable(tbMoveFaciData)
                Dim rw As DataRow = tbMoveFaciData.NewRow
                rw.Item("ServiceType") = "I"
                rw.Item("PRCODE") = InData.Tables("Sno").Rows(0).Item("CMPRCode")
                rw.Item("PRNAME") = DAO.ExecSclr(_DAL.getCD007ByCode, New Object() _
                                                        {Integer.Parse(InData.Tables("Sno").Rows(0).Item("CMPRCode"))}).ToString
                rw.Item("REASONCODE") = InData.Tables("Sno").Rows(0).Item("CMReasonCode")
                rw.Item("REASONNAME") = DAO.ExecSclr(_DAL.getCD014ByCode, New Object() _
                                                        {Integer.Parse(InData.Tables("Sno").Rows(0).Item("CMReasonCode"))}).ToString
                tbMoveFaciData.Rows.Add(rw)
            End If
        End If
        If tbMoveFaciData IsNot Nothing Then
            MoveFaciData = New DataSet
            MoveFaciData.Tables.Add(tbMoveFaciData.Copy)
        End If
        WipData.Tables("Wip").Rows(0).Item("ResvTime") = ResvTime
        '檢核移拆工單是否可存檔
        Using Vali As New Validate(LoginInfo, DAO)
            result = Vali.ChkDataOk(EditMode.Append, WipData)
            If result.ResultBoolean = False Then
                Return result
            End If
        End Using
        '工單存檔
        Using bll As New SaveData(LoginInfo, DAO)
            bll.Ref3ServiceType = Me._Ref3ServiceType
            If (InData.Tables("SNo").Columns.Contains("FaciSeqNo")) AndAlso (Not DBNull.Value.Equals(InData.Tables("SNo").Columns.Contains("FaciSeqNo"))) Then
                bll.Ref3FilterFaciSeqNo = InData.Tables("SNo").Rows(0).Item("FaciSeqNo")
            End If

            If MoveFaciData Is Nothing Then
                result = bll.Save(EditMode.Append, False, WipData, False)
            Else
                result = bll.Save(EditMode.Append, False, WipData, Nothing, False, MoveFaciData)
            End If

            If result.ResultBoolean = False Then
                Return result
            End If
        End Using
        '#8787 客戶派CATV移機單時，若有未結案的促變工單、加購工單，其工單地址需改成移機新址
        If (Not DBNull.Value.Equals(InData.Tables("SNo").Rows(0).Item("UseReInstaddr"))) AndAlso _
                (Integer.Parse(InData.Tables("SNo").Rows(0).Item("UseReInstaddr")) = 1) Then
            result = updOtherSNoAddress(WipData)
            If result.ResultBoolean = False Then
                Return result
            End If
        End If

        '回傳資料
        result.ResultDataSet = GetReturnData(WipData)
     

        WipData.Dispose()
        Return result
    End Function
    '#8787 客戶派CATV移機單時，若有未結案的促變工單、加購工單，其工單地址需改成移機新址(增加此Function)
    Private Function updOtherSNoAddress(ByVal WipData As DataSet) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = String.Empty}        
        Try

            'Update SO009 set ReInstAddrNo={0}0,ReInstAddress={0}1, " & _
            '                                 "servcode = {0}2,strtcode={0}3,salecode={0}4,salename={0}5 " & _
            '                         " Where SNo <> {0}6 And signdate is null And Nvl(PrtCount,0) = 0 And orderno is not null 
            Dim ReInstAddrNo As Object = WipData.Tables("Wip").Rows(0).Item("ReInstAddrNo")
            Dim ReInstAddress As Object = WipData.Tables("Wip").Rows(0).Item("ReInstAddress")
            Dim servcode As Object = WipData.Tables("Wip").Rows(0).Item("servcode")
            Dim strtcode As Object = WipData.Tables("Wip").Rows(0).Item("strtcode")
            Dim SalesCode As Object = WipData.Tables("Wip").Rows(0).Item("SalesCode")
            Dim SalesName As Object = WipData.Tables("Wip").Rows(0).Item("SalesName")
            Dim sno As Object = WipData.Tables("Wip").Rows(0).Item("SNo")
            DAO.ExecSclr(_DAL.updSO009SnoAddr, New Object() {ReInstAddrNo, ReInstAddress, _
                                                             servcode, strtcode, SalesCode, SalesName, ReInstAddrNo, ReInstAddress, sno})

            DAO.ExecSclr(_DAL.updSO007SnoAddr, New Object() {ReInstAddrNo, ReInstAddress, _
                                                             servcode, strtcode, SalesCode, SalesName})
            Return result
        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -1
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

    Private Function GetWipData(InWipData As DataSet, ByRef WipData As DataSet) As Boolean
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
                    '#8802
                    If InWipData.Tables("SNo").Columns.Contains("FaciSeqNo") Then
                        '指定移出變更()
                        Using bll As New CableSoft.SO.BLL.Facility.ChangeFaci.ChangeFaci(LoginInfo, DAO)
                            For Each tFaciSeqNo As String In InDataRow.Item("FaciSeqNo").ToString.Split(",")
                                Dim FaciSeqNo As String = tFaciSeqNo
                                If WipData.Tables("ChangeFacility").AsEnumerable.Where(Function(list) list.IsNull("SeqNo") = False AndAlso list.Item("SeqNo") = FaciSeqNo).Count = 0 Then
                                    Using RetData As DataTable = bll.GetMoveFaci(SNo, FaciSeqNo, True)
                                        For Each Row As DataRow In RetData.Rows                                           
                                            WipData.Tables("ChangeFacility").Rows.Add(CableSoft.BLL.Utility.Utility.CopyDataRow(Row, WipData.Tables("ChangeFacility").NewRow()))
                                        Next
                                        If WipData.Tables.Contains("Charge") Then
                                            For Each drCharge As DataRow In WipData.Tables("Charge").Rows
                                                drCharge("FaciSNo") = DAO.ExecSclr(_DAL.getFaciSNoBySeqNo, New Object() {tFaciSeqNo})
                                                drCharge("FaciSeqNo") = tFaciSeqNo
                                            Next
                                        End If
                                    End Using
                                End If
                            Next
                        End Using
                    End If
                   
                End Using
            End Using
        End Using
        'If ServiceWipData.Tables.Contains("Charge") Then
        '    For Each drCharge As DataRow In ServiceWipData.Tables("Charge").Rows
        '        drCharge("FaciSNO") = FaciRow("FaciSNO")
        '        drCharge("FaciSeqno") = FaciRow("SEQNO")
        '    Next
        '    ServiceWipData.Tables("Charge").AcceptChanges()
        'End If
        '異動工單相關欄位
        If Not UpdateWipHead(InWipData, WipData) Then
            Throw New Exception("UpdateWipHead")
        End If

        Return True
    End Function

    Private Function UpdateWipHead(InWipData As DataSet, ByRef WipData As DataSet) As Boolean
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
            If Not String.IsNullOrEmpty(InDataRow("ReInstAddrNo").ToString) Then
                .Item("ReInstAddrNo") = InDataRow.Item("ReInstAddrNo")
                Using Code As DataTable = DAO.ExecQry(_DAL.GetAddress(), New Object() {InDataRow.Item("ReInstAddrNo")})
                    .Item("ReInstAddress") = Code.Rows(0).Item("Address")
                End Using
            End If
            If Not String.IsNullOrEmpty(InDataRow("NewChargeAddrNo").ToString) Then
                .Item("NewChargeAddrNo") = InDataRow.Item("NewChargeAddrNo")
                Using Code As DataTable = DAO.ExecQry(_DAL.GetAddress(), New Object() {InDataRow.Item("NewChargeAddrNo")})
                    .Item("NewChargeAddress") = Code.Rows(0).Item("Address")
                End Using
            End If
            If Not String.IsNullOrEmpty(InDataRow("NewMailAddrNo").ToString) Then
                .Item("NewMailAddrNo") = InDataRow.Item("NewMailAddrNo")
                Using Code As DataTable = DAO.ExecQry(_DAL.GetAddress(), New Object() {InDataRow.Item("NewMailAddrNo")})
                    .Item("NewMailAddress") = Code.Rows(0).Item("Address")
                End Using
            End If
            If Not String.IsNullOrEmpty(InDataRow("NewTel1").ToString) Then .Item("NewTel1") = InDataRow.Item("NewTel1")
            If Not String.IsNullOrEmpty(InDataRow("NewTel2").ToString) Then .Item("NewTel2") = InDataRow.Item("NewTel2")
            If Not String.IsNullOrEmpty(InDataRow("NewTel3").ToString) Then .Item("NewTel3") = InDataRow.Item("NewTel3")

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
            If String.IsNullOrEmpty(.Item("ReInstDate").ToString) Then
                If InWipData.Tables("SNO").Columns.Contains("ReInstDate") Then
                    If Not String.IsNullOrEmpty(InDataRow("ReInstDate").ToString) Then .Item("ReInstDate") = InDataRow("ReInstDate")
                Else
                    .Item("ReInstDate") = InDataRow("ResvTime")
                End If
            End If

            If Not String.IsNullOrEmpty(InDataRow("Note").ToString) Then .Item("Note") = InDataRow.Item("Note")
            .Item("AcceptTime") = DateTime.Now
            .Item("ModifyFlag") = 1
            .Item("PrintBillFlag") = 0
            .Item("UpdEn") = LoginInfo.EntryName
            .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(DateTime.Now)
            .Item("NewUpdTime") = DateTime.Now
        End With
        WipData.Tables("Wip").Rows(0).AcceptChanges()
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
            '檢核預約時間是否可以預約
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

    Private Function GetReturnData(WipData As DataSet) As DataSet
        Dim RetData As New DataSet With {.DataSetName = "DataSet"}
        Dim RetTable As New DataTable With {.TableName = "SNo"}
        RetTable.Columns.Add(New DataColumn With {.ColumnName = "SNo", .DataType = GetType(String)})
        RetTable.Rows.Add(RetTable.NewRow())
        RetTable.Rows(0).Item("SNo") = WipData.Tables("Wip").Rows(0).Item("SNo")
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



'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'APIID', '命令識別碼', 0, 1, 'Main', -1, 1);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'Compcode', '公司別', 1, 1, 'Main', -1, 2);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'Caller', '呼叫來源', 0, 1, 'Main', -1, 3);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'Seqno', '來源識別碼', 0, 1, 'Main', -1, 4);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'CustId', '客戶編號', 0, 1, 'SNo', -1, 5);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'PRCode', '停拆機類別代碼', 0, 1, 'SNo', -1, 6);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'ReasonCode', '停拆機原因代碼', 0, 1, 'SNo', -1, 7);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'ReInstAddrNo', '新址編號', 0, 1, 'SNO', -1, 8);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'ResvTime', '預約時間', 3, 1, 'SNo', -1, 9);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'AcceptEn', '受理人員代號', 0, 1, 'SNo', -1, 10);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'NewTel1', '新電話(1)', 0, 0, 'SNO', -1, 11);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'NewChargeAddrNo', '新收費地址編號', 0, 0, 'SNO', -1, 12);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'NewMailAddrNo', '新郵寄地址編號', 0, 0, 'SNO', -1, 13);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'Note', '訂單備註', 0, 0, 'SNo', -1, 14);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'NewTel2', '新電話(2)', 0, 0, 'SNO', -1, 15);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'NewTel3', '新電話(3)', 0, 0, 'SNO', -1, 16);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'ReasonDescCode', '停拆機原因細項', 0, 0, 'SNo', -1, 17);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('253', 'CallSeqNo', '互動單號', 0, 1, 'SNo', -1, 18);

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
' Union All 
' Select 1 From SO006 Where CustId = [CustId] And SeqNo = '[CallSeqNo]')) = 0
'Union All
'Select -402 ErrorCode,'新地址編號不存在' ErrorMsg From Dual
'Where (Select Count(*) From SO014 Where AddrNo = '[ReInstAddrNo]') = 0
'Union All
'Select -403 ErrorCode,'新收費地址不存在' ErrorMsg From Dual
'Where (Select Count(*) From SO014 Where AddrNo = '[NewChargeAddrNo]') = 0
'Union All
'Select -404 ErrorCode,'新郵寄地址不存在' ErrorMsg From Dual
'Where (Select Count(*) From SO014 Where AddrNo = '[NewMailAddrNo]') = 0
