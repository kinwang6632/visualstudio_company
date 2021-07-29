Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess
Imports ValidateLanguage = CableSoft.BLL.Language.SO61.WipPRLanguage

Public Class Validate
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New PRDALMultiDB(Me.LoginInfo.Provider)
    'Private _DAL As New PRDAL(Me.LoginInfo.Provider)
    Private _ValidateDAL As New ValidateDALMultiDB(Me.LoginInfo.Provider)
    'Private _ValidateDAL As New ValidateDAL(Me.LoginInfo.Provider)

    Public Sub New()
    End Sub

    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub


    Public Function CheckCanPR(ByVal PRCode As Int32, ByVal Custid As Int32, ByVal ServiceType As String) As RIAResult
        Try
            Dim dtPRCode As DataTable = DAO.ExecQry(_DAL.GetPRCode(String.Empty, String.Empty, PRCode, False, False, False), New Object() {ServiceType})
            Dim intRefNo As Integer = 0
            If dtPRCode.Rows.Count > 0 Then intRefNo = Integer.Parse("0" & dtPRCode.Rows(0)("RefNo").ToString)
            Dim dtServiceType As DataTable = DAO.ExecQry(_DAL.GetServiceType(ServiceType))
            Dim dtSO001 As DataTable = DAO.ExecQry(_DAL.GetSO001, New Object() {Custid})
            Dim dtSO002 As DataTable = DAO.ExecQry(_DAL.GetSO002(ServiceType), New Object() {Custid})
            Dim interdepend As Integer = 0
            If dtServiceType.Rows(0)("DependService").ToString <> ServiceType Then interdepend = 1
            Return CheckCanPR(PRCode, intRefNo, interdepend, dtSO002.Rows(0)("CustStatusCode"), dtSO002.Rows(0)("WipCode3"), Custid, ServiceType, dtSO001.Rows(0)("InstAddrNo"), Nothing)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 檢查停拆機類別是否正常可派(CheckCanPR)
    ''' </summary>
    ''' <param name="PRCode">停拆機類別</param>
    ''' <param name="PRRefNo">停拆機參考號</param>
    ''' <param name="Interdepend">服務依存</param>
    ''' <param name="CustStatusCode">客戶狀態</param>
    ''' <param name="WipCode3">派工類別3</param>
    ''' <param name="Custid">客戶編號</param>
    ''' <param name="ServiceType">公司別</param>
    ''' <param name="InstAddrNo">裝機地址編號</param>
    ''' <param name="WinPRData">拆機單資訊</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckCanPR(ByVal PRCode As Int32, ByVal PRRefNo As Int32,
                               ByVal Interdepend As Int32, ByVal CustStatusCode As Int32,
                               ByVal WipCode3 As String, ByVal Custid As Int32,
                               ByVal ServiceType As String, ByVal InstAddrNo As Int64,
                               ByVal WinPRData As DataTable) As RIAResult
        Try
            Dim dtSystem As DataTable = Nothing
            Dim dtSystem2 As DataTable = Nothing

            Using soUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                dtSystem = soUtil.GetSystem(CableSoft.SO.BLL.Utility.SystemTableType.System, "*", ServiceType)
                dtSystem2 = soUtil.GetSystem(CableSoft.SO.BLL.Utility.SystemTableType.Wip, "*", ServiceType)
            End Using
            '#8789
            Dim abnormalFaci As Integer = 0
            If dtSystem2.Rows.Count > 0 Then
                If Not DBNull.Value.Equals(dtSystem2.Rows(0).Item("abnormalFaci").ToString()) AndAlso _
                    Integer.Parse(dtSystem2.Rows(0).Item("abnormalFaci").ToString()) = 1 Then
                    abnormalFaci = 1
                End If
            End If


            Dim result As New RIAResult() With {.ErrorCode = -99, .ErrorMessage = ""}
            'CD007	停拆移機類別代碼檔
            '1=停機	2=拆機	3=移機	4=移拆
            '5=未繳費拆機	6=拆設備	7=軟關	8=拆分機(CATV)
            '9=取回設備	10=頻道停權	11=停設備	12=帳號停權
            '13=取消申請	14=取消固定IP	15=暫停頻道
            Dim ErrorMessage As String = String.Empty
            If CustStatusCode = 4 Then
                '註銷戶無法產生派工單(CustStatusCode = 4)
                ErrorMessage = ValidateLanguage.CustIsCancel
            ElseIf PRRefNo = 4 Then
                '不可直接新增移拆單(PRRefNo = 4)。
                ErrorMessage = "不可直接新增移拆單"
            ElseIf dtSystem.Rows.Count > 0 Then
                Dim drSystem As DataRow = dtSystem.Rows(0)
                Dim drSystem2 As DataRow = dtSystem2.Rows(0)
                If String.IsNullOrEmpty(drSystem2("CanStop")) OrElse drSystem2("CanStop") = 0 Then
                    'SO042.CanStop = 0 不可直接新增停(分)機單(PRRefNo = 1 Or PRRefNo = 11)。
                    If PRRefNo = 1 OrElse PRRefNo = 11 Then
                        ErrorMessage = ValidateLanguage.WipNotAddPR1
                    End If
                ElseIf Not String.IsNullOrEmpty(drSystem2("FaciRecoupSNO")) Then
                    'SO042.FaciRecoupSNO = 1, 不可直接新增取回單(PRRefNO = 9)。
                    If drSystem2("FaciRecoupSNO") = 1 AndAlso PRRefNo = 9 Then
                        ErrorMessage = ValidateLanguage.WipNotAddPR
                    End If
                End If

                'SO042.PayNowCanPR = 0, 現付制不可直接派拆機單(PRRefNo = 2 or 5 且 SO002.PayKind = 1 )
                If String.IsNullOrEmpty(drSystem("PayNowCanPR")) OrElse drSystem("PayNowCanPR") = 0 Then
                    Using dtSO002 As DataTable = DAO.ExecQry(_DAL.GetSO002(ServiceType), New Object() {Custid})
                        If PRRefNo = 2 OrElse PRRefNo = 5 Then
                            If Not String.IsNullOrEmpty(dtSO002.Rows(0)("PayKind")) And dtSO002.Rows(0)("PayKind") = 1 Then
                                ErrorMessage = ValidateLanguage.WipPRNotPayNow
                            End If
                        End If
                    End Using
                End If
            End If
            If ErrorMessage = String.Empty Then
                '呼叫ChkPRInterdepend, 檢核服務依存是否正常。
                result.ErrorCode = 0
                result.ResultBoolean = True
                result.ErrorMessage = ""
                'result = ChkPRInterdepend(PRRefNo, Interdepend, Custid, instAddrNo, ServiceType, MainSNO) '先不要處理退單的檢核
            Else
                result.ErrorMessage = ErrorMessage
                result.ResultBoolean = False
            End If
            Return result
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function ChkCanResv(ByVal ServCode As String, ByVal WipCode As Int32,
                               ByVal MCode As Int32, ByVal ServiceType As String,
                               ByVal ResvTime As Date,
                               ByVal AcceptTime As Date, ByVal OldResvTime As Date,
                               ByVal Resvdatebefore As Int32,
                               ByVal WorkUnit As Decimal, ByVal IsBookIng As Boolean) As RIAResult
        Return ChkCanResv(ServCode, WipCode, MCode, ServiceType, ResvTime,
                          AcceptTime, OldResvTime, Resvdatebefore, WorkUnit, IsBookIng, ServCode)

    End Function
    ''' <summary>
    ''' 檢查預約時段是否可派工
    ''' </summary>
    ''' <param name="ServCode">服務區</param>
    ''' <param name="WipCode"></param>
    ''' <param name="MCode">裝機類別名稱</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="ResvTime">預約時間</param>
    ''' <param name="AcceptTime">受理時間</param>
    ''' <param name="OldResvTime">舊預約時間</param>
    ''' <param name="Resvdatebefore"></param>
    ''' <param name="WorkUnit">派工點數</param>
    ''' <param name="IsBookIng">預約時間異動時才需要傳True，檢核時傳False</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChkCanResv(ByVal ServCode As String, ByVal WipCode As Int32,
                               ByVal MCode As Int32, ByVal ServiceType As String,
                               ByVal ResvTime As Date,
                               ByVal AcceptTime As Date, ByVal OldResvTime As Date,
                               ByVal Resvdatebefore As Int32,
                               ByVal WorkUnit As Decimal, ByVal IsBookIng As Boolean, ByVal oldServCode As String) As RIAResult
        '2014.07.18 注:結清會呼叫到這個功能
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try

            Return obj.ChkCanResv(BLL.Utility.InvoiceType.PR,
                                  WipCode, ServCode, MCode, ServiceType,
                                  ResvTime, AcceptTime, OldResvTime, Resvdatebefore,
                                  WorkUnit, IsBookIng, Nothing, oldServCode)
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 檢查停拆裝類別服務依存是否可派(ChkPRInterdepend)
    ''' </summary>
    ''' <param name="PRRefNo">拆機派工類別</param>
    ''' <param name="Interdepend"></param>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="instAddrNo"></param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="MainSNO">工單號碼</param>
    ''' <param name="WinPRData">拆機工單資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChkPRInterdepend(ByVal PRRefNo As Integer, ByVal Interdepend As Integer,
                                    ByVal CustId As Integer, ByVal instAddrNo As Int64,
                                    ByVal ServiceType As String, ByVal MainSNO As String,
                                    ByVal WinPRData As DataTable) As RIAResult
        Dim result As New RIAResult With {.ErrorCode = 0, .ErrorMessage = "", .ResultBoolean = True}
        Dim ErrorMessage As String = String.Empty
        Try
            If WinPRData IsNot Nothing AndAlso WinPRData.Select(String.Format("ServiceType='{0}'", ServiceType)).Count > 0 Then
                Return result
            End If
            Dim ServiceInterdepend As Integer = 0
            Using soUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                ServiceInterdepend = soUtil.GetSystem(CableSoft.SO.BLL.Utility.SystemTableType.System, "ServiceInterdepend", ServiceType).Rows(0).Item(0)
            End Using
            If ServiceInterdepend > 0 And Interdepend = 1 Then
                Using dtCD046 As DataTable = DAO.ExecQry(_DAL.GetServiceType(ServiceType))
                    If "1,2,5".IndexOf(PRRefNo.ToString) >= 0 Then
                        If dtCD046.Rows.Count > 0 Then
                            Dim CustStatusSQL As String = String.Empty
                            If PRRefNo = 1 Then
                                CustStatusSQL = " And  B.CustStatusCode = 1"
                            Else
                                CustStatusSQL = " And B.CustStatusCode in (1,2)"
                            End If
                            If Not String.IsNullOrEmpty(dtCD046.Rows(0)("DependService")) And dtCD046.Rows(0)("DependService").ToString.ToUpper <> ServiceType.ToString.ToUpper Then
                                Dim strSQL As String = String.Format("Select Count(*) From SO001 A,SO002 B Where A.CustId = B.CustId And A.InstAddrNo = {0} And B.ServiceType = '{1}'{2}", instAddrNo, dtCD046.Rows(0)("DependService"), CustStatusSQL)
                                '2019.02.25 by Corey RD JACKY通知大家，ExecNqry最好用在 Insert Update Delete，不要用在取資料用。
                                'If DAO.ExecNqry(strSQL) = 0 Then
                                '    ErrorMessage = ValidateLanguage.WipPRNotCancel
                                '    Exit Try
                                'End If
                                Using dtTmp As DataTable = DAO.ExecQry(strSQL)
                                    If dtTmp.Rows.Count = 0 Then
                                        ErrorMessage = ValidateLanguage.WipPRNotCancel
                                        Exit Try
                                    End If
                                End Using
                            End If
                            If ServiceInterdepend = 2 Then
                                Dim strSQL As String = String.Format("Select Count(*) From SO001 A,SO002 B Where A.CustId = B.CustId And A.CustId = {0} And B.ServiceType = '{1}'{2}", CustId, dtCD046.Rows(0)("DependService"), CustStatusSQL)
                                '2019.02.25 by Corey RD JACKY通知大家，ExecNqry最好用在 Insert Update Delete，不要用在取資料用。
                                'If DAO.ExecNqry(strSQL) = 0 Then
                                '    ErrorMessage = ValidateLanguage.WipPRNotCancel
                                '    Exit Try
                                'End If
                                Using dtTmp As DataTable = DAO.ExecQry(strSQL)
                                    If dtTmp.Rows.Count = 0 Then
                                        ErrorMessage = ValidateLanguage.WipPRNotCancel
                                        Exit Try
                                    End If
                                End Using
                            End If
                        End If
                    ElseIf PRRefNo = 3 Then
                        If dtCD046.Rows.Count > 0 Then
                            If Not String.IsNullOrEmpty(dtCD046.Rows(0)("DependService")) And dtCD046.Rows(0)("DependService").ToString.ToUpper <> ServiceType.ToString.ToUpper Then
                                If DAO.ExecNqry(_DAL.GetPRInterdepend, New Object() {CustId, MainSNO}) = 0 Then
                                    ErrorMessage = ValidateLanguage.WipPRNotCancel
                                    Exit Try
                                End If
                            End If
                        End If
                    End If
                End Using
            End If
        Catch ex As Exception
            Throw ex
        End Try
        If ErrorMessage <> String.Empty Then
            result.ResultBoolean = False
            result.ErrorCode = -99
            result.ErrorMessage = ErrorMessage
        Else
            result.ResultBoolean = True
            result.ErrorCode = 0
            result.ErrorMessage = String.Empty
        End If
        Return result
    End Function

    Public Function ChkDataOk(ByVal EditMode As EditMode, ByVal WipData As DataSet) As RIAResult
        Return ChkDataOk(EditMode, WipData, False)
    End Function

    ''' <summary>
    ''' 檢查派工單是否正確
    ''' </summary>
    ''' <param name="EditMode">工單狀態</param>
    ''' <param name="WipData">派工資料</param>
    ''' <param name="ShouldReg">是否順收</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChkDataOk(ByVal EditMode As EditMode, ByVal WipData As DataSet, ByVal ShouldReg As Boolean) As RIAResult
        Dim aRet As New RIAResult()
        Dim WorkCode As DataTable = Nothing
        Dim SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
        Dim WipUtilValidate As New CableSoft.SO.BLL.Wip.Utility.Validate(Me.LoginInfo, DAO)
        Try
            If Not WipData.Tables.Contains("SO001") Then
                Using dtSO001 As DataTable = DAO.ExecQry(_ValidateDAL.GetCust001, New Object() {WipData.Tables("Wip").Rows(0)("CustId")})
                    dtSO001.TableName = "SO001"
                    WipData.Tables.Add(dtSO001.Copy)
                End Using
            End If
            If Not WipData.Tables.Contains("SO002") Then
                Using dtSO002 As DataTable = DAO.ExecQry(_ValidateDAL.GetCust002, New Object() {WipData.Tables("Wip").Rows(0)("CustId"), WipData.Tables("Wip").Rows(0)("ServiceType")})
                    dtSO002.TableName = "SO002"
                    WipData.Tables.Add(dtSO002.Copy)
                End Using
            End If

            Dim PRWipData As DataTable = WipData.Tables("Wip")
            Dim PRChangeFacility As DataTable = WipData.Tables("ChangeFacility")
            Dim Facility As DataTable = WipData.Tables("Facility")
            Dim PrFacility As DataTable = WipData.Tables("PrFacility")
            Dim Charge As DataTable = WipData.Tables("Charge")
            Dim PRCustData1 As DataRow = WipData.Tables("SO001").Rows(0)
            Dim PRCustData2 As DataRow = WipData.Tables("SO002").Rows(0)
            Dim WipSystem As DataTable = SOUtil.GetSystem(CableSoft.SO.BLL.Utility.SystemTableType.Wip, "*", PRWipData.Rows(0).Item("ServiceType"))

            For Each row As DataRow In PRWipData.Rows
                Dim OldWip As DataTable = DAO.ExecQry(_DAL.GetWipData, row("SNo"), False)
                Dim WipCode As DataTable = SOUtil.GetCode(CableSoft.SO.BLL.Utility.CodeType.PRCode, row("PRCode").ToString, False)
                '必要欄位: 停/拆機類別,停/拆機原因,預約時間
                If row.IsNull("PRCode") Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = ValidateLanguage.colNotNullPRCode
                    Return aRet
                End If
                If row.IsNull("ReasonCode") Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = ValidateLanguage.colNotNullPRReason
                    Return aRet
                End If
                If row.IsNull("ResvTime") Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = ValidateLanguage.colNotNullResvTime
                    Return aRet
                End If
                Dim WipRefNo As Integer = 0
                Dim dtCD007tmp As DataTable = DAO.ExecQry(_DAL.GetCD007, row("PRCode"), False)
                If dtCD007tmp.Rows.Count > 0 Then
                    If Not dtCD007tmp.Rows(0).IsNull("RefNo") Then WipRefNo = dtCD007tmp.Rows(0)("RefNo")
                End If
                Select Case WipRefNo
                    Case 2, 5
                        'A.	呼叫 GetChangePRCode , 檢核是否需更換派工類別。
                        'B.	完工時間由未完工變成完工 或 變更完工時間 需做以下檢核:
                        '  .完工時間必須大於前次裝機時間(SO002.InstTime)
                        If Not row.IsNull("FinTime") Then
                            If row.Item("FinTime", DataRowVersion.Current).ToString <> row("FinTime", DataRowVersion.Original).ToString Then
                                If PRCustData2("InstTime") > row.Item("FinTime") Then
                                    aRet.ResultBoolean = False
                                    aRet.ErrorCode = -1
                                    aRet.ErrorMessage = ValidateLanguage.FinTimeError
                                    Return aRet
                                End If
                            End If
                        End If
                    Case 3
                        'A.必要欄位() : 新裝機地址, 新收費地址, 新郵寄地址, 新電話
                        'B.新裝機地址不可等於SO001.InstAddrNo()
                        If row.IsNull("ReInstAddrNo") OrElse row.IsNull("ReInstAddress") Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = ValidateLanguage.colNotNullReInstAddrNo
                            Return aRet
                        End If
                        '#8269 2019.04.03 因為畫面沒有 收費地址和郵寄地址 所以就不需要驗證是否有值的問題
                        'If row.IsNull("NewChargeAddrNo") OrElse row.IsNull("NewChargeAddress") Then
                        '    aRet.ResultBoolean = False
                        '    aRet.ErrorCode = -1
                        '    aRet.ErrorMessage = ValidateLanguage.colNotNullNewChargeAddrNo
                        '    Return aRet
                        'End If
                        'If row.IsNull("NewMailAddrNo") OrElse row.IsNull("NewMailAddress") Then
                        '    aRet.ResultBoolean = False
                        '    aRet.ErrorCode = -1
                        '    aRet.ErrorMessage = ValidateLanguage.colNotNullNewMailAddrNo
                        '    Return aRet
                        'End If
                        If row.IsNull("NewTel1") Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = ValidateLanguage.colNotNullNewTel1
                            Return aRet
                        End If
                        'If row.IsNull("NewTel2") Then
                        '    aRet.ResultBoolean = False
                        '    aRet.ErrorCode = -1
                        '    aRet.ErrorMessage = "必要欄位:新電話2"
                        '    Return aRet
                        'End If
                        'If row.IsNull("NewTel3") Then
                        '    aRet.ResultBoolean = False
                        '    aRet.ErrorCode = -1
                        '    aRet.ErrorMessage = "必要欄位:新電話3"
                        '    Return aRet
                        'End If
                        If row("OldAddrNo") = row("ReInstAddrNo") Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = ValidateLanguage.colNotNullOldAddrNo
                            Return aRet
                        End If
                        If row.IsNull("ReInstDate") OrElse row.IsNull("ReInstDate") Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = ValidateLanguage.colNotNullReInstDate
                            Return aRet
                        End If
                End Select

                If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                    For Each rowChange As DataRow In PRChangeFacility.Rows
                        Dim intChangeFacu As Int16 = DAO.ExecNqry(_DAL.chkPrChangeFacility, rowChange("SEQNO"), False)
                        If intChangeFacu > 0 Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = ValidateLanguage.WipPRHaveRemove
                            Return aRet
                        End If
                    Next
                End If

                If row("FinTime").ToString <> String.Empty Then
                    '(5)當完工時間(Wip.FinTime)有值時, 需做以下檢核:
                    'A.	關聯工單已退單,此工單不得完工: (Select Count(*) From SO009 Where MainSNo = <Wip.SNo> And CustId = <客戶編號> And ServiceType = <服務別> And ReturnCode is not null ) >0
                    'B.	完工時間必須大於前次裝機時間(SO002.InstTime)。
                    'C.	當該派工類別為同區移機(CD007.ReInstAcrossFlag>0)時, 需檢核同區移機中介檔移裝狀態: (Select NStatus From (SO041.ReInstOwner).SO313 Where OCustId = <客戶編號> And OSNo = <工單單號> And OCompCode = <公司別> ) = “退單”, 則不能做完工。
                    Dim chkMainSNO As Int16 = DAO.ExecNqry(_DAL.chkWipPRMainSNO, New Object() {row("SNO"), row("Custid"), row("ServiceType")})
                    If chkMainSNO > 0 Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -1
                        'aRet.ErrorMessage = "關聯工單已退單,此工單不得完工！"
                        aRet.ErrorMessage = ValidateLanguage.OAddressReturnCannotAccept
                        Return aRet
                    End If
                    '2018.10.18 DEBBY 和JACKY討論後，說明是 當初文件寫錯，只有參考號 2,5 才需要做，所以下面這段檢核是多餘的。
                    'If PRCustData2("InstTime") > row.Item("FinTime") Then
                    '    aRet.ResultBoolean = False
                    '    aRet.ErrorCode = -1
                    '    aRet.ErrorMessage = ValidateLanguage.FinTimeError
                    '    Return aRet
                    'End If
                    Dim ReInstAcrossFlag As Int16 = DAO.ExecQry(_DAL.GetCD007, row("PRCode"), False).Rows(0)("ReInstAcrossFlag")
                    If ReInstAcrossFlag > 0 Then
                        Dim ReInstOwner As String = DAO.ExecQry(_DAL.GetSO041).Rows(0)("ReInstOwner").ToString
                        If ReInstOwner.ToString <> String.Empty Then ReInstOwner = ReInstOwner & "."
                        If DAO.ExecQry(_DAL.chkReInstAcross(ReInstOwner), New Object() {row("CustId"), row("SNO"), row("CompCode")}).Rows(0)(0).ToString = "退單" Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            'aRet.ErrorMessage = "派工類別為同區移機，同區移機中介檔移裝狀態退單，則不能做完工！"
                            aRet.ErrorMessage = ValidateLanguage.OAddressReturnCannotFinish
                            Return aRet
                        End If
                    End If
                End If

                If row("ReturnCode").ToString <> String.Empty Then
                    '(6)當退單原因(ReturnCode)有值時, 需做以下檢核:
                    'A.	關聯工單已完工,此工單不得退單: (Select Count(*) From SO009 Where MainSNo = <Wip.SNo> And CustId = <客戶編號> And ServiceType = <服務別> And FinTime is not null ) >0。
                    'B.	當該派工類別為同區移機(CD007.ReInstAcrossFlag>0)時, 需檢核同區移機中介檔移裝狀態: (Select NStatus From (SO041.ReInstOwner).SO313 Where OCustId = <客戶編號> And OSNo = <工單單號> And OCompCode = <公司別> ) = “完工”, 則不能做退單。
                    Dim chkMainSNO As Int16 = DAO.ExecNqry(_DAL.chkWipPRFinTime, New Object() {row("SNO"), row("Custid"), row("ServiceType")})
                    If chkMainSNO > 0 Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -1
                        aRet.ErrorMessage = ValidateLanguage.OtherWipFintTime
                        Return aRet
                    End If
                    Dim ReInstAcrossFlag As Int16 = DAO.ExecQry(_DAL.GetCD007, row("PRCode"), False).Rows(0)("ReInstAcrossFlag")
                    If ReInstAcrossFlag > 0 Then
                        Dim ReInstOwner As String = DAO.ExecQry(_DAL.GetSO041).Rows(0)("ReInstOwner").ToString
                        If ReInstOwner.ToString <> String.Empty Then ReInstOwner = ReInstOwner & "."
                        If DAO.ExecQry(_DAL.chkReInstAcross(ReInstOwner), New Object() {row("CustId"), row("SNO"), row("CompCode")}).Rows(0)(0).ToString = "完工" Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            'aRet.ErrorMessage = "派工類別為同區移機，同區移機中介檔移裝狀態完工，則不能做退單！"
                            aRet.ErrorMessage = ValidateLanguage.OAddressReturnCannotReturn
                            Return aRet
                        End If
                    End If
                End If

                '共用檢核
                Dim result As RIAResult = WipUtilValidate.ChkDataOk(EditMode, 2, PRWipData, OldWip, WipCode, Facility, PrFacility, Charge, Nothing, WipSystem, PRChangeFacility, ShouldReg)
                '設備檢核錯誤碼    : -10001 ~ -19999
                '拆設備檢核錯誤碼  : -20001 ~ -30000
                '收費檢核錯誤碼    : -30001 ~ -40000
                '指定設備檢核錯誤碼: -40001 ~ -50000
                If result.ResultBoolean = False Then
                    Return result
                End If
                aRet.ResultBoolean = True
                aRet.ErrorCode = 0
                aRet.ErrorMessage = String.Empty
            Next
        Catch ex As Exception
            aRet.ErrorCode = -999
            aRet.ErrorMessage = ex.ToString & "--" & ValidateLanguage.DataUpdateError
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 同區移機單、檢核其它單的狀態(因為ChkDataOK 已經有檢核了，所以多寫了)
    ''' </summary>
    ''' <param name="EditMode"></param>
    ''' <param name="Wip"></param>
    ''' <param name="WipCode"></param>
    ''' <param name="SOUtil"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ChkReInstAcrossProcess(ByVal EditMode As EditMode, _
                                            ByVal Wip As DataTable, ByVal WipCode As DataTable, ByVal SOUtil As CableSoft.SO.BLL.Utility.Utility) As RIAResult
        Try
            '檢查共用區移機資料
            '7.  不管是移拆或移裝工單，只要有一單做退單處理，則另一單亦須做退單處理；
            '反之有一單做了完工處理，則另一單亦須做完工處理流程
            Dim ErrorMessage As String = ""
            If WipCode.Rows(0).Item("ReInstAcrossFlag") > 0 Then
                Dim COMOwner As String = SOUtil.GetSystem(CableSoft.SO.BLL.Utility.SystemTableType.System, "ReInstOwner", "").Rows(0).Item(0).ToString
                Using ComInterface As DataTable = DAO.ExecQry(_ValidateDAL.GetCOMInterface(COMOwner), New Object() {Wip.Rows(0).Item("CustId"), Wip.Rows(0).Item("SNo"), Wip.Rows(0).Item("CompCode")})
                    If ComInterface.Rows.Count > 0 Then
                        With ComInterface.Rows(0)
                            If Not .IsNull("NSNOSTATUS") AndAlso .Item("NSNOSTATUS") = ValidateLanguage.WipRunStatus0 AndAlso EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                                '新地址移入單已退單,不可派工!!
                                Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = 0, .ErrorMessage = ValidateLanguage.OAddressReturnCannotAccept}
                            End If
                            If Not .IsNull("NSNOSTATUS") AndAlso .Item("NSNOSTATUS") = ValidateLanguage.WipRunStatus0 AndAlso Not Wip.Rows(0).IsNull("FinTime") Then
                                '新地址移入單已退單,移拆單只可做退單不可完工!!
                                Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = 0, .ErrorMessage = ValidateLanguage.OAddressReturnCannotFinish}
                            End If
                            If Not .IsNull("NSNOSTATUS") AndAlso .Item("NSNOSTATUS") = ValidateLanguage.WipRunStatus1 AndAlso Not Wip.Rows(0).IsNull("ReturnCode") Then
                                '新地址移入單已完工,移拆單只可做完工不可退單!!
                                Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = 0, .ErrorMessage = ValidateLanguage.OAddressReturnCannotReturn}
                            End If
                            '9.  若已於新址做CM開關機動作後工單卻被操作者做退單時，
                            '則之後之CM重開關機動作需由客戶自行至單獨開關機機制中處理，程式將不自動做此步驟。
                            '但程式將show訊息提示操作者"該設備已做過CM開關機動作! 現在退單~請記得自行重做CM開關機，
                            '以利設備順利使用"
                            Dim Count As Integer = Integer.Parse(DAO.ExecQry(_ValidateDAL.ChkMustReOpenCommand(), New Object() {.Item("NSNO")}).Rows(0).Item(0))
                            If Not Wip.Rows(0).IsNull("ReturnCode") AndAlso Count > 0 Then
                                ErrorMessage = ValidateLanguage.OAddressMustReOpenCommand
                            End If
                        End With
                    End If
                End Using
            End If
            Return New RIAResult() With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = ErrorMessage}
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

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
            Catch ex As Exception
            End Try
            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    Protected Overrides Sub Finalize()
        ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
