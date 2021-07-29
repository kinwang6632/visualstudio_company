Option Compare Binary
Option Infer On
Option Explicit On

Imports CableSoft.BLL.Utility
Imports WipPRLanguage = CableSoft.BLL.Language.SO31.WipPRLanguage

Public Class PR
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New PRDAL(Me.LoginInfo.Provider)

    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
    ''' <summary>
    ''' 可新增
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanAppend(ByVal CustId As Integer, ByVal ServiceType As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11131")

            If Not aRet.ResultBoolean Then
                Return aRet
            Else
                If ServiceType = "X" OrElse String.IsNullOrEmpty(ServiceType) Then
                    '沒有傳入服務別或是傳入X，需要判斷所有的狀態種類，所有服務別都不能派拆才秀訊息，其中一個服務別可以派拆就要可以進入工單
                    Dim CanSerViceType As String = String.Empty
                    Dim CanUseRefNo As String = String.Empty
                    Dim ErrCode As Integer = 0
                    Dim ErrMessage As String = String.Empty
                    Using dt As DataTable = DAO.ExecQry(_DAL.GetCustomer(ServiceType), CustId, False)
                        For Each dr As DataRow In dt.Rows
                            If CanAppendChk(String.Format("0{0}", dr.Item("CustStatusCode")), dr.Item("ServiceType"), ErrCode, ErrMessage) Then
                                If ErrCode >= 0 Then
                                    CanSerViceType = String.Format("{0},{1}", CanSerViceType, dr("ServiceType"))
                                    CanUseRefNo = String.Format("{0}+{1}", CanUseRefNo, String.Format("{0}-{1}", dr("ServiceType"), ErrMessage))
                                End If
                            Else
                                aRet.ResultBoolean = False
                                aRet.ErrorCode = -99
                                aRet.ErrorMessage = WipPRLanguage.ChkCustStatusError
                                Exit For
                            End If
                        Next
                        If CanSerViceType = String.Empty Then
                            '沒有服務別資料表是所有服務別都是拆機的狀態須回傳最後的ErrorCode,ErrorMessage
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = ErrCode
                            aRet.ErrorMessage = ErrMessage
                        Else
                            '有正常服務別的客戶資料需要回傳
                            aRet.ResultBoolean = True
                            aRet.ErrorCode = 0
                            aRet.ErrorMessage = ""
                            aRet.ResultXML = String.Format("CanServiceType={0};CanRefNo={1}", CanSerViceType.Substring(1), CanUseRefNo.Substring(1))
                        End If
                    End Using
                Else
                    '有傳入服務別，可針對服務別來判斷
                    Using dt As DataTable = DAO.ExecQry(_DAL.GetCustomer(ServiceType), New Object() {CustId})
                        Dim ErrCode As Integer = 0
                        Dim ErrMessage As String = String.Empty
                        For Each dr As DataRow In dt.Rows
                            If CanAppendChk(String.Format("0{0}", dr.Item("CustStatusCode")), dr.Item("ServiceType"), ErrCode, ErrMessage) Then
                                If ErrCode < 0 Then
                                    aRet.ResultBoolean = False
                                    aRet.ErrorCode = ErrCode
                                    aRet.ErrorMessage = ErrMessage
                                    Exit For
                                End If
                            Else
                                aRet.ResultBoolean = False
                                aRet.ErrorCode = -99
                                aRet.ErrorMessage = WipPRLanguage.ChkCustStatusError
                                Exit For
                            End If
                        Next
                    End Using
                End If
            End If
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function

    Private Function CanAppendChk(ByVal CustStatus As Integer, ByVal ServiceType As String,
                                  ByRef ErrorCode As Integer, ByRef ErrorMessage As String) As Boolean
        ErrorCode = 0
        ErrorMessage = String.Empty
        'Select Case CustStatus
        '    Case 1
        '    Case 4
        '        ErrorCode = -4
        '        ErrorMessage = "註銷戶無法產生派工單！"
        '    Case 5
        '        Using dt2 As DataTable = DAO.ExecQry(_DAL.GetSO042, New Object() {ServiceType})
        '            For Each dr2 As DataRow In dt2.Rows
        '                If IsDBNull(dr2("AbnormalFaci")) OrElse (Int32.Parse(dr2.Item("AbnormalFaci").ToString) = 0) Then
        '                    ErrorCode = 5
        '                    'ErrorMessage = "可傳派工類別: 15"
        '                    'ErrorMessage = "15"
        '                Else
        '                    ErrorCode = 5
        '                    'ErrorMessage = "可派的派工類別: 6,8,10,15"
        '                    'ErrorMessage = "6,8,10,15"
        '                End If
        '            Next
        '        End Using
        '    Case Else
        '        ErrorCode = 99
        '        'ErrorMessage = "可派的派工參考號: 6,7,9,10,15"
        '        ErrorMessage = "6,7,9,10,15"
        'End Select

        '2012.07.02 和Jacky,Karen討論後，未來用 by機 所以不需要特別判斷CustStatusCode,只需要判斷=4就可以
        If CustStatus = 4 Then
            ErrorCode = -4
            ErrorMessage = WipPRLanguage.ChkCustIsCancel
        End If

        Return True
    End Function

    ''' <summary>
    ''' 可修改
    ''' </summary>
    ''' <param name="PRSNO">拆機工單號碼</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanEdit(ByVal PRSNO As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11132")
            If Not aRet.ResultBoolean Then
                Return aRet
            End If
            Using dtWip As DataTable = DAO.ExecQry(_DAL.CanWipData, PRSNO, False)
                If dtWip.Rows.Count <= 0 Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = WipPRLanguage.WipPRDataNothing
                Else
                    If Not dtWip.Columns.Contains("ClsTime") Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -1
                        aRet.ErrorMessage = WipPRLanguage.colClsTimeNothing
                    Else
                        If Not dtWip.Rows(0).IsNull("ClsTime") Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = WipPRLanguage.colClsTimeIsNotNull
                        End If
                    End If
                End If
            End Using
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 可作廢
    ''' </summary>
    ''' <param name="PRSNO">拆機工單號碼</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanDelete(ByVal PRSNO As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11133")
            If Not aRet.ResultBoolean Then
                Return aRet
            End If
            Using dtWip As DataTable = DAO.ExecQry(_DAL.CanWipData, PRSNO, False)
                If dtWip.Rows.Count <= 0 Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = WipPRLanguage.WipPRDataNothing
                Else
                    If Not dtWip.Columns.Contains("ClsTime") Then
                        aRet.ResultBoolean = False
                        aRet.ErrorCode = -1
                        aRet.ErrorMessage = WipPRLanguage.colClsTimeNothing
                    Else
                        If Not dtWip.Rows(0).IsNull("ClsTime") Then
                            aRet.ResultBoolean = False
                            aRet.ErrorCode = -1
                            aRet.ErrorMessage = WipPRLanguage.colClsTimeIsCancel
                        End If
                    End If
                End If
            End Using
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 可顯示
    ''' </summary>
    ''' <param name="PRSNO">拆機工單號碼</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanView(ByVal PRSNO As String) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11134")
            If Not aRet.ResultBoolean Then
                Return aRet
            End If
            Using dtWip As DataTable = DAO.ExecQry(_DAL.CanWipData, PRSNO, False)
                If dtWip.Rows.Count <= 0 Then
                    aRet.ResultBoolean = False
                    aRet.ErrorCode = -1
                    aRet.ErrorMessage = WipPRLanguage.WipPRDataNothing
                    'Else
                    '    If Not dtWip.Columns.Contains("ClsTime") Then
                    '        aRet.ResultBoolean = False
                    '        aRet.ErrorCode = -1
                    '        aRet.ErrorMessage = "無日結欄位可判斷！"
                    '    Else
                    '        If Not dtWip.Rows(0).IsNull("ClsTime") Then
                    '            aRet.ResultBoolean = False
                    '            aRet.ErrorCode = -1
                    '            aRet.ErrorMessage = "已日結不可作廢資料！"
                    '        End If
                    '    End If
                End If
            End Using
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function
    ''' <summary>
    ''' 可列印
    ''' </summary>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanPrint() As RIAResult
        Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO).ChkPriv(Me.LoginInfo.EntryId, "SO11135")
    End Function
    ''' <summary>
    ''' 取得所有權限
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetPriv() As DataTable
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Dim dt As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO1113")
            Return dt
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得一般工單資訊
    ''' </summary>
    ''' <param name="InstCode">派工類別</param>
    ''' <returns>DataSet</returns>
    ''' <remarks>Facility,Charge,ChangeFacility</remarks>
    Public Function GetNormalWip(ByVal CustId As Int32,
                                 ByVal ServiceType As String,
                                 ByVal ResvTime As Date,
                                 ByVal InstCode As Int32) As DataSet
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Return (obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, CustId, ServiceType, ResvTime, InstCode))
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得服務別資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetServiceType(ByVal CanUseServiceType As String) As DataTable
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetServiceType(CanUseServiceType))
        Return dt
    End Function
    ''' <summary>
    ''' 畫面起始載入
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFromDefaultLoad(ByVal CustId As Integer, ByVal ServiceType As String,
                                       ByVal ServCode As String, ByVal CanUseRefNo As String,
                                       ByVal WipCodeValueStr As String, ByVal blnReInstFilter As Boolean,
                                       ByVal ReInstAcrossFlag As Boolean, ByVal SNO As String) As DataSet
        Try
            Dim dsDefaultLoad As New DataSet
            'GetPrData
            Using dsPRData As DataSet = GetPRData(SNO, CustId, ServiceType) 'Wip,SO001,SO002
                If CustId <= 0 Then
                    If dsPRData.Tables.Contains("Wip") Then
                        If dsPRData.Tables("Wip").Rows.Count > 0 Then
                            CustId = dsPRData.Tables("Wip").Rows(0)("Custid")
                            ServiceType = dsPRData.Tables("Wip").Rows(0)("ServiceType")
                        End If
                    End If
                End If
                For Each dtGet As DataTable In dsPRData.Tables
                    dsDefaultLoad.Tables.Add(dtGet.Copy)
                Next
            End Using

            Using dtServiceType As DataTable = GetServiceType(String.Empty)
                dtServiceType.TableName = "ServiceType"
                dsDefaultLoad.Tables.Add(dtServiceType.Copy)
            End Using

            'PRCode
            Using dtPrCode As DataTable = GetPRCode(ServiceType, CanUseRefNo, WipCodeValueStr, blnReInstFilter, ReInstAcrossFlag)
                dtPrCode.TableName = "PrCode"
                dsDefaultLoad.Tables.Add(dtPrCode.Copy)
            End Using
            'PRReasonCode(3)
            Using dtPRReasonCode As DataTable = GetPRReasonCode(ServiceType)
                dtPRReasonCode.TableName = "PRReasonCode"
                dsDefaultLoad.Tables.Add(dtPRReasonCode.Copy)
            End Using
            'ReturnCode(4)
            Using dtReturnCode As DataTable = GetReturnCode(ServiceType)
                dtReturnCode.TableName = "ReturnCode"
                dsDefaultLoad.Tables.Add(dtReturnCode.Copy)
            End Using
            'ReturnDescCode(5)
            Using dtReturnDescCode As DataTable = GetReturnDescCode(ServiceType)
                dtReturnDescCode.TableName = "ReturnDescCode"
                dsDefaultLoad.Tables.Add(dtReturnDescCode.Copy)
            End Using
            'GroupCode(6, ServCode)
            Using dtGroupCod As DataTable = GetGroupCode(ServCode)
                dtGroupCod.TableName = "GroupCod"
                dsDefaultLoad.Tables.Add(dtGroupCod.Copy)
            End Using
            'SignEn(7)
            Using dtSignEn As DataTable = GetSignEn()
                dtSignEn.TableName = "SignEn"
                dsDefaultLoad.Tables.Add(dtSignEn.Copy)
            End Using
            'SatiCode(8)
            Using dtSatiCode As DataTable = GetSatiCode(ServiceType)
                dtSatiCode.TableName = "SatiCode"
                dsDefaultLoad.Tables.Add(dtSatiCode.Copy)
            End Using
            'WorkerEn(9, 0)
            Using dtWorkerEn As DataTable = GetWorkerEn(0)
                dtWorkerEn.TableName = "WorkerEn0"
                dsDefaultLoad.Tables.Add(dtWorkerEn.Copy)
            End Using
            'WorkerEn(10, 1)
            Using dtWorkerEn As DataTable = GetWorkerEn(1)
                dtWorkerEn.TableName = "WorkerEn1"
                dsDefaultLoad.Tables.Add(dtWorkerEn.Copy)
            End Using
            'CustRtnCode(11)
            Using dtCustRtnCode As DataTable = GetCustRtnCode(ServiceType)
                dtCustRtnCode.TableName = "CustRtnCode"
                dsDefaultLoad.Tables.Add(dtCustRtnCode.Copy)
            End Using
            'SO042(12)
            Using dtSO042 As DataTable = GetSO042(ServiceType)
                dtSO042.TableName = "SO042"
                dsDefaultLoad.Tables.Add(dtSO042.Copy)
            End Using
            'SO041(13)
            Using dtSO041 As DataTable = DAO.ExecQry(_DAL.GetSO041)
                dtSO041.TableName = "SO041"
                dsDefaultLoad.Tables.Add(dtSO041.Copy)
            End Using
            'GetPriv(14)
            Using dtGetPriv As DataTable = GetPriv()
                dtGetPriv.TableName = "GetPriv"
                dsDefaultLoad.Tables.Add(dtGetPriv.Copy)
            End Using
            'GetUserPriv(15)
            Using dtGetUserPriv As DataTable = GetUserPriv()
                dtGetUserPriv.TableName = "GetUserPriv"
                dsDefaultLoad.Tables.Add(dtGetUserPriv.Copy)
            End Using
            'GetAddressData(16)
            Using dtGetAddressData As DataTable = GetAddressData(CustId)
                dtGetAddressData.TableName = "GetAddressData"
                dsDefaultLoad.Tables.Add(dtGetAddressData.Copy)
            End Using

            Return dsDefaultLoad
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選停拆機類別(GetPRCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="CanUseRefNo">可使用之參考號</param>
    ''' <param name="WipCodeValueStr">可使用之工單代碼</param>
    ''' <param name="blnReInstFilter">是否過濾 移機跨區</param>
    ''' <param name="ReInstAcrossFlag">移機跨區</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPRCode(ByVal ServiceType As String, ByVal CanUseRefNo As String,
                              ByVal WipCodeValueStr As String, ByVal blnReInstFilter As Boolean, ByVal ReInstAcrossFlag As Boolean) As DataTable
        Try
            '#6503 2013.05.03 增加判斷SO042 過濾不要的參考號
            Dim CanNotUseRefNo As String = String.Empty
            Using dtSO042 As DataTable = DAO.ExecQry(_DAL.GetSO042, New Object() {ServiceType})
                If dtSO042.Rows.Count > 0 Then
                    If Not dtSO042.Rows(0).IsNull("FilterPRRefNo") Then
                        CanNotUseRefNo = dtSO042.Rows(0)("FilterPRRefNo")
                    End If
                End If
            End Using
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRCode(CanUseRefNo, CanNotUseRefNo, WipCodeValueStr, blnReInstFilter, ReInstAcrossFlag), ServiceType, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選停拆機類別(GetPRCodeByContactRefNo)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="ContactRefno">互動管理參考號</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetPRCodeByContactRefNo(ByVal ServiceType As String, ByVal ContactRefno As Integer) As DataTable
        Try
            Dim RefNo As String = String.Empty
            '(1)	ContactRefNo = 4:移機 則 派工參考號過濾3
            '(2)	ContactRefNo = 5:拆機 則 派工參考號過濾2,5,6,8
            '(3)	ContactRefNo = 6:停機則 派工參考號過濾1,11
            '(4)	ContactRefNo = 11:拆設備則 派工參考號過濾6,8
            '(5)	ContactRefNo = 27,28:同/跨區移機則 派工參考號過濾2,6,8
            '(6)	ContactRefNo = 30:關機 則 派工參考號過濾7
            '(7)	ContactRefNo = 32,33:關機 則 派工參考號過濾10
            '(8)	ContactRefNo = 30:關機 則 派工參考號過濾7
            '(9)	ContactRefNo = 39:暫停頻道 則 派工參考號過濾15
            '(10)	ContactRefNo = 42:暫停頻道 則 派工參考號過濾14
            Select Case ContactRefno
                Case 4
                    RefNo = "3"
                Case 5
                    RefNo = "2,5,6,8"
                Case 6
                    RefNo = "1,11"
                Case 11
                    RefNo = "6,8"
                Case 27, 28
                    RefNo = "2,6,8"
                Case 30
                    RefNo = "7"
                Case 32, 33
                    RefNo = "10"
                Case 39
                    RefNo = "15"
                Case 42
                    RefNo = "14"
                Case Else
                    RefNo = ""
            End Select
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRCodeByContactRefNo, New Object() {ServiceType, RefNo})
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選停拆移機原因(GetPRReasonCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetPRReasonCode(ByVal ServiceType As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRReasonCode, ServiceType, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選停拆移機原因(GetPRReasonDescCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="PRReasonCode">停拆移機原因</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetPRReasonDescCode(ByVal ServiceType As String, ByVal PRReasonCode As Integer) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRReasonDescCode, New Object() {ServiceType, PRReasonCode})
            dt.TableName = "ReasonDescCode"
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選工程組別(GetGroupCode)
    ''' </summary>
    ''' <param name="ServCode">服務區</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetGroupCode(ByVal ServCode As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetGroupCode, ServCode, False)
            If dt.Rows.Count <= 0 Then dt = DAO.ExecQry("Select CodeNo,Description,RefNo From CD003 Where StopFlag = 0 Order by CodeNo")
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選工作人員(GetWorkerEn)
    ''' </summary>
    ''' <param name="Type">工程人員種類 (0:工程人員1,1:工程人員2)</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetWorkerEn(ByVal Type As Integer) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetWorkerEn)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選退單原因(GetReturnCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetReturnCode(ByVal ServiceType As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetReturnCode, ServiceType, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選退單原因分類(GetReturnDescCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetReturnDescCode(ByVal ServiceType As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetReturnDescCode, String.Format("%{0}%", ServiceType), False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選簽收人員(GetSignEn)
    ''' </summary>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetSignEn() As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetSignEn)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得可選服務滿意度(GetSatiCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetSatiCode(ByVal ServiceType As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetSatiCode, ServiceType, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得客戶流向(GetCustRtnCode)
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetCustRtnCode(ByVal ServiceType As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetCustRtnCode, ServiceType, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得維修資料
    ''' </summary>
    ''' <param name="SNo">工單單號</param>
    ''' <returns>DataSet</returns>
    ''' <remarks></remarks>
    Public Function GetPRData(ByVal SNo As String) As DataSet
        Try
            Return GetPRData(SNo, 0, String.Empty)
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function GetPRData(ByVal SNo As String, ByVal Custid As Integer, ByVal ServiceType As String) As DataSet
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Try
            Dim dsPRData As DataSet = Nothing
            dsPRData = obj.GetWipDetail(SNo, False, BLL.Utility.InvoiceType.PR)
            If Custid <= 0 Then
                Custid = dsPRData.Tables("Wip").Rows(0)("Custid")
                ServiceType = dsPRData.Tables("Wip").Rows(0)("ServiceType")
            End If
            Using dtSO001 As DataTable = DAO.ExecQry(_DAL.GetSO001, Custid, False)
                If dtSO001 IsNot Nothing Then
                    dtSO001.TableName = "SO001"
                    dsPRData.Tables.Add(dtSO001.Copy)
                End If
            End Using
            Using dtSO002 As DataTable = DAO.ExecQry(_DAL.GetSO002(ServiceType), New Object() {Custid})
                If dtSO002 IsNot Nothing Then
                    dtSO002.TableName = "SO002"
                    dsPRData.Tables.Add(dtSO002.Copy)
                End If
            End Using

            Return dsPRData
        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得客戶主檔
    ''' </summary>
    ''' <param name="Custid">客戶編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCustOmer(ByVal ServiceType As String, ByVal Custid As Integer) As DataSet
        Dim ds As New DataSet()
        Try
            Using dtSO001 As DataTable = DAO.ExecQry(_DAL.GetSO001, Custid, False)
                If dtSO001 IsNot Nothing Then
                    dtSO001.TableName = "SO001"
                    ds.Tables.Add(dtSO001.Copy)
                End If
            End Using
            Using dtSO002 As DataTable = DAO.ExecQry(_DAL.GetSO002(ServiceType), New Object() {Custid})
                If dtSO002 IsNot Nothing Then
                    dtSO002.TableName = "SO002"
                    ds.Tables.Add(dtSO002.Copy)
                End If
            End Using
        Catch ex As Exception
            Throw ex
        Finally
            ds.Dispose()
        End Try
        Return ds
    End Function
    ''' <summary>
    ''' 取得轉換派工類別(GetChangePRCode)
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="PRRefNo">派工類別參考號</param>
    ''' <returns>Collection</returns>
    ''' <remarks></remarks>
    Public Function GetChangePRCode(ByVal CustId As Integer, ByVal ServiceType As String, ByVal PRRefNo As Integer) As DataTable
        Try
            Dim dtRtn As DataTable = Nothing
            Select Case PRRefNo
                Case 2, 5, 6
                    Dim Refno As String = String.Empty
                    Select Case ServiceType.ToUpper
                        Case "C"
                            Refno = "10"
                        Case "D"
                            Refno = "3"
                        Case "I"
                            Refno = "2,5,7,8"
                        Case "P"
                            Refno = "6"
                    End Select
                    Dim Count004 As Int16 = DAO.ExecNqry(_DAL.GetChangePRCode, New Object() {CustId, Refno})
                    If Count004 = 0 Then
                        dtRtn = DAO.ExecQry(_DAL.GetPRCodeByContactRefNo, New Object() {ServiceType, "2,5"})
                    Else
                        dtRtn = DAO.ExecQry(_DAL.GetPRCodeByContactRefNo, New Object() {ServiceType, "6"})
                    End If
            End Select
            Return dtRtn
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 檢核結清資料是否能修改
    ''' </summary>
    ''' <returns>RiaResult</returns>
    ''' <remarks></remarks>
    Public Function GetSO1132Priv() As RIAResult
        Try
            Return New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO).ChkPriv(Me.LoginInfo.EntryId, "SO1132C")
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得SO042工單設定
    ''' </summary>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSO042(ByVal ServiceType As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetSO042, ServiceType, False)
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetUserPriv() As DataTable
        Try
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                Dim dt As DataTable = SOUtil.GetPriv(LoginInfo.EntryId, New String() {"SO1113", "SO1113"}, LoginInfo.GroupId)
                Return dt
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ''' <summary>
    ''' 取得地址資料
    ''' </summary>
    ''' <param name="Custid">客戶編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAddressData(ByVal Custid) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetAddressData, New Object() {Custid, LoginInfo.CompCode})
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 取得一般工單資料
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="WorkCodeValue">派工類別</param>
    ''' <param name="ResvTime">預約時間</param>
    ''' <param name="Full">完整</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNormalCalculateData(ByVal CustId As Integer,
                                           ByVal ServiceType As String,
                                           ByVal WorkCodeValue As Integer,
                                           ByVal ResvTime As DateTime,
                                           ByVal SNo As String,
                                           ByVal Full As Boolean) As DataSet
        Return GetNormalCalculateData(CustId, ServiceType, WorkCodeValue, ResvTime, SNo, Full, Nothing)
    End Function
    ''' <summary>
    ''' 取得一般工單資料
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="WorkCodeValue">派工類別</param>
    ''' <param name="ResvTime">預約時間</param>
    ''' <param name="Full">完整</param>
    ''' <param name="WipOtherData">Change:收費資料 Contact:互動資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNormalCalculateData(ByVal CustId As Integer,
                                           ByVal ServiceType As String,
                                           ByVal WorkCodeValue As Integer,
                                           ByVal ResvTime As DateTime,
                                           ByVal SNo As String,
                                           ByVal Full As Boolean,
                                           ByVal WipOtherData As DataSet) As DataSet
        Try
            Dim WipData As DataSet = Nothing
            Dim dtChangeData As DataTable = Nothing
            Dim dtContact As DataTable = Nothing
            Using WipUtil As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
                If WipOtherData IsNot Nothing Then
                    If WipOtherData.Tables.Contains("Contact") Then
                        dtContact = WipOtherData.Tables("Contact").Copy
                    End If
                End If
                '取得收費/設備資料
                WipData = WipUtil.GetWipCalculateData(BLL.Utility.InvoiceType.PR, CustId, ServiceType, SNo, ResvTime, WorkCodeValue, dtContact)
                Dim Wip As DataTable = WipData.Tables("Wip")
                '取得工單資料
                If Wip.Rows.Count <= 0 Then
                    If Not GetNormalWip(CustId, ServiceType, SNo, ResvTime, WorkCodeValue, 0, Wip) Then
                        Throw New Exception("GetNormalWip")
                    End If
                End If
                If Full Then
                    '取得指定設備資料
                    Using Charge As DataTable = WipData.Tables("Charge")
                        Using Facility As DataTable = WipData.Tables("Facility")
                            Dim ChangeFacility As DataTable = WipData.Tables("ChangeFacility")
                            If Not GetChangeFaci(CustId, SNo, Wip, Charge, Facility, ChangeFacility) Then
                                Throw New Exception("GetChangeFaci")
                            End If
                            If WipOtherData IsNot Nothing Then
                                If WipOtherData.Tables.Contains("ChangeData") Then
                                    dtChangeData = WipOtherData.Tables("ChangeData").Copy
                                End If
                            End If
                            If Not WipUtil.GetDefaultChangeFaci(BLL.Utility.InvoiceType.PR, Wip, dtChangeData, ChangeFacility) Then
                                Throw New Exception("GetDefaultChangeFaci")
                            End If
                        End Using
                    End Using

                    Using dsCust As DataSet = GetCustOmer(ServiceType, CustId)
                        For Each dtTmp As DataTable In dsCust.Tables
                            If WipData.Tables.Contains(dtTmp.TableName) Then
                                WipData.Tables.Remove(dtTmp.TableName)
                            End If
                            WipData.Tables.Add(dtTmp.Copy)
                        Next
                    End Using
                    Using dtPriv As DataTable = GetPriv()
                        dtPriv.TableName = "Priv"
                        WipData.Tables.Add(dtPriv.Copy)
                    End Using
                End If
                WipData.AcceptChanges()
            End Using
            Return WipData
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '產生一般工單
    Private Function GetNormalWip(ByVal CustId As Integer, ByVal ServiceType As String,
                                  ByVal SNo As String,
                                  ByVal ResvTime As DateTime, ByVal WorkCodeValue As Integer,
                                  ByVal WorkUnit As Decimal,
                                  ByRef Wip As DataTable) As Boolean
        Try
            Dim WipRow As DataRow = Wip.NewRow
            Dim Address As DataTable = DAO.ExecQry(_DAL.GetAddressData(), New Object() {CustId, LoginInfo.CompCode})
            Dim SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
            Dim PRCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, WorkCodeValue.ToString, False)
            Dim GroupCode As DataTable = DAO.ExecQry(_DAL.GetGroupCode(), New Object() {Address.Rows(0).Item("ServCode")})
            Dim dtCustId As DataTable = DAO.ExecQry(_DAL.GetSO001, New Object() {CustId})
            With WipRow
                .Item("CustId") = CustId
                If dtCustId.Rows.Count > 0 Then
                    '#6933 2015.02.10 測試不OK 增加CustName,Tel1 兩個欄位
                    .Item("CustName") = dtCustId.Rows(0).Item("CUSTNAME")
                    .Item("Tel1") = dtCustId.Rows(0).Item("Tel1")
                End If
                .Item("SNo") = SNo
                .Item("OldAddrNo") = Address.Rows(0).Item("AddrNo")
                .Item("OldAddress") = Address.Rows(0).Item("Address")
                .Item("PRCode") = PRCode.Rows(0).Item("CodeNo")
                .Item("PRName") = PRCode.Rows(0).Item("Description")
                .Item("ResvTime") = ResvTime
                .Item("AcceptTime") = DateTime.Now
                .Item("AcceptEn") = LoginInfo.EntryId
                .Item("AcceptName") = LoginInfo.EntryName
                .Item("ServCode") = Address.Rows(0).Item("ServCode")
                .Item("StrtCode") = Address.Rows(0).Item("StrtCode")
                If GroupCode.Rows.Count > 0 Then
                    .Item("GroupCode") = GroupCode.Rows(0).Item("CodeNo")
                    .Item("GroupName") = GroupCode.Rows(0).Item("Description")
                End If
                If WorkUnit > 0 Then
                    .Item("WorkUnit") = WorkUnit
                Else
                    .Item("WorkUnit") = PRCode.Rows(0).Item("WorkUnit")
                End If
                .Item("CompCode") = LoginInfo.CompCode
                .Item("InstCount") = 1

                .Item("ServiceType") = ServiceType
                .Item("NodeNo") = Address.Rows(0).Item("NodeNo")
                .Item("SalesCode") = Address.Rows(0).Item("SalesCode")
                .Item("SalesName") = Address.Rows(0).Item("SalesName")

                .Item("WorkServCode") = .Item("ServCode")
                .Item("ModifyFlag") = 1
            End With
            Wip.Rows.Add(WipRow)
            GroupCode.Dispose()
            Address.Dispose()
            SOUtil.Dispose()
            PRCode.Dispose()
            dtCustId.Dispose()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '產生指定設備資料
    Private Function GetChangeFaci(ByVal CustId As Integer, ByVal SNo As String,
                                   ByVal Wip As DataTable, ByVal Charge As DataTable,
                                   ByVal Facility As DataTable, ByRef ChangeFacility As DataTable) As Boolean
        Try
            If String.IsNullOrEmpty(SNo) Then
                SNo = Wip.Rows(0).Item("SNo")
            End If
            Dim WipRow As DataRow = Wip.Select(String.Format("SNo = '{0}'", SNo)).FirstOrDefault
            Using PRCode As DataTable = New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO).GetCode(BLL.Utility.CodeType.PRCode, WipRow.Item("PRCode").ToString(), False)
                Dim InstRefNo As Integer = 0
                Dim ReInstAcrossFlag As Integer = False
                If Not PRCode.Rows(0).IsNull("RefNo") Then
                    InstRefNo = Integer.Parse(PRCode.Rows(0).Item("RefNo"))
                End If
                If Not PRCode.Rows(0).IsNull("ReInstAcrossFlag") Then
                    ReInstAcrossFlag = Integer.Parse(PRCode.Rows(0).Item("ReInstAcrossFlag")) = 1
                End If
                Dim FaciRows() As DataRow = Facility.Select(String.Format("SNO = '{0}'", SNo))
                Dim WipUtil As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
                '4.1.如有新增設備需做新增設備的指定
                If Not WipUtil.GetChangeFacility(Utility.FaciChangeType.Append, WipRow, FaciRows, FaciRows, Nothing, ChangeFacility) Then
                    Throw New Exception("GetChangeFaci")
                End If
                WipUtil.Dispose()
            End Using
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '互動進入，如過有指定設備的話需要將設備填入ChangeFacility
    Public Function GetDefChangeFaciData(ByVal dsWipData As DataSet, ByVal WipType As Int32,
                                         WipRefNo As Int32, ReInstAcrossFlag As Boolean) As DataSet
        Try
            Dim FaciSeqNo As String = String.Empty
            '判斷是否有(互動)資料
            If dsWipData.Tables.Contains("Contact") Then
                '判斷欄位(設備流水號)是否存在
                If dsWipData.Tables("Contact").Columns.Contains("FaciSeqNo") Then
                    '判斷互動資料是否有內容
                    If dsWipData.Tables("Contact").Rows.Count > 0 Then
                        '判斷設備流水號內容是否為NULL
                        If dsWipData.Tables("Contact").Rows(0).IsNull("FaciSeqNo") Then
                            Return dsWipData
                        Else
                            FaciSeqNo = dsWipData.Tables("Contact").Rows(0)("FaciSeqNo")
                        End If
                    Else
                        Return dsWipData
                    End If
                Else
                    Return dsWipData
                End If
            Else
                Return dsWipData
            End If

            Using objUtly As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
                '取得KindCode的值
                Dim KindCode As Int32 = objUtly.GetCanChangeKind(BLL.Utility.InvoiceType.PR, WipRefNo, ReInstAcrossFlag).Rows(0)("KindCode")
                '判斷KindCode，呼叫不一樣的功能，取回工單資料
                Dim dsChangeWip As DataSet = Nothing
                Dim dtChangeFacility As DataTable = Nothing
                Using dtFacility As DataTable = DAO.ExecQry(_DAL.GetFaciSeqNoData, New Object() {dsWipData.Tables("Wip").Rows(0)("CustId"), FaciSeqNo})
                    Using objChangeFaci As New CableSoft.SO.BLL.Facility.ChangeFaci.ChangeFaci(LoginInfo, DAO)
                        Dim ChooseServiceID As String = Nothing
                        Select Case KindCode
                            Case 304
                                dsChangeWip = objChangeFaci.GetPRFaci(dsWipData.Tables("Wip").Rows(0)("SNO"), FaciSeqNo)
                                '取回工單資料找出 ChangeFacility
                                If dsChangeWip.Tables.Contains("ChangeFacility") Then dtChangeFacility = dsChangeWip.Tables("ChangeFacility").Copy
                                '取回工單資料找出 PrFacility,並且取代 dsWipData.Tables("PrFacility")
                                If dsWipData.Tables.Contains("PrFacility") Then dsWipData.Tables.Remove("PrFacility")
                                dsWipData.Tables.Add(dsChangeWip.Tables("PrFacility").Copy)
                            Case 308
                                dsChangeWip = objChangeFaci.GetMovePRFaci(dsWipData.Tables("Wip").Rows(0)("SNO"), FaciSeqNo)
                                '取回工單資料找出 ChangeFacility
                                If dsChangeWip.Tables.Contains("ChangeFacility") Then dtChangeFacility = dsChangeWip.Tables("ChangeFacility").Copy
                                '取回工單資料找出 PrFacility,並且取代 dsWipData.Tables("PrFacility")
                                If dsWipData.Tables.Contains("PrFacility") Then dsWipData.Tables.Remove("PrFacility")
                                dsWipData.Tables.Add(dsChangeWip.Tables("PrFacility").Copy)
                            Case Else
                                dtChangeFacility = DAO.ExecQry("Select * From SO004D Where 0=1")
                                objUtly.GetChangeFacility(KindCode, dsWipData.Tables("Wip").Rows(0), dtFacility.Rows(0), dtFacility.Rows(0), Nothing, dtChangeFacility)
                        End Select
                        Using ChargeData As DataTable = DAO.ExecQry(_DAL.Get003CData, New Object() {dsWipData.Tables("Wip").Rows(0)("CustId"), FaciSeqNo})
                            For Each row As DataRow In ChargeData.Rows
                                ChooseServiceID = String.Format("{0},{1}", ChooseServiceID, row("ServiceId"))
                            Next
                            If ChooseServiceID IsNot Nothing Then ChooseServiceID = ChooseServiceID.Substring(1)
                        End Using
                        '取代目前WipData內的ChangeFacility
                        If dsWipData.Tables.Contains("ChangeFacility") Then
                            dsWipData.Tables.Remove("ChangeFacility")
                        End If
                        If ChooseServiceID IsNot Nothing Then dtChangeFacility.Rows(0)("ChooseServiceID") = ChooseServiceID
                        dtChangeFacility.TableName = "ChangeFacility"
                        dsWipData.Tables.Add(dtChangeFacility.Copy)
                    End Using
                End Using
            End Using
            Return dsWipData
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetCanMoveServiceType(ByVal CustId As Integer, ByVal ServiceType As String) As DataSet
        Try
            Using dsWipOther As New DataSet()
                Using dtOtherServiceType As DataTable = DAO.ExecQry(_DAL.GetCanMoveServiceType, New Object() {CustId, ServiceType})
                    dtOtherServiceType.TableName = "Wip"
                    dsWipOther.Tables.Add(dtOtherServiceType.Copy)
                    For Each drRow As DataRow In dtOtherServiceType.Rows
                        Using dtCode As DataTable = DAO.ExecQry(_DAL.GetPRCode(3, String.Empty, String.Empty, False, False), New Object() {drRow("ServiceType")})
                            dtCode.TableName = drRow("ServiceType") & "PrCode"
                            dsWipOther.Tables.Add(dtCode.Copy)
                        End Using
                        Using dtReason As DataTable = DAO.ExecQry(_DAL.GetPRReasonCode, New Object() {drRow("ServiceType")})
                            dtReason.TableName = drRow("ServiceType") & "ReasonCode"
                            dsWipOther.Tables.Add(dtReason.Copy)
                        End Using
                    Next
                End Using
                Return dsWipOther
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetChangeFacilityPinCode(ByVal CustId As Integer, ByVal InSeqNo As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetChangeFacilityPinCode(String.Format("'{0}'", InSeqNo.Replace("'", "").Replace(",", "','"))), New Object() {CustId})
            Return dt
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function GetChangeFacilitySEQNO(ByVal dtChangeFaci As DataTable) As String
        Try
            Dim SEQNO As String = String.Empty
            For Each drChange As DataRow In dtChangeFaci.Rows
                If drChange("SEQNO") IsNot Nothing Then SEQNO = String.Format("{0},{1}", SEQNO, drChange("SEQNO"))
            Next
            If Not String.IsNullOrEmpty(SEQNO) Then
                If SEQNO.Substring(0, 1) = "," Then SEQNO = SEQNO.Substring(1)
            End If
            Return SEQNO
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "刪除暫存點數"
    Public Function DelResvPoint() As Boolean
        Try
            Dim WipUtil As New CableSoft.SO.BLL.Wip.Utility.SaveData(LoginInfo, DAO)
            Return WipUtil.DelResvPoint()
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
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
