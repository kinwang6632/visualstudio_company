Option Compare Binary
Option Infer On
Option Explicit On

Imports CableSoft.BLL.Utility
Imports WipPRLanguage = CableSoft.BLL.Language.SO61.WipPRLanguage

Public Class PR
    Inherits BLLBasic
    Implements IDisposable

    'Private _DAL As New PRDAL(Me.LoginInfo.Provider)
    Private _DAL As New PRDALMultiDB(Me.LoginInfo.Provider)

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
    ''' 可新增(HTML5)
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanAppend(ByVal CustId As Integer) As RIAResult
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim aRet As RIAResult = Nothing
        Try
            aRet = obj.ChkPriv(Me.LoginInfo.EntryId, "SO11131")
        Finally
            obj.Dispose()
        End Try
        Return aRet
    End Function

    ''' <summary>
    ''' 可新增(Silverlight)
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

            '2019.04.03 Corey 以前沒有服務別，所以所有的服務別都要檢核。
            '                 現在畫面已經有服務別的功能，所以只需要檢核C服務就好。因為可能其他服務別(P)是拆機，依樣會秀訊息
            If ServiceType = "X" OrElse String.IsNullOrEmpty(ServiceType) Then ServiceType = "C"

            If Not aRet.ResultBoolean Then
                Return aRet
            Else
                If ServiceType = "X" OrElse String.IsNullOrEmpty(ServiceType) Then
                    '沒有傳入服務別或是傳入X，需要判斷所有的狀態種類，所有服務別都不能派拆才秀訊息，其中一個服務別可以派拆就要可以進入工單
                    Dim CanSerViceType As String = String.Empty
                    Dim CanUseRefNo As String = String.Empty
                    Dim ErrCode As Integer = 0
                    Dim ErrMessage As String = String.Empty
                    Dim CanRefNo As String = String.Empty
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
                        Dim CanRefNo As String = String.Empty
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

        '2012.07.02 和Jacky,Karen討論後，未來用 by機 所以不需要特別判斷CustStatusCode,只需要判斷=4就可以
        If CustStatus = 4 Then
            ErrorCode = -4
            ErrorMessage = WipPRLanguage.ChkCustIsCancel
        End If

        'Select Case CustStatus
        '    Case 1
        '    Case 4
        '        ErrorCode = -4
        '        'ErrorMessage = "註銷戶無法產生派工單！"
        '        ErrorMessage = WipPRLanguage.ChkCustIsCancel
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

        Return True
    End Function

    
    ''' <summary>
    ''' 可修改
    ''' </summary>
    ''' <param name="PRSNO">拆機工單號碼</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanEdit(ByVal PRSNO As String) As RIAResult
        Dim numMID As String = "SO11132"
        Dim result As RIAResult = New RIAResult() With {.ResultBoolean = True}
        Using Wip As DataTable = DAO.ExecQry(_DAL.GetWipData("ClsTime"), New Object() {PRSNO})
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                Dim rData As DataTable = SOUtil.GetPriv(LoginInfo.EntryId, numMID)
                If rData.Rows.Count = 0 OrElse CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(rData.Rows(0).Item("GroupX")) = 0 Then
                    '判斷是否有修改的權限
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -99, .ErrorMessage = WipPRLanguage.NoPriv}
                ElseIf Wip.Rows.Count = 0 Then
                    '判斷是否有工單資料
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -99, .ErrorMessage = WipPRLanguage.WipPRDataNothing}
                ElseIf Not Wip.Columns.Contains("ClsTime") Then
                    '判斷是否有日結欄位
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = WipPRLanguage.colClsTimeNothing}
                ElseIf Wip.Rows(0).IsNull("ClsTime") = False Then
                    '已日結不可修改資料(Wip.ClsTime is not Null)。
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -99, .ErrorMessage = WipPRLanguage.colClsTimeIsNotNull}
                Else
                    Dim ChkManager As DataRow = rData.AsEnumerable.Where(Function(list) list.Item("Mid") = numMID & "1").FirstOrDefault()
                    If ChkManager Is Nothing OrElse CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(ChkManager.Item("GroupX")) = 0 Then
                        result.ErrorCode = 1
                    End If
                End If
            End Using
        End Using
        Return result
    End Function
    ''' <summary>
    ''' 可作廢
    ''' </summary>
    ''' <param name="PRSNO">拆機工單號碼</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function CanDelete(ByVal PRSNO As String) As RIAResult
        Dim numMID As String = "SO11133"
        Dim result As RIAResult = New RIAResult() With {.ResultBoolean = True}
        Using Wip As DataTable = DAO.ExecQry(_DAL.GetWipData("ClsTime"), New Object() {PRSNO})
            Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                '(1)	Priv=CableSoft.SO.BLL.Utility.Utility.ChkPriv(UserId,’SO11112’)
                Dim rData As DataTable = SOUtil.GetPriv(LoginInfo.EntryId, numMID)
                If rData.Rows.Count = 0 OrElse CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(rData.Rows(0).Item("GroupX")) = 0 Then
                    '判斷是否有修改的權限
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -99, .ErrorMessage = WipPRLanguage.NoPriv}
                ElseIf Wip.Rows.Count = 0 Then
                    '判斷是否有工單資料
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -99, .ErrorMessage = WipPRLanguage.WipPRDataNothing}
                ElseIf Not Wip.Columns.Contains("ClsTime") Then
                    '判斷是否有日結欄位
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -1, .ErrorMessage = WipPRLanguage.colClsTimeNothing}
                ElseIf Wip.Rows(0).IsNull("ClsTime") = False Then
                    '已日結不可修改資料(Wip.ClsTime is not Null)。
                    result = New RIAResult With {.ResultBoolean = False, .ErrorCode = -99, .ErrorMessage = WipPRLanguage.colClsTimeIsCancel}
                Else
                    '2018.06.22 by Corey 增加修改資料需主管驗證
                    Dim ChkManager As DataRow = rData.AsEnumerable.Where(Function(list) list.Item("Mid") = numMID & "1").FirstOrDefault()
                    If ChkManager Is Nothing OrElse CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(ChkManager.Item("GroupX")) = 0 Then
                        result.ErrorCode = 1
                    End If
                End If
            End Using
        End Using
        Return result
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
            Using dtWip As DataTable = DAO.ExecQry(_DAL.GetWipData("ClsTime"), PRSNO, False)
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
    ''' (新增狀態)需要取得可以使用的服務別資料
    ''' </summary>
    ''' <param name="EditMode">工單狀態</param>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">特定服務別 C,D,I,P 或是 X,空的表示不考慮服務別全都要</param>
    ''' <returns>RIAResult</returns>
    ''' <remarks></remarks>
    Public Function GetAppendCanUseServiceType(ByVal EditMode As EditMode, ByVal CustId As Integer, ByVal ServiceType As String) As RIAResult
        Try
            Dim retOK As Boolean = False, retResultXML As String = String.Empty
            Dim ErrCode As Integer = 0
            Dim ErrMessage As String = String.Empty
            Dim CanSerViceType As String = String.Empty
            Dim CanUseRefNo As String = String.Empty

            '2019.04.03 Corey 以前沒有服務別，所以所有的服務別都要檢核。
            '                 現在畫面已經有服務別的功能，所以只需要檢核C服務就好。因為可能其他服務別(P)是拆機，依樣會秀訊息
            If ServiceType = "X" OrElse String.IsNullOrEmpty(ServiceType) Then ServiceType = "C"

            If ServiceType = "X" OrElse String.IsNullOrEmpty(ServiceType) Then
                '沒有傳入服務別或是傳入X，需要判斷所有的狀態種類，所有服務別都不能派拆才秀訊息，其中一個服務別可以派拆就要可以進入工單
                Using dt As DataTable = DAO.ExecQry(_DAL.GetCustomer(ServiceType), CustId, False)
                    Dim strXML As String = String.Empty
                    For Each dr As DataRow In dt.Rows
                        If ServiceTypeCanAppend(dr, ErrCode, ErrMessage, CanUseRefNo) Then
                            If ErrCode >= 0 Then
                                strXML = String.Format("{0}-CanServiceType={1};CanRefNo={2};ErrMessage={3}", strXML, CanSerViceType, CanUseRefNo, ErrMessage)
                            End If
                        Else
                            ErrCode = -99
                            ErrMessage = WipPRLanguage.ChkCustStatusError
                            Exit For
                        End If
                    Next
                    If Not String.IsNullOrEmpty(strXML) Then
                        '有正常服務別的客戶資料需要回傳
                        If strXML.Substring(0) = "-" Then strXML = strXML.Substring(1)
                        retOK = True
                        retResultXML = strXML
                    Else
                        '沒有服務別資料表是所有服務別都是拆機的狀態須回傳最後的ErrorCode,ErrorMessage
                    End If
                End Using
            Else
                '有傳入服務別，可針對服務別來判斷
                Using dt As DataTable = DAO.ExecQry(_DAL.GetCustomer(ServiceType), New Object() {CustId})
                    CanSerViceType = ServiceType
                    If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                        For Each dr As DataRow In dt.Rows
                            If ServiceTypeCanAppend(dr, ErrCode, ErrMessage, CanUseRefNo) Then
                                If ErrCode < 0 Then
                                    Exit For
                                Else
                                    retOK = True
                                End If
                            Else
                                ErrCode = -99
                                ErrMessage = WipPRLanguage.ChkCustStatusError
                                Exit For
                            End If
                        Next
                    End If

                    If Not String.IsNullOrEmpty(CanUseRefNo) Then
                        '有正常服務別的客戶資料需要回傳
                        retOK = True
                        retResultXML = String.Format("CanServiceType={0};CanRefNo={1};ErrMessage={2}", CanSerViceType, CanUseRefNo, ErrMessage)
                    End If
                End Using
            End If
            Return New RIAResult() With {.ResultBoolean = retOK, .ErrorCode = ErrCode, .ErrorMessage = ErrMessage, .ResultXML = retResultXML}
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ServiceTypeCanAppend(ByVal drData As DataRow, ByRef ErrorCode As Integer, ByRef ErrorMessage As String,
                                          ByRef CanUseRefNo As String) As Boolean
        ErrorCode = 0
        ErrorMessage = String.Empty
        Dim intCustStatusCode As Integer = Int16.Parse("0" & drData("CustStatusCode").ToString)
        '#7536 2017.08.22 by Corey 原本2012.07.02討論不做，需求又要做了。所以要將下面功能打開。
        Select Case intCustStatusCode
            Case 1
            Case 4
                ErrorCode = -4
                'ErrorMessage = "註銷戶無法產生派工單！"
                ErrorMessage = WipPRLanguage.ChkCustIsCancel
            Case 2, 3, 6, 5
                Using dt2 As DataTable = DAO.ExecQry(_DAL.GetSO042, New Object() {drData("ServiceType")})
                    For Each dr2 As DataRow In dt2.Rows
                        If intCustStatusCode = 2 And Int32.Parse(dr2.Item("StopPR").ToString) = 1 Then
                            ErrorCode = 10
                            ErrorMessage = WipPRLanguage.CustomerUcanAdd(drData("CustStatusName").ToString)
                            CanUseRefNo = "2,6,8,10"
                        Else
                            If ",2,3,6,".Contains(String.Format(",{0},", intCustStatusCode.ToString)) Or (Int32.Parse(dr2.Item("AbnormalFaci").ToString) = 1 AndAlso intCustStatusCode = 5) Then
                                ErrorMessage = WipPRLanguage.CustomerIsPR(drData("CustStatusName").ToString)
                                If intCustStatusCode = 5 Then
                                    ErrorCode = 10
                                    CanUseRefNo = "6,8,10"
                                Else
                                    ErrorCode = 10
                                    CanUseRefNo = "6,7,9,10"
                                End If
                            End If
                        End If
                    Next
                End Using

                'Case Else
                '    ErrorCode = 99
                '    ErrorMessage = "可派的派工參考號: 6,7,9,10,15"
                '    ErrorMessage = "6,7,9,10,15"
        End Select

        If Int16.Parse("0" & drData("WipCode3").ToString) <> 0 Then
            'ErrorMessage = "該客戶已是[" & drData("WipName3").ToString & "]派工中狀態, 是否確認新增?"
            ErrorCode = 10
            ErrorMessage = WipPRLanguage.CustomerIsWap(drData("WipName3").ToString)
        End If

        Return True
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

    Public Function GetServiceType(ByVal CanUseServiceType As String, ByVal lngCustid As Integer, ByVal strFaciSEQNO As String) As DataTable
        If Not String.IsNullOrEmpty(strFaciSEQNO) Then
            Using dtFaci As DataTable = DAO.ExecQry(_DAL.GetFaciSeqNoData, New Object() {lngCustid, strFaciSEQNO})
                If dtFaci.Rows.Count > 0 Then
                    Dim strServiceType As String = dtFaci.Rows(0)("ServiceType").ToString
                    If Not String.IsNullOrEmpty(strServiceType) Then
                        If Not String.Format(",{0},", CanUseServiceType).Contains(String.Format(",{0},", strServiceType)) Then
                            CanUseServiceType = CanUseServiceType & "," & strServiceType
                            If CanUseServiceType.Substring(0, 1) = "," Then CanUseServiceType = CanUseServiceType.Substring(1)
                        End If
                    End If
                End If
            End Using
        End If
        Dim dt As DataTable = DAO.ExecQry(_DAL.GetServiceType(CanUseServiceType))
        Return dt
    End Function

   
    Private Function GetWipRefNoByContact(ContactRefNo As Integer, ByRef ReInstAcrossFlag As Boolean) As String
        Dim WorkRefNoStr As String = Nothing
        Select Case ContactRefNo
            Case 1, 2, 7, 8, 10, 13, 14, 16, 24, 25, 29, 35, 38, 40, 41, 43, 47, 55
                Select Case ContactRefNo
                    Case 1      '新裝機
                        WorkRefNoStr = "1,5,7,17"
                    Case 2      '設備加裝
                        WorkRefNoStr = "2,3,4"
                    Case 7      '復機
                        WorkRefNoStr = "5,7,17"
                    Case 8      '客戶改裝
                        WorkRefNoStr = 6
                    Case 10     '裝e-Box
                        WorkRefNoStr = 9
                    Case 16     'CM 升降級
                        WorkRefNoStr = "10,11"
                    Case 24     'STB 升降級
                        WorkRefNoStr = 12
                    Case 25     'CM促案變更
                        WorkRefNoStr = 14
                    Case 29     '續約
                        WorkRefNoStr = 15
                    Case 38, 40 '加約,恢復頻道
                        WorkRefNoStr = 21
                    Case 41     '41=申請固定IP 
                        WorkRefNoStr = 13
                    Case 43     '43=DVR升降容量
                        WorkRefNoStr = "22,23,24"
                    Case 47     '47=頻道更換(裝機單) 
                        WorkRefNoStr = "25"
                    Case 55     '55=開訊
                        WorkRefNoStr = "8"
                    Case Else
                End Select
            Case 3, 36
            Case 4, 5, 6, 11, 27, 28, 30, 32, 33, 34, 37, 39, 42
                Select Case ContactRefNo
                    Case 4          '移機
                        WorkRefNoStr = 3
                    Case 5          '拆機
                        WorkRefNoStr = "2,5,6,8"
                    Case 6          '停機
                        WorkRefNoStr = "1,11"
                    Case 11         '拆設備
                        WorkRefNoStr = "6,8"
                    Case 27, 28     '同/跨區移機
                        WorkRefNoStr = "2,6,8"
                        ReInstAcrossFlag = True
                    Case 30         '關機
                        WorkRefNoStr = "7"
                    Case 32, 33, 56 '32=頻道結清,33=頻道更換(拆裝),56=PVR拆機
                        WorkRefNoStr = "10"
                    Case 39         '暫停頻道
                        WorkRefNoStr = "15"
                    Case 42         '42=取消固定IP
                        WorkRefNoStr = "14"
                    Case Else
                End Select
            Case Else
        End Select
        Return WorkRefNoStr
    End Function

    ''' <summary>
    ''' 畫面起始載入(HTML5)
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="SNO">工單號碼</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="ServCode">指定服務區</param>
    ''' <param name="WipRefNo">工單參考號</param>
    ''' <param name="WipCodeValueStr">特定派工類別號碼</param>
    ''' <param name="dsInitData">畫面載入條件</param>
    ''' <param name="strRetMsgCheck">回傳訊息，給前端判斷用的</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetInitData2(ByVal CustId As Integer, ByVal SNO As String, ByVal ServiceType As String,
                                ByVal ServCode As String, ByVal WipRefNo As Integer, ByVal WipCodeValueStr As String,
                                ByRef dsInitData As DataSet, ByRef strRetMsgCheck As String, ByVal EditMode As EditMode) As DataSet
        Try
            Dim dsDefaultLoad As New DataSet
            Dim dsDefaultLoad2 As New DataSet
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
                    If Not ",SO001,SO002,".Contains(dtGet.TableName) Then
                        dsDefaultLoad.Tables.Add(dtGet.Copy)
                    End If
                Next


                Dim CanUseRefNo As String = String.Empty, blnReInstFilter As Boolean = False, ReInstAcrossFlag As Boolean = False
                If WipRefNo > 0 Then
                    blnReInstFilter = True
                    CanUseRefNo = GetWipRefNoByContact(WipRefNo, ReInstAcrossFlag)
                End If
                If String.IsNullOrEmpty(CanUseRefNo) Then
                    Dim aRet As RIAResult = GetAppendCanUseServiceType(EditMode, CustId, ServiceType)
                    If aRet.ResultBoolean Then
                        Dim strReturn As String = aRet.ResultXML
                        If Not String.IsNullOrEmpty(strReturn) Then
                            Dim aryData() As String = strReturn.Split("-")
                            For Each strData As String In aryData
                                Dim aryData2() As String = strData.Split(";")
                                If strData.Contains(String.Format("CanServiceType={0};", ServiceType)) Then
                                    For Each strServ As String In aryData2
                                        Dim aryCanUse() As String = strServ.Split("=")
                                        If aryCanUse(0).ToUpper = "CanRefNo".ToUpper Then
                                            CanUseRefNo = aryCanUse(1)
                                        End If
                                        If aryCanUse(0).ToUpper = "ErrMessage".ToUpper Then
                                            strRetMsgCheck = aryCanUse(1)
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    End If
                End If
                '#8475 By Kin 2019/08/19
                Dim strCD002Code As String = "-X"
                If dsDefaultLoad.Tables("Wip").Rows.Count > 0 AndAlso Not DBNull.Value.Equals(dsDefaultLoad.Tables("Wip").Rows(0).Item("ServCode")) Then
                    strCD002Code = dsDefaultLoad.Tables("Wip").Rows(0).Item("ServCode")
                End If
                Using dtCD002 As DataTable = GetCD002(strCD002Code)
                    dsDefaultLoad2.Tables.Add(dtCD002.Copy)
                End Using
                '客戶主檔資料
                Using dtSO001 As DataTable = DAO.ExecQry(_DAL.GetSO001, CustId, False)
                    If dtSO001 IsNot Nothing Then
                        dtSO001.TableName = "SO001"
                        dsDefaultLoad2.Tables.Add(dtSO001.Copy)
                    End If
                    If dtSO001.Rows.Count > 0 Then
                        '#8173 2019.03.08 by Corey 增加利用SO001.ID 找到SO137申請人資料
                        Dim IDData As String = dtSO001.Rows(0)("ID").ToString
                        If Not String.IsNullOrEmpty(IDData) Then
                            Using dtSO137 As DataTable = DAO.ExecQry(_DAL.GetDeclarantData(), New Object() {IDData})
                                dtSO137.TableName = "SO137"
                                dsDefaultLoad2.Tables.Add(dtSO137.Copy)
                            End Using
                        End If
                    End If
                End Using
                Using dtSO002 As DataTable = DAO.ExecQry(_DAL.GetSO002(ServiceType), New Object() {CustId})
                    If dtSO002 IsNot Nothing Then
                        dtSO002.TableName = "SO002"
                        dsDefaultLoad2.Tables.Add(dtSO002.Copy)
                    End If
                End Using

                'ServiceType
                Using dtServiceType As DataTable = GetServiceType(String.Empty)
                    dtServiceType.TableName = "ServiceType"
                    dsDefaultLoad2.Tables.Add(dtServiceType.Copy)
                End Using

                'PRCode
                Using dtPrCode As DataTable = GetPRCode(ServiceType, CanUseRefNo, WipCodeValueStr, blnReInstFilter, ReInstAcrossFlag)
                    dtPrCode.TableName = "PRCode"
                    dsDefaultLoad2.Tables.Add(dtPrCode.Copy)
                End Using
                'ALLPRCode
                Using dtAllPrCode As DataTable = DAO.ExecQry(_DAL.GetALLPRCode)
                    dtAllPrCode.TableName = "ALLPRCode"
                    dsDefaultLoad2.Tables.Add(dtAllPrCode.Copy)
                End Using

                'PRReasonCode(3)
                Using dtPRReasonCode As DataTable = GetPRReasonCode(ServiceType)
                    dtPRReasonCode.TableName = "PRReasonCode"
                    dsDefaultLoad2.Tables.Add(dtPRReasonCode.Copy)
                End Using
                If dsPRData.Tables("Wip").Rows.Count > 0 Then
                    Using dtPRReasonDescCode As DataTable = GetPRReasonDescCode(ServiceType, dsPRData.Tables("Wip").Rows(0).Item("ReasonCode"))
                        dsDefaultLoad2.Tables.Add(dtPRReasonDescCode.Copy)
                    End Using
                Else
                    Using dtPRReasonDescCode As DataTable = GetPRReasonDescCode(ServiceType, -1)
                        dsDefaultLoad2.Tables.Add(dtPRReasonDescCode.Copy)
                    End Using
                End If

                'ReturnCode(4)
                Using dtReturnCode As DataTable = GetReturnCode(ServiceType)
                    dtReturnCode.TableName = "ReturnCode"
                    dsDefaultLoad2.Tables.Add(dtReturnCode.Copy)
                End Using
                'ReturnDescCode(5)
                Using dtReturnDescCode As DataTable = GetReturnDescCode(ServiceType)
                    dtReturnDescCode.TableName = "ReturnDescCode"
                    dsDefaultLoad2.Tables.Add(dtReturnDescCode.Copy)
                End Using
                'GroupCode(6, ServCode)
                Using dtGroupCod As DataTable = GetGroupCode(ServCode)
                    dtGroupCod.TableName = "GroupCode"
                    dsDefaultLoad2.Tables.Add(dtGroupCod.Copy)
                End Using
                'SignEn(7)
                Using dtSignEn As DataTable = GetSignEn()
                    dtSignEn.TableName = "SignEn"
                    dsDefaultLoad2.Tables.Add(dtSignEn.Copy)
                End Using
                'SatiCode(8)
                Using dtSatiCode As DataTable = GetSatiCode(ServiceType)
                    dtSatiCode.TableName = "SatiCode"
                    dsDefaultLoad2.Tables.Add(dtSatiCode.Copy)
                End Using
                'WorkerEn(9, 0)
                Using dtWorkerEn As DataTable = GetWorkerEn(0)
                    dtWorkerEn.TableName = "WorkerEn0"
                    dsDefaultLoad2.Tables.Add(dtWorkerEn.Copy)
                End Using
                'WorkerEn(10, 1)
                Using dtWorkerEn As DataTable = GetWorkerEn(1)
                    dtWorkerEn.TableName = "WorkerEn1"
                    dsDefaultLoad2.Tables.Add(dtWorkerEn.Copy)
                End Using
                '#8475 CD002 By Kin 


                'CustRtnCode(11)
                Using dtCustRtnCode As DataTable = GetCustRtnCode(ServiceType)
                    dtCustRtnCode.TableName = "CustRtnCode"
                    dsDefaultLoad2.Tables.Add(dtCustRtnCode.Copy)
                End Using
                'SO042(12)
                Using dtSO042 As DataTable = GetSO042(ServiceType)
                    dtSO042.TableName = "SO042"
                    dsDefaultLoad2.Tables.Add(dtSO042.Copy)
                End Using
                'SO041(13)
                Using dtSO041 As DataTable = DAO.ExecQry(_DAL.GetSO041)
                    dtSO041.TableName = "SO041"
                    dsDefaultLoad2.Tables.Add(dtSO041.Copy)
                End Using
                'GetPriv(14)
                Using dtGetPriv As DataTable = GetPriv()
                    dtGetPriv.TableName = "GetPriv"
                    dsDefaultLoad2.Tables.Add(dtGetPriv.Copy)
                End Using
                'GetUserPriv(15)
                Using dtGetUserPriv As DataTable = GetUserPriv()
                    dtGetUserPriv.TableName = "UserPriv"
                    dsDefaultLoad2.Tables.Add(dtGetUserPriv.Copy)
                End Using
                'GetAddressData(16)
                Using dtGetAddressData As DataTable = GetAddressData(CustId)
                    dtGetAddressData.TableName = "GetAddressData"
                    dsDefaultLoad2.Tables.Add(dtGetAddressData.Copy)
                End Using
                '2018.06.20 by Corey 增加取得 SO029B 鎖定權限
                Using bll As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                    Using dt As DataTable = bll.GetFieldPrivMappingData("SO1113A", EditMode)
                        dt.TableName = "FieldPriv"
                        dsDefaultLoad2.Tables.Add(dt.Copy)
                    End Using
                End Using

                dsInitData = dsDefaultLoad2.Copy
            End Using
            Return dsDefaultLoad
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 畫面起始載入(Silverlight)
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
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRCode(CanUseRefNo, CanNotUseRefNo, WipCodeValueStr, blnReInstFilter, ReInstAcrossFlag, False), ServiceType, False)
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
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetPRReasonCode, New Object() {ServiceType})
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
    Public Function GetCD002(ByVal codeNo As String) As DataTable
        Try
            Dim dt As DataTable = DAO.ExecQry(_DAL.GetCD002, New Object() {codeNo, LoginInfo.CompCode})
            dt.TableName = "CD002"
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
                Dim dt As DataTable = SOUtil.GetPriv(LoginInfo.EntryId, New String() {"SO1113", "SO1113"})
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
    ''' <param name="OtherData">Change:收費資料 Contact:互動資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNormalCalculateData(ByVal CustId As Integer,
                                           ByVal ServiceType As String,
                                           ByVal WorkCodeValue As Integer,
                                           ByVal ResvTime As DateTime,
                                           ByVal SNo As String,
                                           ByVal Full As Boolean,
                                           ByVal OtherData As DataSet) As DataSet
        Return GetNormalCalculateData(CustId, ServiceType, WorkCodeValue, ResvTime, SNo, Full, OtherData, Nothing)
    End Function
    ''' <summary>
    ''' 取得一般工單資料
    ''' </summary>
    ''' <param name="CustId">客戶編號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="WorkCodeValue">派工類別</param>
    ''' <param name="ResvTime">預約時間</param>
    ''' <param name="Full">完整</param>
    ''' <param name="OtherData">Change:收費資料 Contact:互動資料</param>
    ''' <param name="oldWipData">原工單資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNormalCalculateData(ByVal CustId As Integer,
                                           ByVal ServiceType As String,
                                           ByVal WorkCodeValue As Integer,
                                           ByVal ResvTime As DateTime,
                                           ByVal SNo As String,
                                           ByVal Full As Boolean,
                                           ByVal OtherData As DataSet,
                                           ByVal oldWipData As DataSet) As DataSet
        Try
            Dim WipData As DataSet = Nothing
            Dim dtChangeData As DataTable = Nothing
            Dim dtContact As DataTable = Nothing

            Using WipUtil As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
                If OtherData IsNot Nothing Then
                    If OtherData.Tables.Contains("Contact") Then
                        dtContact = OtherData.Tables("Contact").Copy
                    End If
                End If
                '取得收費/設備資料
                'obj.GetWipCalculateData(BLL.Utility.InvoiceType.Maintain, CustId, ServiceType, Nothing, ResvTime, InstCode, dtContact, dsWipData)
                WipData = WipUtil.GetWipCalculateData(BLL.Utility.InvoiceType.PR, CustId, ServiceType, SNo, ResvTime, WorkCodeValue, dtContact, oldWipData)
                Dim Wip As DataTable = WipData.Tables("Wip")
                '取得工單資料
                Dim WorkCode As DataTable = Nothing
                If Wip.Rows.Count <= 0 Then
                    If Not GetNormalWip(CustId, ServiceType, SNo, ResvTime, WorkCodeValue, 0, Wip, WorkCode) Then
                        Throw New Exception("GetNormalWip")
                    End If
                    If WipData.Tables.Contains(Wip.TableName) Then
                        WipData.Tables.Remove(Wip.TableName)
                    End If
                    WipData.Tables.Add(Wip.Copy)
                End If
                '#8432 restore the facisno,faciseqno to the new table of charge by kin 2019/07/04
                If oldWipData IsNot Nothing AndAlso oldWipData.Tables.Contains("Charge") Then
                    For Each rwOldCharge As DataRow In oldWipData.Tables("Charge").Rows
                        For Each rwCharge As DataRow In WipData.Tables("Charge").Rows
                            If (rwCharge("BillNo") = rwOldCharge("BillNo")) AndAlso (rwCharge("Item") = rwOldCharge("Item")) Then
                                If Not DBNull.Value.Equals(rwOldCharge("FaciSeqNo")) Then
                                    rwCharge("FaciSeqNo") = rwOldCharge("FaciSeqNo")
                                End If
                                If Not DBNull.Value.Equals(rwOldCharge("FaciSNo")) Then
                                    rwCharge("FaciSNo") = rwOldCharge("FaciSNo")
                                End If
                                rwCharge.AcceptChanges()
                            End If
                        Next
                    Next
                End If
                
                If Full Then
                    '取得指定設備資料(結清功能呼叫拆機工單才會傳ChangeData)
                    Using Charge As DataTable = WipData.Tables("Charge")
                        Using Facility As DataTable = WipData.Tables("Facility")
                            Dim ChangeFacility As DataTable = WipData.Tables("ChangeFacility")
                            If Not GetChangeFaci(CustId, SNo, Wip, Charge, Facility, ChangeFacility) Then
                                Throw New Exception("GetChangeFaci")
                            End If
                            If OtherData IsNot Nothing Then
                                If OtherData.Tables.Contains("ChangeData") Then
                                    dtChangeData = OtherData.Tables("ChangeData").Copy
                                End If
                            End If
                            If Not WipUtil.GetDefaultChangeFaci(BLL.Utility.InvoiceType.PR, Wip, dtChangeData, ChangeFacility) Then
                                Throw New Exception("GetDefaultChangeFaci")
                            End If
                        End Using
                    End Using

                    '如果"互動"過來，有指定設備的話需要指定該設備
                    If dtContact IsNot Nothing Then
                        '#8062 2019.01.24 by Corey 因為文件說明沒有效果，所以改用JACKY提供的功能使用。
                        If dtContact.Rows.Count > 0 Then
                            Dim strFaciSeqno As String = dtContact.Rows(0).Item("FaciSeqNo").ToString
                            Using bll As New CableSoft.SO.BLL.Facility.ChangeFaci.ChangeFaci(LoginInfo, DAO)
                                Dim intWipRefNo As Integer = 0
                                Dim intReInstAcrossFlag As Integer = 0
                                Using bllUlty As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                                    Using PRCode As DataTable = bllUlty.GetCode(SO.BLL.Utility.CodeType.PRCode, WorkCodeValue.ToString, False)
                                        If PRCode.Rows.Count > 0 Then
                                            '取得派工類別參考號 及 移機跨區種類 
                                            intWipRefNo = Int32.Parse("0" & PRCode.Rows(0)("RefNo").ToString)
                                            intReInstAcrossFlag = Int32.Parse("0" & PRCode.Rows(0)("ReInstAcrossFlag").ToString)
                                        End If
                                    End Using
                                End Using
                                If intWipRefNo > 0 Then '有參考號才需要取SO004D KindCode
                                    Dim intKindCode As Integer = 0
                                    Using dtKindCode As DataTable = bll.GetCanChangeKind(2, intWipRefNo, intReInstAcrossFlag)
                                        If dtKindCode.Rows.Count > 0 Then
                                            intKindCode = dtKindCode.Rows(0)("KindCode")
                                        End If
                                    End Using
                                    If intKindCode > 0 Then '判斷 KindCode 決定 Kind種類
                                        '注: intKindCode = 304 or 308 呼叫功能不一樣，其他內容都依樣
                                        Select Case intKindCode
                                            Case 304 And Not String.IsNullOrEmpty(strFaciSeqno) '拆除
                                                Using RetData As DataSet = bll.GetPRFaci(SNo, strFaciSeqno)
                                                    For Each Table As String In New String() {"Facility", "PRFacility", "ChangeFacility"}
                                                        For Each Row As DataRow In RetData.Tables(Table).Rows
                                                            WipData.Tables(Table).Rows.Add(CableSoft.BLL.Utility.Utility.CopyDataRow(Row, WipData.Tables(Table).NewRow()))
                                                        Next
                                                    Next
                                                End Using
                                            Case 308 And Not String.IsNullOrEmpty(strFaciSeqno) '移機
                                                Using RetData As DataSet = bll.GetMovePRFaci(SNo, strFaciSeqno)
                                                    For Each Table As String In New String() {"Facility", "PRFacility", "ChangeFacility"}
                                                        For Each Row As DataRow In RetData.Tables(Table).Rows
                                                            WipData.Tables(Table).Rows.Add(CableSoft.BLL.Utility.Utility.CopyDataRow(Row, WipData.Tables(Table).NewRow()))
                                                        Next
                                                    Next
                                                End Using
                                            Case Else
                                                '互動指定設備
                                                If Not WipUtil.ContactChangeFacility(SO.BLL.Utility.InvoiceType.PR, dtContact, WorkCode, WipData) Then
                                                    Throw New Exception("ContactChangeFacility")
                                                End If
                                        End Select
                                    End If
                                End If
                            End Using
                        End If
                    End If

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
                                  ByRef Wip As DataTable, ByRef WorkCode As DataTable) As Boolean
        Try
            Dim WipRow As DataRow = Wip.NewRow
            With WipRow
                .Item("CustId") = CustId
                .Item("SNo") = SNo

                Using dtCustId As DataTable = DAO.ExecQry(_DAL.GetSO001, New Object() {CustId})
                    If dtCustId.Rows.Count > 0 Then
                        '#6933 2015.02.10 測試不OK 增加CustName,Tel1 兩個欄位
                        .Item("CustName") = dtCustId.Rows(0).Item("CUSTNAME")
                        .Item("Tel1") = dtCustId.Rows(0).Item("Tel1")
                    End If
                End Using
                'SOUtil.GetSystem(BLL.Utility.SystemTableType.System,""

                Using SOUtil As New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
                    Using PRCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, WorkCodeValue.ToString, False)
                        WorkCode = PRCode.Copy
                        Dim rowPRCode As DataRow = PRCode.Rows(0)
                        .Item("PRCode") = rowPRCode("CodeNo")
                        .Item("PRName") = rowPRCode("Description")
                        If WorkUnit > 0 Then
                            .Item("WorkUnit") = WorkUnit
                        Else
                            .Item("WorkUnit") = rowPRCode("WorkUnit")
                        End If
                        If Int32.Parse("0" & rowPRCode("REFNO").ToString) = 3 Then
                            Using dtSystem As DataTable = SOUtil.GetSystem(BLL.Utility.SystemTableType.System, "MovePRResvDay", Nothing)
                                Dim intMovePRResvDay As Integer = 0
                                intMovePRResvDay = Int32.Parse("0" & dtSystem.Rows(0)(0).ToString)
                                .Item("ReInstDate") = ResvTime.AddDays(intMovePRResvDay)
                            End Using
                        End If
                    End Using
                End Using

                Using Address As DataTable = DAO.ExecQry(_DAL.GetAddressData(), New Object() {CustId, LoginInfo.CompCode})
                    .Item("OldAddrNo") = Address.Rows(0).Item("AddrNo")
                    .Item("OldAddress") = Address.Rows(0).Item("Address")
                    .Item("ServCode") = Address.Rows(0).Item("ServCode")
                    .Item("StrtCode") = Address.Rows(0).Item("StrtCode")
                    .Item("NodeNo") = Address.Rows(0).Item("NodeNo")
                    .Item("SalesCode") = Address.Rows(0).Item("SalesCode")
                    .Item("SalesName") = Address.Rows(0).Item("SalesName")
                    Using GroupCode As DataTable = DAO.ExecQry(_DAL.GetGroupCode(), New Object() {Address.Rows(0).Item("ServCode")})
                        If GroupCode.Rows.Count > 0 Then
                            .Item("GroupCode") = GroupCode.Rows(0).Item("CodeNo")
                            .Item("GroupName") = GroupCode.Rows(0).Item("Description")
                        End If
                    End Using
                End Using
                .Item("ResvTime") = ResvTime
                .Item("AcceptTime") = DateTime.Now
                .Item("AcceptEn") = LoginInfo.EntryId
                .Item("AcceptName") = LoginInfo.EntryName
                .Item("CompCode") = LoginInfo.CompCode
                .Item("InstCount") = 1
                .Item("ServiceType") = ServiceType

                .Item("WorkServCode") = .Item("ServCode")
                .Item("ModifyFlag") = 1
            End With
            Wip.Rows.Add(WipRow)
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
    Public Function setChooseServiceIdByCitem(ByVal CustId As String, ByVal CitemCode As Object, ByVal FaciSeqNo As String) As String
        Dim ChooseServiceID As String = Nothing

        Using ChargeData As DataTable = DAO.ExecQry(_DAL.getServiceIdByCitemCode, New Object() {CitemCode, CustId, FaciSeqNo})
            For Each row As DataRow In ChargeData.Rows
                ChooseServiceID = String.Format("{0},{1}", ChooseServiceID, row("ServiceId"))
            Next
            If ChooseServiceID IsNot Nothing Then ChooseServiceID = ChooseServiceID.Substring(1)
        End Using
        Return ChooseServiceID
    End Function
    Public Function setSO004DChooseserviceid(ByVal CustId As String, ByVal FaciSeqNo As String) As String
        Dim ChooseServiceID As String = Nothing
        If String.IsNullOrEmpty(FaciSeqNo) Then FaciSeqNo = "-1"
        Using ChargeData As DataTable = DAO.ExecQry(_DAL.Get003CData, New Object() {CustId, FaciSeqNo})
            For Each row As DataRow In ChargeData.Rows
                ChooseServiceID = String.Format("{0},{1}", ChooseServiceID, row("ServiceId"))
            Next
            If ChooseServiceID IsNot Nothing Then ChooseServiceID = ChooseServiceID.Substring(1)
        End Using
        Return ChooseServiceID
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
                Dim strUseServiceType As String = String.Empty
                Dim strCalcFaciRefNo As String = String.Empty '計費設備參考號
                Using dtCD046 As DataTable = DAO.ExecQry(_DAL.GetServiceType(String.Empty))
                    If dtCD046.Rows.Count > 0 Then
                        For Each drCD046 As DataRow In dtCD046.Rows
                            If drCD046("CodeNo") <> "C" Then
                                '因為該功能是CATV連動產生其他服務別工單，所以CATV不需要考慮進來。
                                strUseServiceType = String.Format("{0},{1}", strUseServiceType, drCD046("CodeNo"))
                            End If
                        Next
                        If Not String.IsNullOrEmpty(strUseServiceType) Then
                            If strUseServiceType.Substring(0) = "," Then
                                strUseServiceType = strUseServiceType.Substring(1)
                            End If
                        End If
                    End If
                End Using
                strCalcFaciRefNo = CableSoft.SO.BLL.Utility.Utility.GetServiceCanChooseRefNo(DAO, strUseServiceType, False, True)

                Using dtOtherServiceType As DataTable = DAO.ExecQry(_DAL.GetCanMoveServiceType(strCalcFaciRefNo), New Object() {CustId, ServiceType})
                    dtOtherServiceType.TableName = "Wip"
                    dsWipOther.Tables.Add(dtOtherServiceType.Copy)
                    For Each drRow As DataRow In dtOtherServiceType.Rows
                        Using dtCode As DataTable = DAO.ExecQry(_DAL.GetPRCode(3, String.Empty, String.Empty, False, False, True), New Object() {drRow("ServiceType")})
                            dtCode.TableName = drRow("ServiceType") & "PrCode"
                            dsWipOther.Tables.Add(dtCode.Copy)
                        End Using
                        Using dtReason As DataTable = DAO.ExecQry(_DAL.GetPRReasonCode(), New Object() {drRow("ServiceType")})
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
