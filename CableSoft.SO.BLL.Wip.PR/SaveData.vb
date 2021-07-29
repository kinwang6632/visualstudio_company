Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess
Imports System.IO
Imports SaveDataLanguage = CableSoft.BLL.Language.SO31.WipPRLanguage

Public Class SaveData
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New SaveDataDAL(Me.LoginInfo.Provider)
    Private SOUtil As CableSoft.SO.BLL.Utility.Utility = Nothing

    Private Const fMaintain_Wip As String = "Wip"
    Private Const fMaintain_Facility As String = "Facility"
    Private Const fMaintain_PRFacility As String = "PRFacility"
    Private Const fMaintain_Charge As String = "Charge"
    Private Const fMaintain_ChangeFacility As String = "ChangeFacility"
    Private Const fMaintain_Parameter As String = "WipPara"


    Public Sub New()
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As DbConnection)
        MyBase.New(LoginInfo, DBConnection)
        
    End Sub
    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub

    Public Function Save(ByVal EditMode As EditMode, ByVal ShouldReg As Boolean,
                         ByVal WipData As DataSet) As Boolean
        Try
            Dim ria As RIAResult = Save(EditMode, ShouldReg, WipData, Nothing, False)
            Return ria.ResultBoolean
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function Save(ByVal EditMode As EditMode, ByVal ShouldReg As Boolean,
                         ByVal WipData As DataSet, ByVal WipInstData As DataSet) As Boolean
        Try
            Dim ria As RIAResult = Save(EditMode, ShouldReg, WipData, WipInstData, False)
            Return ria.ResultBoolean
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function Save(ByVal EditMode As EditMode, ByVal ShouldReg As Boolean,
                         ByVal WipData As DataSet, ByVal ReturnRIA As Boolean) As RIAResult
        Return Save(EditMode, ShouldReg, WipData, Nothing, ReturnRIA)
    End Function
    Public Function Save(ByVal EditMode As EditMode, ByVal ShouldReg As Boolean,
                     ByVal WipData As DataSet, ByVal WipInstData As DataSet,
                     ByVal ReturnRIA As Boolean) As RIAResult
        Return Save(EditMode, ShouldReg, WipData, WipInstData, ReturnRIA, Nothing)
    End Function
    Public Function Save(ByVal EditMode As EditMode, ByVal ShouldReg As Boolean,
                         ByVal WipData As DataSet, ByVal WipInstData As DataSet,
                         ByVal ReturnRIA As Boolean, ByVal MoveFaciData As DataSet) As RIAResult
        SOUtil = New CableSoft.SO.BLL.Utility.Utility(LoginInfo, DAO)
        Dim trans As DbTransaction = Nothing
        Dim MyTrans As Boolean
        Dim cn As System.Data.Common.DbConnection = DAO.GetConn()
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim AutoCloseConn As Boolean = DAO.AutoCloseConn
        Dim dtContact As DataTable = Nothing
        Dim result As RIAResult = Nothing
        Try
            Using WipUtil As New CableSoft.SO.BLL.Wip.Utility.SaveData(LoginInfo, DAO)
                Using Wip As DataTable = WipData.Tables("Wip")
                    CSLog = New CableSoft.SO.BLL.DataLog.DataLog(LoginInfo, DAO)
                    If DAO.Transaction Is Nothing Then
                        MyTrans = True
                        If String.IsNullOrEmpty(cn.ConnectionString) Then cn.ConnectionString = Me.LoginInfo.ConnectionString
                        If cn.State <> ConnectionState.Open Then cn.Open()
                        trans = cn.BeginTransaction
                        DAO.AutoCloseConn = False
                    Else
                        MyTrans = False
                        If cn.State = ConnectionState.Closed Then
                            cn.Open()
                        End If
                        trans = DAO.Transaction
                    End If
                    DAO.Transaction = trans

                    '(1)拆復異動資料(SO001/SO002/SO003) 當新增且派工類別為5時, 則需要客戶主
                    '   A.CMCode = Select CodeNo ,Description From CD031 Where CodeNo = <SO044.CMCode>
                    '   B.PTCode = Select CodeNo ,Description From CD032 Where CodeNo = 1
                    'C.	當收費參數.Para26 = 1 則需將週期性收費項目帳號及發票清成預設:
                    '   1.Update SO033 Set CMCode=<CMCode.CodeNo>,CMName=<CMCode.Description>,PTCode=<PTCode.CodeNo>,PTName=<PTCode.Description>,BankCode=null,BankName=null,AccountNo=null,InvSeqNo=null Where CustId = <CustId> And ServiceType = <ServiceType>
                    'D.	當收費參數.ClearInvDat = 1 則需將客戶主檔的帳號及發票清成預設:
                    '   1.Update SO001 Set CMCode=<CMCode.CodeNo>,CMName=<CMCode.Description>,InvoiceType=<SO044.InvoiceType>,InvNo=null,InvTitle=null,InvAddress=null,InvPurposeCode=null,InvPurposeName=null,InvoiceKind=0,Email=null,DenRecCode=null,DenRecName=null,DenRecDate=null,MailAddrNo=InstAddrNo,MailAddress=InstAddress,CustNote=null,ChargeNote=null Where CustId = <CustId>
                    Dim WipRow As DataRow = Wip.Rows(0)
                    Dim WipCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PRCode, WipRow.Item("PRCode").ToString, True)
                    Dim WipRefNo As Integer
                    Dim OldSNo As String = WipRow.Item("Sno")

                    If Not WipCode.Rows(0).IsNull("RefNo") Then
                        WipRefNo = Integer.Parse(WipCode.Rows(0).Item("RefNo"))
                    Else
                        WipRefNo = 0
                    End If

                    If WipData IsNot Nothing Then
                        If WipData.Tables.Contains("Contact") Then
                            dtContact = WipData.Tables("Contact").Copy
                        End If
                    End If

                    '2014.11.21 因為拆機工單Client端最後SAVE時有 _dsPRData.AcceptChanges() ，裡面有備註原因。
                    '           所以處理方式就是判斷"修改"工單，就針對工單資料，先利用該工單號碼取得資料庫內的資料，再將資料由DataSet填寫回去，並將資料庫的Dara取代DataSet內。
                    If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                        Dim dbWipdata As DataTable = DAO.ExecQry(_DAL.GetDBWip, New Object() {OldSNo})
                        If dbWipdata.Rows.Count > 0 Then
                            dbWipdata.TableName = "WIP"
                            For Each dbWip As DataRow In dbWipdata.Rows
                                If dbWip("SNO") = WipRow("SNO") Then
                                    For Each dbColumns As DataColumn In dbWipdata.Columns
                                        If dbWipdata.Columns.Contains(dbColumns.ColumnName) Then
                                            If WipRow.Table.Columns.Contains(dbColumns.ColumnName) Then
                                                If dbWip(dbColumns.ColumnName).ToString <> WipRow(dbColumns.ColumnName).ToString Then
                                                    dbWip(dbColumns.ColumnName) = WipRow(dbColumns.ColumnName)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                            WipData.Tables.Remove("Wip")
                            WipData.Tables.Add(dbWipdata.Copy)
                            WipRow = WipData.Tables("Wip").Rows(0)
                        End If
                    End If

                    If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                        If WipRefNo = 5 Then
                            If Not PRChangeCustomerInv(WipCode, WipRow) Then
                                Throw New Exception("ChangeAccountInvoice")
                                'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "ChangeAccountInvoice"}
                            End If
                        End If
                        '#6721 2014.06.05 增加判斷新增工單時，要將指定設備的PinCode回填到工單內
                        If Not PRChangeFaciPinCodeToWip(EditMode, WipRow, WipData) Then
                            Throw New Exception("PRChangeFaciPinCodeToWip")
                            'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "PRChangeFaciPinCodeToWip"}
                        End If
                    End If

                    '停復異動資料
                    If WipRefNo = 7 Then
                        If Not StopChangeData(WipRow) Then
                            Throw New Exception("StopChangeData")
                            'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "StopChangeData"}
                        End If
                    End If
                    '(3)更新工單資料: 呼叫CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeWip。
                    If Not WipUtil.ChangeWip(EditMode, BLL.Utility.InvoiceType.PR, WipData, ShouldReg) Then
                        Throw New Exception("Wip.Utility.ChangeWip")
                        'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeWip"}
                    End If
                    ''(4)更新設備(裝): 呼叫CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeFacility。
                    '2014.03.19 by Corey 因為此功能Utility已經都有直接處理了，所以不需要再呼叫。
                    'If Not WipUtil.ChangeFacility(EditMode, BLL.Utility.InvoiceType.PR, WipData) Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeFacility"}
                    'End If
                    ''(5)更新設備(拆): 呼叫CableSoft.SO.BLL.Wip.Utility.SaveData.ChangePRFacility。
                    '2014.03.19 by Corey 因為此功能Utility已經都有直接處理了，所以不需要再呼叫。
                    'If Not WipUtil.ChangePRFacility(EditMode, BLL.Utility.InvoiceType.PR, WipData) Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeFacility"}
                    'End If
                    ''(6)更新指定設備: 呼叫CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeChangeFacility。
                    '2014.03.19 by Corey 因為此功能Utility已經都有直接處理了，所以不需要再呼叫。
                    'If Not WipUtil.ChangeChangeFacility(EditMode, BLL.Utility.InvoiceType.PR, WipData) Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeChangeFacility"}
                    'End If
                    '(7)更新收費資料: 呼叫CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeCharge。
                    If Not WipUtil.ChangeCharge(EditMode, ShouldReg, WipData) Then
                        Throw New Exception("Wip.Utility.ChangeCharge")
                        'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeCharge"}
                    End If

                    '拆機不用處理 (8)更新結清資料: 呼叫CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeCloseData
                    'If Not WipUtil.ChangeCloseData(EditMode, BLL.Utility.InvoiceType.PR, WipData) Then
                    '    Throw New Exception("Wip.Utility.ChangeCloseData")
                    'End If
                    ''(9)更新預約明細(SO010): CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeResvDetail
                    '2014.03.19 by Corey 因為此功能Utility已經都有直接處理了，所以不需要再呼叫。
                    'If Not WipUtil.ChangeResvDetail(WipRow.Item("SNO"), WipRow.Item("CompCode"), WipRow.Item("ResvTime"),
                    '                                WipRow.Item("WorkServCode"), WipRow.Item("ServiceType"), Not WipRow.IsNull("SignDate")) Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeResvDetail"}
                    'End If
                    ''(10)更新裝機未完一覽表(SO072): CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeResvLog
                    '2014.03.19 by Corey 因為此功能Utility已經都有直接處理了，所以不需要再呼叫。
                    'If Not WipUtil.ChangeResvLog(BLL.Utility.InvoiceType.PR, Wip) Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeResvLog"}
                    'End If
                    '(11)回填已送命令工單單號(SO314/SO307/SO180/SEND_NAGRA/SO555/SO005B/SOAC0202/SOAC0202TMP): CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeCommandData
                    If Not WipUtil.ChangeCommandData(OldSNo, WipData.Tables("Wip")) Then
                        Throw New Exception("Wip.Utility.ChangeCommandData")
                        'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeCommandData"}
                    End If
                    ''(12)	更新派工服務件數(SO083/SO083A): CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeResvTempPoint
                    '2014.03.19 by Corey 因為此功能Utility已經都有直接處理了，所以不需要再呼叫。
                    'If Not WipUtil.ChangeResvTempPoint(BLL.Utility.InvoiceType.PR, Wip) Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeResvTempPoint"}
                    'End If
                    ''(13)	刪除派工服務件數暫存檔(SO085): CableSoft.SO.BLL.Wip.Utility.SaveData.DelResvPoint
                    '2014.03.19 by Corey 因為此功能Utility已經都有直接處理了，所以不需要再呼叫。
                    'If Not WipUtil.DelResvPoint() Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.DelResvPoint"}
                    'End If
                    '拆機不用處理 (14)	更新客戶促銷明細檔(SO098): 如工單有促銷方案及消息來源, 則做以下動作:
                    'A.	檢查客編/服務別/促銷方案/消息來源於SO098存不存在,不存在則新增, 回填欄位如B所列。
                    'B.	ServiceType=服務別,CompCode=公司別,CustId=客編,BulletinCode/BulletinName=消息來源, MediaCode/MediaName=介紹媒介, PromCode/PromName=促銷方案, ProcDate=受理時間
                    'If Not ChangeCustomerPromData(Wip) Then 
                    '    Throw New Exception("ChangeCustomerPromData")
                    'End If
                    ''(15)更新訂單資訊(SO105): CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeOrderData
                    'If Not WipUtil.ChangeOrderData(BLL.Utility.InvoiceType.PR, Wip) Then
                    '    Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Wip.Utility.ChangeOrderData"}
                    'End If
                    '(16)更新業務預約資訊(SO100): CableSoft.SO.BLL.Wip.Utility.SaveData.ChangeSalesData

                    '(17)工單完工或退單時需要回填SO313的資料
                    If Not GetUpdataSO313(EditMode, WipRow) Then
                        Throw New Exception("Update SO313 : OSNOSTATUS")
                        'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Update SO313 : OSNOSTATUS"}
                    End If

                    '(17-1)同區移機 需要將設備的相關狀態資料給Update
                    If Not PrMoveToFacility(EditMode, WipData) Then
                        Throw New Exception("WipPR Move to WipInstall Combo")
                        'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "WipPR Move to WipInstall Combo"}
                    End If

                    '(18)完工時間由未完工到完工,需判斷計價設備如果最後一台，則需要更新SO002.PRTIME
                    If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipRow, "FinTime") Then
                        If Not ChkFaciToUpd002(WipData) Then
                            Throw New Exception("Check Facility is Zero to Update SO002.PRTime")
                            'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Check Facility is Zero to Update SO002.PRTime"}
                        End If
                    End If

                    '(19)完工時間由未完工到完工且拆機參考號為2,3,5,6才做以下動作:
                    If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipRow, "FinTime") AndAlso
                        ",2,3,4,5,6,".Contains(String.Format(",{0},", WipRefNo)) Then
                        If ",2,3,4,5,".Contains(String.Format(",{0},", WipRefNo)) Then
                            '更新大樓資料(SO017)
                            Dim strErrMsg As String = String.Empty
                            If Not UpdMduidData(WipRow, WipRefNo, strErrMsg) Then
                                Throw New Exception("Update SO017 : MduId Data" & strErrMsg)
                                'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Update SO017 : MduId Data" & strErrMsg}
                            End If
                            '更新地址客戶歷史檔(SO015)
                            If Not ChangeAddress(WipRefNo, WipRow) Then
                                Throw New Exception("Update SO015 : ChangeAddress")
                                'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Update SO015 : ChangeAddress"}
                            End If
                            'C.	更新客戶基本資料(SO001)(派工參考號 3 才做):
                            If Not UpdCustomerData(WipRow, WipRefNo) Then
                                Throw New Exception("Update SO001 : Customer Data")
                                'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Update SO001 : Customer Data"}
                            End If
                            '更新街道資料(CD017)
                            If Not UpdStrtData(WipRow, WipRefNo) Then
                                Throw New Exception("Update CD017 : StrtCode Count")
                                'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Update CD017 : StrtCode Count"}
                            End If
                        End If
                        '#6721 工單完工時 並且 工單參考號=6，才需要更新PinCode(SO004) 
                        If WipRefNo = 6 Then
                            If Not UpdPinCode(EditMode, WipRow, WipData.Tables("ChangeFacility")) AndAlso Not WipRow.IsNull("PinCode") Then
                                Throw New Exception("Update SO004 : PinCode")
                                'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Update SO004 : PinCode"}
                            End If
                        End If
                    End If

                    '(20) 工單退單需要做個別處理事項
                    If Not ReturnWip(EditMode, WipRow, ShouldReg, WipRefNo) Then
                        Throw New Exception("Return Wip and OtherWip Error")
                        'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Return Wip and OtherWip Error"}
                    End If

                    '(End After 21)同區移機有裝機工單，須呼叫裝機存檔。
                    If WipInstData IsNot Nothing Then
                        '(20-1)同區移機需要 先將中間檔(SO313)資料填入相關資料
                        If Not PrMoveToInstall(EditMode, WipData, WipInstData) Then
                            Throw New Exception("WipPR Move to WipInstall Combo")
                            'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "WipPR Move to WipInstall Combo"}
                        End If

                        Using ObjInstSave As New CableSoft.SO.BLL.Wip.Install.SaveData(LoginInfo, DAO)
                            '2012.07.10 因為工單同區移機時，裝機工單要有拆機工單資料，所以需要將WipPR的資料寫入WipInstall，TableName=MovePRData
                            Dim WipPRtoInst As DataTable = WipData.Tables("Wip").Copy
                            WipPRtoInst.TableName = "MovePRData"
                            WipInstData.Tables.Add(WipPRtoInst)
                            If Not ObjInstSave.Save(EditMode, ShouldReg, WipInstData) Then
                                Throw New Exception("Update Install Wip Error.")
                                'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "Update Install Wip Error."}
                            End If
                        End Using
                    End If

                    Using SOWipUtil As New CableSoft.SO.BLL.Utility.Wip(LoginInfo, DAO)
                        Dim RetCode As Int16 = 0
                        Dim P_RETMSG As String = ""
                        RetCode = SOWipUtil.SF_ADJSTATUS1(Nothing, WipRow.Item("CustId"), 1, 0, WipRow.Item("CompCode"),
                                                          WipRow.Item("ServiceType"), P_RETMSG)
                        '更新客戶狀態(SF_ADJSTATUS1)
                        If RetCode < 0 Then
                            Throw New Exception(String.Format("Wip.SF_ADJSTATUS1-ReturnCode:{0},ReturnMessage:{1}", RetCode, P_RETMSG))
                            'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = String.Format("Wip.SF_ADJSTATUS1-ReturnCode:{0},ReturnMessage:{1}", RetCode, P_RETMSG)}
                        End If
                    End Using

                    '因為呼叫WipUtil.ChangeWip 共用的檢核程式裡面有LOG一次，因此外層就不需要再LOG一次
                    'Dim aResult As RIAResult = CSLog.SummaryExpansion(DataLog.OpType.Update, "SO009", Wip, Int32.Parse(Integer.Parse(DateTime.Now.ToString("yyyyMMdd"))))
                    'If Not aResult.ResultBoolean Then
                    '    Select Case aResult.ErrorCode
                    '        Case -6
                    '            If MyTrans Then trans.Rollback()
                    '            Return False
                    '    End Select
                    'End If

                    Using ControlCMD As New CableSoft.SO.BLL.Wip.ControlCommand.ControlCommand(LoginInfo, DAO)
                        Dim RiaSendCmd As RIAResult = New RIAResult
                        RiaSendCmd = ControlCMD.Execute(EditMode, WipData)
                        If RiaSendCmd.ResultBoolean Then
                            WipData = RiaSendCmd.ResultDataSet
                        Else
                            Throw New Exception("SendCmd Error: " & RiaSendCmd.ErrorMessage)
                            'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "SendCmd Error: " & RiaSendCmd.ErrorMessage}
                        End If
                    End Using

                    If Not WipUtil.ChangeWipFinalProcess(EditMode, BLL.Utility.InvoiceType.PR, WipData) Then
                        Throw New Exception("WipUtil.ChangeWipFinalProcess")
                        'Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = -999, .ErrorMessage = "WipUtil.ChangeWipFinalProcess"}
                    End If
                    '2014/05/05 Jacky 6729 Luke 移機將其他客戶底下全服務移機功能
                    If MoveFaciData IsNot Nothing Then
                        Dim gresult As RIAResult = OtherServiceMovePR(EditMode, WipData, MoveFaciData)
                        If gresult.ResultBoolean = False Then
                            Throw New Exception(gresult.ErrorMessage)
                            'Return result
                        End If
                    End If
                    WipCode.Dispose()
                    If MyTrans Then
                        trans.Commit()
                    End If
                End Using
            End Using
            result = New RIAResult() With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = String.Empty}
        Catch ex As Exception
            If MyTrans Then trans.Rollback()
            result = New RIAResult() With {.ResultBoolean = False, .ErrorCode = -99, .ErrorMessage = ex.ToString()}
        Finally
            If MyTrans Then
                cn.Close()
                DAO.AutoCloseConn = AutoCloseConn
                DAO.Dispose()
                cn.Dispose()
                trans.Dispose()
            End If
            SOUtil.Dispose()
        End Try
        Return result
    End Function

    Private Function ReturnSaveFunc(MyTrans As Boolean, ErrorCode As Integer, ErroeMessage As String,
                                    trans As DbTransaction, cn As DbConnection, AutoCloseConn As Boolean) As RIAResult
        '2014.12.15 因為Save功能如果直接回傳
        If MyTrans Then
            trans.Rollback()
            cn.Close()
            DAO.AutoCloseConn = AutoCloseConn
            DAO.Dispose()
            cn.Dispose()
            trans.Dispose()
        End If
        SOUtil.Dispose()
        Return New RIAResult() With {.ResultBoolean = False, .ErrorCode = ErrorCode, .ErrorMessage = ErroeMessage}
    End Function



#Region "同區移機功能 "
    ''' <summary>
    ''' 新增工單要填寫相關資料 工單退單需要清除相關資料
    ''' </summary>
    ''' <param name="EditMode">工單狀態</param>
    ''' <param name="WipPR">拆機工單資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PrMoveToFacility(EditMode As EditMode, ByVal WipPR As DataSet) As Boolean
        Try
            Dim WipPRRow As DataRow = WipPR.Tables("Wip").Rows(0)
            Dim ReInstOwner As String = DAO.ExecQry(_DAL.GetSO041).Rows(0)("ReInstOwner").ToString
            Dim dtSO313 As DataTable = DAO.ExecQry(_DAL.GetSO313(ReInstOwner), New Object() {LoginInfo.CompCode, WipPRRow("CustId"), WipPRRow("SNO")})
            For Each drSO313 As DataRow In dtSO313.Rows
                '新增 需要將設備資料註記
                If EditMode = CableSoft.BLL.Utility.EditMode.Append Then
                    For Each drSO004D As DataRow In WipPR.Tables("ChangeFacility").Rows
                        DAO.ExecNqry("Update SO004 Set PRFlag = 1 Where " & _DAL.PrMoveToFacility,
                                     New Object() {dtSO313.Rows(0)("OCustId"), LoginInfo.CompCode, dtSO313.Rows(0)("ServiceType"), drSO004D("SEQNO")})
                    Next
                End If
                If EditMode = CableSoft.BLL.Utility.EditMode.Edit Then
                    '退單需將設備資料還原
                    If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipPRRow, "SignDate") Then
                        If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipPRRow, "ReturnCode") Then
                            For Each drSO004D As DataRow In WipPR.Tables("ChangeFacility").Rows
                                DAO.ExecNqry("Update SO004 Set PRFlag = 0 ,PRDate = Null, GetDate = Null, PRSNo = null Where " & _DAL.PrMoveToFacility,
                                             New Object() {dtSO313.Rows(0)("OCustId"), LoginInfo.CompCode, dtSO313.Rows(0)("ServiceType"), drSO004D("SEQNO")})
                            Next
                        End If
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 產生同區移出單時要產生同區移入單時需要有中間檔(SO313)
    ''' </summary>
    ''' <param name="EditMode">工單狀態</param>
    ''' <param name="WipPR">拆機工單資料</param>
    ''' <param name="WipInst">裝機工單資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PrMoveToInstall(EditMode As EditMode, ByVal WipPR As DataSet, ByVal WipInst As DataSet) As Boolean
        Try
            '新增模式才需要填寫SO313資料
            If EditMode <> CableSoft.BLL.Utility.EditMode.Append Then Return True
            Dim ReInstOwner As String = DAO.ExecQry(_DAL.GetSO041).Rows(0)("ReInstOwner").ToString
            Dim dtSO313 As DataTable = DAO.ExecQry(_DAL.GetSO313(ReInstOwner), New Object() {LoginInfo.CompCode, WipPR.Tables("Wip").Rows(0)("CustId"), WipPR.Tables("Wip").Rows(0)("SNO")})
            Dim drSO313 As DataRow = Nothing
            Dim dtCustData As DataTable = Nothing
            Dim blnAddNew As Boolean = False
            Dim strWhere As String = "0=1"
            Dim drWipPR As DataRow = WipPR.Tables("Wip").Rows(0)
            Dim drWipInst As DataRow = WipInst.Tables("Wip").Rows(0)
            Dim dtProposer As DataRow = Nothing '申請人資料

            If dtSO313.Rows.Count > 0 Then
                strWhere = String.Format("OCompCode={0} and OCustID={1} and OSNO='{2}'", LoginInfo.CompCode, WipPR.Tables("Wip").Rows(0)("CustID"), WipPR.Tables("Wip").Rows(0)("SNO"))
                drSO313 = dtSO313.Rows(0)
            Else
                drSO313 = dtSO313.NewRow
                dtSO313.Rows.Add(drSO313)
                blnAddNew = True
            End If

            For Each dtTmp As DataTable In WipPR.Tables
                If dtTmp.TableName.ToUpper = "SO001".ToUpper Then '2014.09.12 原本用Customer 改用SO001的就可以。
                    If Not IsDBNull(dtTmp.Rows(0)("ID")) Then
                        Dim Cust_ID As String = dtTmp.Rows(0)("ID")
                        If Not String.IsNullOrEmpty(Cust_ID) Then
                            dtCustData = DAO.ExecQry(_DAL.GetSO137, New Object() {Cust_ID})
                        End If
                    End If
                End If
            Next

            With drSO313
                Dim CompName As String = DAO.ExecQry(_DAL.GetCD039, New Object() {LoginInfo.CompCode}).Rows(0)("Description").ToString
                .Item("OCompCode") = LoginInfo.CompCode
                .Item("OCompName") = CompName
                .Item("OCustId") = drWipPR("CustId")
                .Item("OSNO") = drWipPR("SNO")
                .Item("OAddrNo") = drWipPR("OldAddrNo")
                .Item("OAddress") = drWipPR("OldAddress")
                .Item("NCompCode") = LoginInfo.CompCode
                .Item("NCompName") = CompName
                .Item("NCustId") = drWipInst("CustId")
                .Item("NSNO") = drWipInst("SNO")
                .Item("NAddrNo") = drWipInst("AddrNo")
                .Item("NAddress") = drWipInst("Address")
                .Item("ServiceType") = drWipPR("ServiceType")
                .Item("CStatus") = "同區移機"
                .Item("UpdTime") = DateTimeUtility.GetDTString(DateTime.Now)
                .Item("UpdEN") = LoginInfo.EntryId
                .Item("CustName") = drWipPR("CustName")
                .Item("Tel1") = drWipPR("Tel1")
                If dtCustData IsNot Nothing Then
                    If dtCustData.Rows.Count > 0 Then
                        .Item("Birthday") = dtCustData.Rows(0)("Birthday")
                        .Item("ID") = dtCustData.Rows(0)("ID")
                    End If
                End If
            End With
            dtSO313.AcceptChanges()
            If blnAddNew Then
                If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.InsertRow, dtSO313, "SO313", strWhere, True, , False) Then
                    Return False
                End If
            Else
                If dtSO313.Rows.Count > 0 Then
                    If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, dtSO313, "SO313", strWhere, True, , False) Then
                        Return False
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 完工或退單需要連動將SO313.OSNOSTATUS 填入(完工或退單)
    ''' </summary>
    ''' <param name="EditMode">工單狀態</param>
    ''' <param name="WipPR">拆機工單資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetUpdataSO313(EditMode As EditMode, ByVal WipPR As DataRow) As Boolean
        Try
            If EditMode = CableSoft.BLL.Utility.EditMode.Edit AndAlso Not WipPR.IsNull("SignDate") Then
                Dim ReInstOwner As String = DAO.ExecQry(_DAL.GetSO041).Rows(0)("ReInstOwner").ToString
                Dim blnUpdate As Boolean = False
                Using dtSO313 As DataTable = DAO.ExecQry(_DAL.GetSO313(ReInstOwner), New Object() {LoginInfo.CompCode, WipPR("CustId"), WipPR("SNO")})
                    For Each drRow As DataRow In dtSO313.Rows
                        Dim StatusName As String = Nothing
                        If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipPR, "FinTime") Then StatusName = SaveDataLanguage.WipRunStatus1
                        If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipPR, "ReturnCode") Then StatusName = SaveDataLanguage.WipRunStatus0
                        If Not String.IsNullOrEmpty(StatusName) Then
                            blnUpdate = True
                            drRow("OSNOSTATUS") = StatusName
                            drRow("UpdTime") = DateTimeUtility.GetDTString(DateTime.Now)
                            drRow("UpdEN") = LoginInfo.EntryId
                        End If
                    Next
                    If dtSO313.Rows.Count > 0 AndAlso blnUpdate Then
                        Dim strWhere As String = String.Format("OCompCode={0} and OCustID={1} and OSNO='{2}'", LoginInfo.CompCode, WipPR("CustID"), WipPR("SNO"))
                        If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, dtSO313, "SO313", strWhere, True, , False) Then
                            Return False
                        End If
                    End If
                End Using
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "工單退單處理 SO042(ReInstSyncReturn,FaciBackSyncReturn)"
    Private Function ReturnWip(EditMode As EditMode, WipRow As DataRow, ShouldReg As Boolean, WipRefNo As String) As Boolean
        Try
            If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipRow, "ReturnCode") Then
                '１。SO042.FaciBackSyncReturn 拆機退單同步退取回單
                Dim dtSO042 As DataTable = SOUtil.GetSystem(BLL.Utility.SystemTableType.Wip, "ReInstSyncReturn,FaciBackSyncReturn", WipRow("ServiceType"))
                If dtSO042.Rows(0)("FaciBackSyncReturn") = 1 AndAlso ",2,5,6,".Contains(String.Format(",{0},", WipRefNo)) Then
                    Dim rtnWipData As DataTable = DAO.ExecQry(_DAL.GetRtnWip, New Object() {WipRow("Custid"), WipRow("SNO")})
                    For Each rtnWipRow As DataRow In rtnWipData.Rows
                        Dim rtnWip As DataSet = Nothing
                        Using WipUtil As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
                            rtnWip = WipUtil.GetWipCalculateData(BLL.Utility.InvoiceType.PR, rtnWipRow("CustId"), rtnWipRow("ServiceType"), rtnWipRow("SNo"), rtnWipRow("ResvTime"), rtnWipRow("PrCode"))
                            Dim WipRow2 As DataRow = rtnWip.Tables("Wip").Rows(0)
                            With WipRow2
                                .Item("ReturnCode") = WipRow("ReturnCode")
                                .Item("ReturnName") = WipRow("ReturnName")
                                .Item("SignDate") = WipRow("SignDate")
                                .Item("SignEn") = WipRow("SignEn")
                                .Item("SignName") = WipRow("SignName")
                                .Item("CallOkTime") = DateTime.Parse(String.Format("{0} {1:HH:mm:ss}", DateTime.Parse(.Item("SignDate")).ToString("yyyy/MM/dd"), DateTime.Now))
                                .Item("Note") = IIf(.Item("Note") & "" = "", "", "; ") & SaveDataLanguage.WipPRandReturn
                                .Item("UpdTime") = WipRow("UpdTime")
                                .Item("UpdEn") = WipRow("UpdEn")
                                .Item("NewUpdTime") = DateTime.Now
                            End With
                            If Not Save(EditMode, ShouldReg, rtnWip) Then
                                Throw New Exception("SaveData.ReturnWip")
                                Exit Function
                            End If
                        End Using
                    Next
                End If

                '２。SO042.ReInstSyncReturn 同區移機連動退單
                If dtSO042.Rows(0)("ReInstSyncReturn") = 1 Then
                    Dim ReInstOwner As String = DAO.ExecQry(_DAL.GetSO041).Rows(0)("ReInstOwner").ToString
                    Dim dtSO313 As DataTable = DAO.ExecQry(_DAL.GetSO313(ReInstOwner), New Object() {LoginInfo.CompCode, WipRow("Custid"), WipRow("SNO")})
                    For Each drSO313 As DataRow In dtSO313.Rows
                        Using Inst As New CableSoft.SO.BLL.Wip.Install.Install(LoginInfo, DAO)
                            Using WipInst As DataSet = Inst.GetInstallData(drSO313("NSNO"))
                                If WipInst.Tables("Wip").Rows.Count > 0 Then
                                    If WipInst.Tables("Wip").Rows(0).IsNull("ReturnCode") Then
                                        With WipInst.Tables("Wip").Rows(0)
                                            .Item("ReturnCode") = WipRow("ReturnCode")
                                            .Item("ReturnName") = WipRow("ReturnName")
                                            .Item("CallOkTime") = DateTime.Parse(WipRow("SignDate")).AddTicks(DateTime.Now.TimeOfDay.Ticks)
                                            .Item("Note") = String.Format("{0}{1}{2}", .Item("Note"), IIf(.Item("Note") & "" = "", "", "; "), SaveDataLanguage.WipPRRemoveAndReturn)
                                            .Item("SignDate") = WipRow("SignDate")
                                            .Item("SignEn") = WipRow("SignEn")
                                            .Item("SignName") = WipRow("SignName")
                                        End With

                                        Using PRFacility As DataTable = WipInst.Tables("PRFacility")
                                            For Each drPRFacility As DataRow In PRFacility.Rows
                                                drPRFacility.Item("GetDate") = DBNull.Value
                                            Next
                                        End Using
                                        '呼叫拆機工單存檔
                                        Using InstSave As New CableSoft.SO.BLL.Wip.Install.SaveData(LoginInfo, DAO)
                                            If Not InstSave.Save(EditMode.Edit, False, WipInst) Then
                                                Throw New Exception("SaveData.ReturnWip.GetReInstInstData")
                                            End If
                                        End Using
                                    End If
                                End If
                            End Using
                        End Using
                    Next
                End If
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "更新客戶基本資料(SO001)"
    Friend Function UpdCustomerData(WipData As DataRow, WipRefNo As String) As Boolean
        'C.	更新客戶基本資料(SO001)(派工參考號 3 才做):
        '1.	NInst = Select A.*,B.ClctMethod From SO014 A,SO017 B Where A.MduId = B.MduId(+) And A.AddrNo = <新裝機地址編>。
        '2.	異動欄位如下: InstAddrNo= <新裝機地址編>, InstAddress = <新裝機地址>, ChargeAddrNo = <新收費地址編>, ChargeAddress= <新收費地址>, 
        '                 MailAddrNo = <新郵寄地址編>, MailAddress = <新郵寄地址>, ServCode= <NInst.ServCode>, ServArea= <NInst.ServName>, 
        '                 ClctAreaCode = <NInst.ClctAreaCode>, ClctAreaName= <NInst.ClctAreaName>, MduId= <NInst.MduId>, 
        '                 ChargeType = (Decode(<NInst.ClctMethod>,1,3,2,2,3,2,1),UpdTime = <異動時間>, UpdEn = <操作人員>。
        '3.	Update SO001 Set <異動欄位> Where CustId = <客戶編號>。
        Try
            If WipRefNo = 3 Then
                Using NInst As DataTable = DAO.ExecQry(_DAL.GetNewAddress, New Object() {WipData("ReInstAddrNo")})
                    Using dtSO001 As DataTable = DAO.ExecQry(_DAL.GetSO001, New Object() {WipData("CustId")})
                        If dtSO001.Rows.Count > 0 AndAlso NInst.Rows.Count > 0 Then
                            With dtSO001.Rows(0)
                                .Item("InstAddrNo") = WipData("ReInstAddrNo")
                                .Item("InstAddress") = WipData("ReInstAddress")
                                .Item("ChargeAddrNo") = WipData("NewChargeAddrNo")
                                .Item("ChargeAddress") = WipData("NewChargeAddress")
                                .Item("MailAddrNo") = WipData("NewMailAddrNo")
                                .Item("MailAddress") = WipData("NewMailAddress")
                                .Item("ServCode") = NInst.Rows(0)("ServCode")
                                .Item("ServArea") = NInst.Rows(0)("ServName")
                                .Item("ClctAreaCode") = NInst.Rows(0)("ClctAreaCode")
                                .Item("ClctAreaName") = NInst.Rows(0)("ClctAreaName")
                                .Item("MduId") = NInst.Rows(0)("MduId")
                                .Item("ChargeType") = NInst.Rows(0)("ClctMethod")
                                .Item("UpdTime") = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
                                .Item("UpdEn") = LoginInfo.EntryName
                                .Item("NewUpdTime") = DateTime.Now
                            End With
                            If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, dtSO001, "SO001", "Custid=" & dtSO001.Rows(0)("Custid"), False, , False) Then
                                Return False
                            End If
                        End If
                    End Using
                End Using
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "更新街道資料(CD017)"
    Private Function UpdStrtData(WipData As DataRow, WipRefNo As String) As Boolean
        Try
            'D.	服務別為C 則需更新街道資料(CD017):
            '1.	異動欄位如下: InstCnt = InstCnt –1(派工參考號 2, 4, 5 才做)
            '2.	異動欄位如下: InstCnt = InstCnt + 1(派工參考號 3 才做)。
            '3.	Update CD017 Set <異動欄位> Where CodeNo = <工單StrtCode>。
            If String.Compare(WipData("ServiceType"), "C", False) = 0 Then
                Dim AddrNo As String = String.Empty
                Dim IsAddCount As Boolean = False
                If ",2,4,5,".Contains(String.Format(",{0},", WipRefNo)) Then
                    AddrNo = WipData("OldAddrNo")
                Else
                    AddrNo = WipData("ReInstAddrNo")
                    IsAddCount = True
                End If

                Dim dtSO014 As DataTable = DAO.ExecQry(_DAL.GetSO014, New Object() {AddrNo})
                If dtSO014.Rows.Count > 0 Then
                    Using dtCD017 As DataTable = DAO.ExecQry(_DAL.GetCD017, New Object() {dtSO014.Rows(0)("StrtCode")}, "dsCD017", "CD017")
                        If dtCD017.Rows.Count > 0 Then
                            If IsAddCount Then
                                dtCD017.Rows(0)("InstCnt") += 1
                            Else
                                dtCD017.Rows(0)("InstCnt") -= 1
                            End If
                            If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, dtCD017, "CD017", "CodeNo=" & dtCD017.Rows(0)("CodeNO"), False, , False) Then
                                Return False
                            End If
                        End If
                    End Using
                End If
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "更新大樓戶數資料(SO017)"
    Private Function UpdMduidData(WipData As DataRow, WipRefNo As Integer, ByRef ErrMsg As String) As Boolean
        Try
            'A.	當SO001.MduId 有值 則需更新大樓資料(SO017): 
            '1.	更新欄位(派工參考號 2, 4, 5 才做): 
            'i.	    服務別= C: InstCnt  = InstCnt  - 1(SO001.ChargeType = 3 才做), PerInstCnt  = PerInstCnt  - 1(SO001.ChargeType <> 3 才做), UnInstCnt  = UnInstCnt  + 1 
            'ii.	服務別= I: InstCnt1 = InstCnt1 - 1(SO001.ChargeType = 3 才做), PerInstCnt1 = PerInstCnt1 - 1(SO001.ChargeType <> 3 才做), UnInstCnt1 = UnInstCnt1 + 1 
            'iii.	服務別= D: InstCnt2 = InstCnt2 - 1(SO001.ChargeType = 3 才做), PerInstCnt2 = PerInstCnt2 - 1(SO001.ChargeType <> 3 才做), UnInstCnt2 = UnInstCnt2 + 1 
            'iv.	服務別= P: InstCnt3 = InstCnt3 - 1(SO001.ChargeType = 3 才做), PerInstCnt3 = PerInstCnt3 - 1(SO001.ChargeType <> 3 才做), UnInstCnt3 = UnInstCnt3 + 1 
            '2.	更新欄位(派工參考號 3 才做): 
            'i.	    服務別= C: InstCnt  = InstCnt  + 1(SO001.ChargeType = 3 才做), PerInstCnt  = PerInstCnt  + 1(SO001.ChargeType <> 3 才做), UnInstCnt  = UnInstCnt  - 1 
            'ii.	服務別= I: InstCnt1 = InstCnt1 + 1(SO001.ChargeType = 3 才做), PerInstCnt1 = PerInstCnt1 + 1(SO001.ChargeType <> 3 才做), UnInstCnt1 = UnInstCnt1 - 1 
            'iii.	服務別= D: InstCnt2 = InstCnt2 + 1(SO001.ChargeType = 3 才做), PerInstCnt2 = PerInstCnt2 + 1(SO001.ChargeType <> 3 才做), UnInstCnt2 = UnInstCnt2 - 1 
            'iv.	服務別= P: InstCnt3 = InstCnt3 + 1(SO001.ChargeType = 3 才做), PerInstCnt3 = PerInstCnt3 + 1(SO001.ChargeType <> 3 才做), UnInstCnt3 = UnInstCnt3 - 1 
            '3.	Update SO017 Set <說明1的欄位> Where MduId = <SO001.MduId>
            Using dtSO001_Old As DataTable = DAO.ExecQry("Select * From SO001 Where Custid=" & WipData("Custid"))
                Dim dtSO001_New As DataTable = Nothing
                Dim HaveNewMdu As Boolean = False
                '新地址資料有對應到大樓
                If dtSO001_Old.Rows.Count = 0 Then
                    ErrMsg = SaveDataLanguage.PRtoCustNothing
                    Return False
                End If
                If Not IsDBNull(WipData("ReInstAddrNo")) Then
                    '新地址資料有值，需要填寫計算新地址資料
                    dtSO001_New = DAO.ExecQry(String.Format("Select * from SO001 Where InstAddrNo='{0}'", WipData("ReInstAddrNo")))
                    If dtSO001_New.Rows.Count > 0 Then
                        If Not IsDBNull(dtSO001_New.Rows(0)("MduID")) Then
                            HaveNewMdu = True
                        End If
                    End If
                End If
                '沒有大樓資料就不需要再填寫計算客戶數
                If Not IsDBNull(dtSO001_Old.Rows(0)("Mduid")) Then
                    If Not UpdSO017Count(WipData("ServiceType"), dtSO001_Old.Rows(0)("Mduid"), dtSO001_Old.Rows(0)("ChargeType"), WipRefNo, ErrMsg) Then
                        Return False
                    End If
                End If
                If HaveNewMdu Then
                    If Not IsDBNull(dtSO001_New.Rows(0)("Mduid")) Then
                        If Not UpdSO017Count(WipData("ServiceType"), dtSO001_New.Rows(0)("Mduid"), dtSO001_New.Rows(0)("ChargeType"), WipRefNo, ErrMsg) Then
                            Return False
                        End If
                    End If
                End If
            End Using
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function UpdSO017Count(ByVal ServiceType As String, ByVal Mduid As String,
                                   ByVal ChargeType As Int32, ByVal WipRefNo As String,
                                   ByVal ErrMsg As String) As Boolean
        Try
            Dim SerID As String = String.Empty
            Dim strWhere As String = String.Format("MduId = '{0}'", Mduid)
            Using dtSO017 As DataTable = DAO.ExecQry("Select * From SO017 Where " & strWhere)
                If dtSO017.Rows.Count = 0 Then
                    '沒有對應到大樓資料所以不需要對應
                    Return True
                End If
                Dim drSO017 As DataRow = dtSO017.Rows(0)
                Dim intAdd As Int32 = 1
                If ",2,4,5".Contains(String.Format(",{0},", WipRefNo)) Then intAdd = -1
                Select Case ServiceType
                    Case "C"
                    Case "D"
                        SerID = "1"
                    Case "I"
                        SerID = "2"
                    Case "P"
                        SerID = "3"
                End Select
                With drSO017
                    If ChargeType = 3 Then
                        .Item("InstCnt" & SerID) = .Item("InstCnt" & SerID) + intAdd
                    Else
                        .Item("PerInstCnt" & SerID) = .Item("PerInstCnt" & SerID) + intAdd
                    End If
                    .Item("UnInstCnt" & SerID) = .Item("UnInstCnt" & SerID) - intAdd
                End With
                dtSO017.AcceptChanges()
                If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, dtSO017, "SO017", strWhere, True, , False) Then
                    ErrMsg = SaveDataLanguage.UpdBuildingError
                    Return False
                End If
            End Using
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "拆機復機順完工拆機單"
    Private Function PRtoInstallFintime(EditMode As EditMode, ShouldReg As Boolean, WipData As DataRow) As Boolean
        Try
            '1.取該服務別的停拆移機單派工參考號為2,5 且 未結案 的工單(Wip/Facility/Charge/PRFacility/ChangeFacility)。
            '  2.將Wip 回填成已完工, 欄位內容如下: 
            '   i.FinTime = <工單FinTime-1分鐘>, CallOkTime= <工單FinTime-1分鐘>,SignDate = <工單SignDate>,SignEn = <工單SignEn>,SignName = <工單SignName>,FinUnit = WorkUnit, Note = Note || ‘拆機復機順完工’, UpdTime = <工單UpdTime>, UpdEn = <工單UpdEn>
            '  3.呼叫CableSoft.SO.BLL.Wip.PR.SaveData.Save()

            '步驟1
            Using OtherWip As DataTable = DAO.ExecQry(_DAL.GetOtherWip, New Object() {WipData("Custid"), WipData("ServiceType")}, False)
                OtherWip.TableName = "Wip"
                For Each WipDataRow As DataRow In OtherWip.Rows
                    Dim OtherWipData As DataSet = Nothing
                    Using WipUtil As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
                        '取得收費/設備資料
                        OtherWipData = WipUtil.GetWipCalculateData(BLL.Utility.InvoiceType.PR, WipDataRow("CustId"), WipDataRow("ServiceType"), WipDataRow("SNo"), WipDataRow("ResvTime"), WipDataRow("PRCODE"))
                        '步驟2
                        Dim OtherWipRow As DataRow = OtherWipData.Tables("Wip").Rows(0)
                        With OtherWipRow
                            .Item("FinTime") = DateTime.Parse(WipData("FinTime")).AddMinutes(-1)        '工單FinTime-1分鐘
                            .Item("CallOkTime") = DateTime.Parse(WipData("FinTime")).AddMinutes(-1)     '工單FinTime-1分鐘
                            .Item("SignDate") = WipData("SignDate")                                     '工單SignDate
                            .Item("SignEn") = WipData("SignEn")                                         '工單SignEn
                            .Item("SignName") = WipData("SignName")                                     '工單SignName
                            .Item("FinUnit") = WipData("WorkUnit")                                      '工單WorkUnit
                            .Item("Note") = WipData("Note") & SaveDataLanguage.WipPRandFin
                            .Item("UpdTime") = WipData("UpdTime")                                       '工單UpdTime
                            .Item("UpdEN") = WipData("UpdEN")                                           '工單UpdEn
                            .Item("NewUpdTime") = DateTime.Now
                        End With
                        '步驟3
                        If Not Save(EditMode, ShouldReg, OtherWipData) Then
                            Throw New Exception("SaveData.PRtoInstallFintime")
                            Exit Function
                        End If
                    End Using
                Next
                Return True
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "更新地址客戶歷史檔(SO015)"

    Private Function ChangeAddress(ByVal RefNo As Int32, WipData As DataRow) As Boolean
        Try
            'B.	更新地址客戶歷史檔(SO015):
            '1.	新增一筆資料至SO015 欄位如下(派工參考號 3 才做): SNO = <工單單號>, PinCode=<工單PinCode>, AddrNo=<工單AddrNo>, CustId=<客戶編號>, CustName=<工單CustName>, InDate=<完工時間取日期>, OutDate=null, PRFlag=null, Address=<工單Address>, CompCode=<公司別>。
            '2.	更新SO015(派工參考號 2,4,5才做): PrSNo=<工單號碼>, PrPinCode=<工單.PinCode>, OutDate=<完工時間>, PRFlag = Decode(<派工參考號>,4,4,2)。
            If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipData, "FinTime") Then
                Dim strWhere As String = String.Format("AddrNo = {0} And CompCode = {1} And CustId = {2} And PRFlag is Null ", WipData("OldAddrNo"), LoginInfo.CompCode, WipData("CustId"))
                Dim dtSO015 As DataTable
                Dim drSO015 As DataRow = Nothing
                Dim blnAddNew As Boolean = False

                If RefNo = 3 Then
                    blnAddNew = True
                    strWhere = "0 = 1"
                    dtSO015 = DAO.ExecQry(String.Format("Select SO015.RowId,SO015.* From SO015 Where {0} Order by  InDate Desc", strWhere))
                    drSO015 = dtSO015.NewRow
                    dtSO015.Rows.Add(drSO015)
                    With drSO015
                        .Item("CustID") = WipData("CustId")
                        .Item("CustName") = WipData("CustName")
                        .Item("PRFlag") = 3
                        .Item("SNo") = WipData("SNo")
                        .Item("PinCode") = WipData("PinCode")
                        .Item("InDate") = Date.Parse(WipData("FinTime")).ToString("yyyy/MM/dd")
                        .Item("AddrNo") = WipData("ReInstAddrNo")
                        .Item("Address") = WipData("ReInstAddress")
                        .Item("CompCode") = LoginInfo.CompCode
                    End With
                Else
                    dtSO015 = DAO.ExecQry(String.Format("Select SO015.RowId,SO015.* From SO015 Where {0} Order by  InDate Desc", strWhere))
                    If dtSO015.Rows.Count > 0 Then
                        drSO015 = dtSO015.Rows(0)
                        With drSO015
                            .Item("PrSNo") = WipData("SNo")
                            .Item("PrPinCode") = WipData("PinCode")
                            .Item("OutDate") = Date.Parse(WipData("FinTime")).ToString("yyyy/MM/dd")
                            If RefNo = 4 Then
                                .Item("PRFlag") = 4 '移拆
                            Else
                                .Item("PRFlag") = 2 '拆機
                            End If
                        End With
                    End If
                End If
                dtSO015.AcceptChanges()
                If blnAddNew Then
                    If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.InsertRow, dtSO015, "SO015", strWhere, True, , False) Then
                        Return False
                    End If
                Else
                    If dtSO015.Rows.Count > 0 Then
                        If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, dtSO015, "SO015", strWhere, True, , False) Then
                            Return False
                        End If
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "拆復異動資料(SO001/SO002/SO003)"
    Private Function PRChangeCustomerInv(ByVal PrCode As DataTable,
                                          ByVal PrDataRow As DataRow) As Boolean
        Try
            Dim ChargeSys As DataTable = SOUtil.GetSystem(BLL.Utility.SystemTableType.Charge, "Para26, ClearInvDat", PrDataRow("ServiceType"))
            Dim DefaultSys As DataTable = SOUtil.GetSystem(BLL.Utility.SystemTableType.DefaultS, "CMCode,InvoiceType", PrDataRow("ServiceType"))
            Dim DefCMCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.CMCode, DefaultSys.Rows(0).Item("CMCode").ToString, True)
            Dim DefPTCode As DataTable = SOUtil.GetCode(BLL.Utility.CodeType.PTCode, "1", True)
            '當收費參數.Para26 = 1 則需將週期性收費項目帳號及發票清成預設
            If ChargeSys.Rows(0).Item("Para26") = 1 Then
                Dim CyclePeriod As DataTable = DAO.ExecQry(_DAL.GetCyclePeriodInvDef())
                Dim CyclePeriodRow As DataRow = CyclePeriod.NewRow
                'CMCode,CMName,PTCode,PTName,BankCode,BankName,AccountNo,InvSeqNo
                With CyclePeriodRow
                    .Item("CMCode") = DefCMCode.Rows(0).Item("CodeNo")
                    .Item("CMName") = DefCMCode.Rows(0).Item("Description")
                    .Item("PTCode") = DefPTCode.Rows(0).Item("CodeNo")
                    .Item("PTName") = DefCMCode.Rows(0).Item("Description")
                    .Item("BankCode") = DBNull.Value
                    .Item("BankName") = DBNull.Value
                    .Item("AccountNo") = DBNull.Value
                    .Item("InvSeqNo") = DBNull.Value
                End With
                CyclePeriod.Rows.Add(CyclePeriodRow)
                If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO,
                    UpdateMode.UpdateRow, CyclePeriod, "SO003", "CustId = " & PrDataRow("CustId") &
                        " And ServiceType = '" & PrDataRow("ServiceType") & "'", , , False) Then
                    Throw New Exception("CyclePeriod")
                End If
                CyclePeriod.Dispose()
            End If
            '當收費參數.ClearInvDat = 1 則需將客戶主檔的帳號及發票清成預設
            If ChargeSys.Rows(0).Item("ClearInvDat") = 1 Then
                Dim Customer As DataTable = DAO.ExecQry(_DAL.GetCustomerInvDef(), New Object() {PrDataRow("CustId"), PrDataRow("CompCode"), PrDataRow("ServiceType")})
                Dim CustomerRow As DataRow = Customer.NewRow
                'CMCode,CMName,InvoiceType,InvNo,InvTitle,InvAddress,InvPurposeCode,
                'InvPurposeName,InvoiceKind,Email,DenRecCode,DenRecName,DenRecDate,CustNote,ChargeNote
                With CustomerRow
                    .Item("CMCode") = DefCMCode.Rows(0).Item("CodeNo")
                    .Item("CMName") = DefCMCode.Rows(0).Item("Description")
                    .Item("InvoiceType") = DefaultSys.Rows(0).Item("InvoiceType")
                    .Item("InvNo") = DBNull.Value
                    .Item("InvTitle") = DBNull.Value
                    .Item("InvAddress") = DBNull.Value
                    .Item("InvPurposeCode") = DBNull.Value
                    .Item("InvPurposeName") = DBNull.Value
                    .Item("InvoiceKind") = 0
                    .Item("Email") = DBNull.Value
                    .Item("DenRecCode") = DBNull.Value
                    .Item("DenRecName") = DBNull.Value
                    .Item("DenRecDate") = DBNull.Value
                    .Item("MailAddrNo") = .Item("InstAddrNo")
                    .Item("MailAddress") = .Item("InstAddress")
                    .Item("CustNote") = DBNull.Value
                    .Item("ChargeNote") = DBNull.Value
                End With
                Customer.Rows.Add(CustomerRow)
                '更新客戶資料SO001
                If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO,
                    UpdateMode.UpdateRow, Customer, "SO001", "CustId = " & PrDataRow("CustId") &
                        " And ServiceType = '" & PrDataRow("ServiceType") & "'") Then
                    Throw New Exception("Update Customer 001")
                End If
                Customer.Columns.Remove("CustNote")
                Customer.Columns.Remove("ChargeNote")
                Customer.Columns.Remove("MailAddrNo")
                Customer.Columns.Remove("MailAddress")
                Customer.Columns.Remove("InstAddrNo")
                Customer.Columns.Remove("InstAddress")
                '更新客戶資料SO002
                If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO,
                    UpdateMode.UpdateRow, Customer, "SO002", "CustId = " & PrDataRow("CustId") &
                        " And ServiceType = '" & PrDataRow("ServiceType") & "'") Then
                    Throw New Exception("Update Customer 002")
                End If
                '如果都沒有其他服務在用的帳號則將SO002A/SO106帳號帳號停用
                DAO.ExecNqry(_DAL.GetStopAccountNo("SO002A"), New Object() {PrDataRow("CustId"), PrDataRow("ServiceType"), DateTime.Now})
                DAO.ExecNqry(_DAL.GetStopAccountNo("SO106"), New Object() {PrDataRow("CustId"), PrDataRow("ServiceType"), DateTime.Now})
                Customer.Dispose()
            End If

            DefCMCode.Dispose()
            ChargeSys.Dispose()
            DefaultSys.Dispose()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "停復異動資料(SO003)"
    Private Function StopChangeData(ByVal PRDataRow As DataRow) As Boolean
        Try
            '參考號為7 完工時將該服務的週期性收費資料停用
            'If PRDataRow.IsNull(PRDataRow.Table.Columns("Fintime"), DataRowVersion.Original) AndAlso
            '   Not PRDataRow.IsNull(PRDataRow.Table.Columns("Fintime"), DataRowVersion.Current) Then
            If CableSoft.BLL.Utility.Utility.CheckNullToNotNull(PRDataRow, "FinTime") Then
                DAO.ExecNqry(_DAL.GetStopPeriodCycle, New Object() {PRDataRow.Item("CustId"), PRDataRow.Item("CompCode"), PRDataRow.Item("ServiceType")})
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "更新客戶促銷資料檔"
    Private Function ChangeCustomerPromData(ByVal Wip As DataTable) As Boolean
        Try
            '如工單有促銷方案及消息來源, 則做以下動作:
            '1.先取得客戶促銷明細檔
            '2.如不存在則將資料新增到客戶促銷明細檔,回填欄位如下:
            '2.1.ServiceType=服務別,CompCode=公司別,CustId=客編,BulletinCode/BulletinName=消息來源,
            '2.2.MediaCode/MediaName=介紹媒介, PromCode/PromName=促銷方案, ProcDate=受理時間
            Dim WipRow As DataRow = Wip.Rows(0)
            If Not WipRow.IsNull("PromCode") AndAlso Not WipRow.IsNull("BulletinCode") Then
                Using CustomerProm As DataTable = DAO.ExecQry(_DAL.GetCustomerPromData(), New Object() {WipRow.Item("CustId"), WipRow.Item("ServiceType"), WipRow.Item("PromCode"), WipRow.Item("BulletinCode")})
                    If CustomerProm.Rows.Count = 0 Then
                        Dim CustomerPromRow As DataRow = CustomerProm.NewRow
                        CustomerProm.Rows.Add(CustomerPromRow)
                        With CustomerPromRow
                            .Item("ServiceType") = WipRow.Item("ServiceType")
                            .Item("CompCode") = WipRow.Item("CompCode")
                            .Item("CustId") = WipRow.Item("CustId")
                            .Item("BulletinCode") = WipRow.Item("BulletinCode")
                            .Item("BulletinName") = WipRow.Item("BulletinName")
                            .Item("MediaCode") = WipRow.Item("MediaCode")
                            .Item("MediaName") = WipRow.Item("MediaName")
                            .Item("PromCode") = WipRow.Item("PromCode")
                            .Item("PromName") = WipRow.Item("PromName")
                            .Item("ProcDate") = WipRow.Item("AcceptTime")
                        End With
                        If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, UpdateMode.InsertRow, CustomerProm, "SO098", "", , , False) Then
                            Throw New Exception("Update ChangeCustomerPromData")
                        End If
                    End If
                End Using
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "判斷設備是否最後一台計價，需要將SO002.PRTime填值"
    Private Function ChkFaciToUpd002(ByVal WipData As DataSet) As Boolean
        Try
            Dim FaciSEQNo As String = String.Empty
            Dim drWip As DataRow = WipData.Tables("Wip").Rows(0)
            For Each DrTmp As DataRow In WipData.Tables("PRFacility").Rows
                FaciSEQNo = String.Format("{0},{1}", FaciSEQNo, DrTmp("SEQNO"))
            Next
            If Not String.IsNullOrEmpty(FaciSEQNo) Then
                FaciSEQNo = FaciSEQNo.Substring(1)
                Dim CountFaci As Int32 = Int32.Parse("0" & DAO.ExecQry(_DAL.FaciCount(FaciSEQNo, "2,3,5,6,7,8,10"), New Object() {drWip("CustID"), drWip("ServiceType")}).Rows(0)(0).ToString)
                If CountFaci = 0 Then
                    DAO.ExecNqry(_DAL.FaciToUpSO002, New Object() {drWip("CustID"), drWip("ServiceType"), drWip("FinTime")})
                End If
            End If
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "新增工單如果有指定設備，須抓取設備PinCode來回填到工單內"
    Private Function PRChangeFaciPinCodeToWip(ByVal EditMode As EditMode, ByRef WipRow As DataRow, ByVal WipData As DataSet) As Boolean
        Try
            '新增工單才需要判斷指定設備的PinCode帶入工單內，否則回傳成功。
            If EditMode <> CableSoft.BLL.Utility.EditMode.Append Then Return True
            Dim dtChangeFaci As DataTable = WipData.Tables("ChangeFacility").Copy
            Dim strSEQNO As String = String.Empty

            Using objPR As New PR(LoginInfo, DAO)
                strSEQNO = objPR.GetChangeFacilitySEQNO(dtChangeFaci)
                '沒有指定設備則直接回傳成功
                If String.IsNullOrEmpty(strSEQNO) Then Return True
                Using dtPincode As DataTable = objPR.GetChangeFacilityPinCode(WipRow("Custid"), strSEQNO)
                    If dtPincode.Rows.Count > 0 Then
                        If Not WipRow.IsNull("PinCode") Then
                            If WipRow("PinCode") <> dtPincode.Rows(0)("PinCode") Then
                                WipRow("PinCode") = dtPincode.Rows(0)("PinCode")
                            End If
                        Else
                            WipRow("PinCode") = dtPincode.Rows(0)("PinCode")
                        End If
                    End If
                End Using
            End Using
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "將工單的SO009.PinCode回填到SO004.PinCode"
    Private Function UpdPinCode(EditMode As EditMode, ByVal WipRow As DataRow, ByVal dtChangeFacility As DataTable) As Boolean
        Try
            '#6721 2014.05.02 by Corey 
            '問題需求判斷完工時，要將工單的SO009.PinCode填寫SO004.PinCode。討論後判斷SO004.PRSNO是這張工單的資料才需要填寫到SO004內。
            If EditMode = CableSoft.BLL.Utility.EditMode.Edit AndAlso CableSoft.BLL.Utility.Utility.CheckNullToNotNull(WipRow, "FinTime") Then
                Dim strFaciSeqno As String = String.Empty
                strFaciSeqno = CableSoft.BLL.Utility.Utility.GetRowFieldString(dtChangeFacility, "SEQNO", "'")
                Dim strWhere As String = String.Format(" PRSNO='{0}' and Custid={1} and SEQNO in ({2})", WipRow("SNO"), WipRow("CustId"), strFaciSeqno)
                Dim strPinCode As String = "Null"
                If Not WipRow.IsNull("PinCode") Then strPinCode = String.Format("'{0}'", WipRow("PinCode"))
                DAO.ExecNqry(String.Format("Update SO004 Set PinCode={0} Where {1}", strPinCode, strWhere))
                '#6819 2014.07.08 by Corey 因為使用CableSoft.BLL.Utility.Utility.ExecuteCommand不能大量Updata大量資料，所以改用直接下語法的方式。
                'Using dtSO004 As DataTable = DAO.ExecQry("Select * From SO004 Where " & strWhere)
                '    For Each drSO004 As DataRow In dtSO004.Rows
                '        If Not WipRow.IsNull("PinCode") Then
                '            drSO004("PinCode") = WipRow("PinCode")
                '        Else
                '            drSO004("PinCode") = Nothing
                '        End If
                '    Next
                '    dtSO004.AcceptChanges()
                '    If dtSO004.Rows.Count > 0 Then
                '        If Not CableSoft.BLL.Utility.Utility.ExecuteCommand(DAO, CableSoft.Utility.DataAccess.UpdateMode.UpdateRow, dtSO004, "SO004", strWhere, True, , False) Then
                '            Return False
                '        End If
                '    End If
                'End Using
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "移機順產生其他服務移機單"
    Private Function OtherServiceMovePR(EditMode As EditMode,
                                        WipData As DataSet, MoveFaciData As DataSet) As RIAResult
        Try
            If MoveFaciData.Tables.Contains("Wip") Then
                Using bll As New CableSoft.SO.BLL.Wip.Utility.Utility(LoginInfo, DAO)
                    Dim MainWip As DataTable = WipData.Tables("Wip")
                    Dim MainWorkCode As DataTable = SOUtil.GetCode(SO.BLL.Utility.CodeType.PRCode, "InterDependRefNo", "CodeNo = " & MainWip.Rows(0).Item("PRCode"))
                    Dim InterDependRefNo As String = Nothing
                    If MainWorkCode.Rows(0).IsNull("InterDependRefNo") = False Then
                        InterDependRefNo = MainWorkCode.Rows(0).Item("InterDependRefNo")
                    End If
                    For Each TempWipRow As DataRow In MoveFaciData.Tables("Wip").Rows
                        '取得工單資料結構
                        Using ServiceWipData As DataSet = bll.GetWipCalculateData(SO.BLL.Utility.InvoiceType.PR, MainWip.Rows(0).Item("CustId"), TempWipRow.Item("ServiceType"), MainWip.Rows(0).Item("ResvTime"), TempWipRow.Item("PRCode"))
                            Dim newWipRow As DataRow = CableSoft.BLL.Utility.Utility.CopyDataRow(MainWip.Rows(0), ServiceWipData.Tables("Wip").NewRow())
                            Using WorkCode As DataTable = SOUtil.GetCode(SO.BLL.Utility.CodeType.PRCode, "Description,WorkUnit,InterDependRefNo", "CodeNo = " & TempWipRow.Item("PRCode"))
                                '回填工單要改的欄位
                                With newWipRow
                                    .Item("ServiceType") = TempWipRow.Item("ServiceType")
                                    .Item("PRCode") = TempWipRow.Item("PRCode")
                                    .Item("PRName") = TempWipRow.Item("PRName")
                                    .Item("ReasonCode") = TempWipRow.Item("ReasonCode")
                                    .Item("ReasonName") = TempWipRow.Item("ReasonName")
                                    .Item("MainSNo") = MainWip.Rows(0).Item("SNo")
                                    .Item("WorkUnit") = WorkCode.Rows(0).Item("WorkUnit")
                                    .Item("SNo") = SOUtil.GetFalseSNo(SO.BLL.Utility.InvoiceType.PR, .Item("ServiceType"))
                                End With
                                ServiceWipData.Tables("Wip").Rows.Add(newWipRow)
                                '取得SO004D的相關資料
                                Using FaciData As DataTable = DAO.ExecQry(_DAL.GetMoveFaciData(InterDependRefNo), New Object() {newWipRow.Item("CustId"), newWipRow.Item("ServiceType")})
                                    Using bllChangeFaci As New CableSoft.SO.BLL.Facility.ChangeFaci.ChangeFaci(LoginInfo, DAO)
                                        For Each FaciRow As DataRow In FaciData.Rows
                                            Dim ServiceIds As String = bllChangeFaci.GetChooseServiceIDs(FaciRow.Item("CustId"), FaciRow.Item("SeqNo"))
                                            Dim Delete003Citems As String = bllChangeFaci.GetDelete003Citem(ServiceIds)
                                            Using TempChangeFaci As DataTable = ServiceWipData.Tables("ChangeFacility").Clone()
                                                If Not bll.GetChangeFacility(Utility.FaciChangeType.FaciMove, newWipRow, FaciRow, FaciRow, Delete003Citems, ServiceIds, TempChangeFaci) Then
                                                    Throw New Exception("OtherServiceMovePR_GetChangeFacility")
                                                End If
                                                CableSoft.BLL.Utility.Utility.CopyDataTable(TempChangeFaci, ServiceWipData.Tables("ChangeFacility"), False)
                                            End Using
                                        Next
                                    End Using
                                End Using
                            End Using

                            Using PR As New PR(LoginInfo, DAO)
                                Dim ChangeFaciSeqNo As String = PR.GetChangeFacilitySEQNO(ServiceWipData.Tables("ChangeFacility"))
                                If Not String.IsNullOrEmpty(ChangeFaciSeqNo) Then
                                    Using dtPinCode As DataTable = PR.GetChangeFacilityPinCode(MainWip.Rows(0).Item("CustId"), ChangeFaciSeqNo)
                                        If dtPinCode.Rows.Count > 0 Then
                                            If Not dtPinCode.Rows(0).IsNull("PinCode") Then
                                                ServiceWipData.Tables("Wip").Rows(0)("PinCode") = dtPinCode.Rows(0)("PinCode")
                                            Else
                                                ServiceWipData.Tables("Wip").Rows(0)("PinCode") = Nothing
                                            End If
                                        Else
                                            ServiceWipData.Tables("Wip").Rows(0)("PinCode") = Nothing
                                        End If
                                    End Using
                                End If
                            End Using

                            '呼叫工單存檔
                            Dim result As RIAResult = Save(EditMode, False, ServiceWipData, False)
                            If result.ResultBoolean = False Then
                                Return result
                            End If
                        End Using
                    Next
                    MainWorkCode.Dispose()
                End Using
            End If
            Return New RIAResult With {.ResultBoolean = True}
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

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
